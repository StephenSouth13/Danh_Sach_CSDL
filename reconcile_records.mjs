import fs from "node:fs";
import path from "node:path";

const [targetDir, refDir, outDir] = process.argv.slice(2);
if (!targetDir || !refDir || !outDir) {
  throw new Error("Usage: node reconcile_records.mjs <target-xlsx-dir> <ref-xlsx-dir> <out-dir>");
}

const read = (p) => fs.readFileSync(p, "utf8");
const write = (p, s) => fs.writeFileSync(p, s, "utf8");

function decodeXml(s = "") {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function escapeXml(s = "") {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function colIndex(ref) {
  const letters = String(ref).match(/^[A-Z]+/)?.[0] || "";
  let n = 0;
  for (const ch of letters) n = n * 26 + ch.charCodeAt(0) - 64;
  return n;
}

function sheetMap(dir) {
  const wb = read(path.join(dir, "xl", "workbook.xml"));
  const rels = read(path.join(dir, "xl", "_rels", "workbook.xml.rels"));
  const relMap = new Map();
  for (const m of rels.matchAll(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)) {
    relMap.set(m[1], m[2]);
  }
  const map = new Map();
  for (const m of wb.matchAll(/<sheet\b[^>]*name="([^"]+)"[^>]*sheetId="([^"]+)"[^>]*(?:r:id|id)="([^"]+)"/g)) {
    let target = relMap.get(m[3]);
    if (!target) continue;
    if (!target.startsWith("xl/")) target = `xl/${target}`;
    map.set(decodeXml(m[1]), { name: decodeXml(m[1]), id: m[2], relId: m[3], target });
  }
  return map;
}

function sharedStrings(dir) {
  const p = path.join(dir, "xl", "sharedStrings.xml");
  if (!fs.existsSync(p)) return [];
  const xml = read(p);
  const result = [];
  for (const si of xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)) {
    let text = "";
    for (const t of si[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)) text += decodeXml(t[1]);
    result.push(text);
  }
  return result;
}

function cellValue(cellXml, t, shared) {
  if (t === "s") {
    const v = cellXml.match(/<v>([\s\S]*?)<\/v>/)?.[1];
    return shared[Number(v)] ?? "";
  }
  if (t === "inlineStr") {
    let text = "";
    for (const m of cellXml.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)) text += decodeXml(m[1]);
    return text;
  }
  const v = cellXml.match(/<v>([\s\S]*?)<\/v>/)?.[1];
  return v ? decodeXml(v) : "";
}

function rowsFromSheet(dir, sheet, shared) {
  const xml = read(path.join(dir, sheet.target));
  const rows = [];
  for (const row of xml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)) {
    const values = new Map();
    for (const cell of row[2].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
      const ref = cell[1].match(/\br="([^"]+)"/)?.[1];
      if (!ref) continue;
      const t = cell[1].match(/\bt="([^"]+)"/)?.[1] || "";
      values.set(colIndex(ref), cellValue(cell[2], t, shared));
    }
    rows.push({ rowNumber: Number(row[1]), values });
  }
  return rows;
}

function normalize(text = "") {
  return String(text)
    .toLowerCase()
    .replace(/^https?:\/\/(www\.)?/, "")
    .replace(/['"`]/g, " ")
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .replace(/đ/g, "d")
    .replace(/\b(ky|ki)\s*luc\s*gia\b/g, " ")
    .replace(/\b(klg|gs|ts|pgs|ths|nsut|nsnd|ong|ba|anh|chi|thay|hoa si|nha bao|nghe si|giao su|vien si)\b/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function tokens(text) {
  const n = normalize(text);
  return n ? [...new Set(n.split(/\s+/).filter((t) => t.length > 1))] : [];
}

function tokenScore(a, b) {
  if (!a.length || !b.length) return 0;
  const bs = new Set(b);
  let common = 0;
  for (const t of a) if (bs.has(t)) common++;
  const ca = common / a.length;
  const cb = common / b.length;
  return Math.max(ca, cb) * 0.65 + Math.min(ca, cb) * 0.35;
}

function cleanUrl(url = "") {
  return String(url).trim().toLowerCase().replace(/^http:\/\//, "https://").replace(/\/+$/, "");
}

function refRecord(sheet, row, owner, title, date, link) {
  return {
    sheet,
    row,
    owner: owner || "",
    title: title || "",
    date: date || "",
    link: link || "",
    cleanLink: cleanUrl(link),
    ownerTokens: tokens(owner),
    titleTokens: tokens(title),
  };
}

function loadRefs() {
  const shared = sharedStrings(refDir);
  const sheets = sheetMap(refDir);
  const refs = [];
  for (const name of ["CSDL KLVN", "CSDL KLVW", "CSDL KL CA", "CSDL KLTG"]) {
    const sheet = sheets.get(name);
    if (!sheet) throw new Error(`Missing sheet: ${name}`);
    const rows = rowsFromSheet(refDir, sheet, shared);
    for (const r of rows) {
      if (r.rowNumber < 3) continue;
      if (name === "CSDL KLVN") refs.push(refRecord(name, r.rowNumber, r.values.get(6), r.values.get(5), r.values.get(4), ""));
      else if (name === "CSDL KLVW") refs.push(refRecord(name, r.rowNumber, r.values.get(7), r.values.get(6), r.values.get(5), ""));
      else refs.push(refRecord(name, r.rowNumber, r.values.get(2), r.values.get(3), r.values.get(5), r.values.get(6)));
    }
  }
  return refs.filter((r) => `${r.owner}${r.title}${r.link}`.trim());
}

function buildIndex(refs) {
  const url = new Map();
  const token = new Map();
  for (const r of refs) {
    if (r.cleanLink) url.set(r.cleanLink, r);
    for (const t of [...r.ownerTokens, ...r.titleTokens]) {
      if (t.length < 4) continue;
      if (!token.has(t)) token.set(t, []);
      token.get(t).push(r);
    }
  }
  return { url, token };
}

function findBest(record, refs, index) {
  const url = cleanUrl(record.url);
  const nameTokens = tokens(record.name);
  const titleTokens = tokens(record.title);
  if (url && index.url.has(url)) {
    const ref = index.url.get(url);
    return { ref, titleScore: 1, nameScore: tokenScore(nameTokens, ref.ownerTokens), urlExact: true };
  }

  const candidates = new Map();
  for (const t of [...nameTokens, ...titleTokens]) {
    if (t.length < 4) continue;
    for (const r of index.token.get(t) || []) candidates.set(`${r.sheet}:${r.row}`, r);
  }
  const list = candidates.size ? [...candidates.values()] : refs;
  let best = null;
  let bestScore = -1;
  for (const ref of list) {
    const titleScore = tokenScore(titleTokens, ref.titleTokens);
    const nameScore = tokenScore(nameTokens, ref.ownerTokens);
    const score = titleScore * 0.72 + nameScore * 0.28;
    if (score > bestScore) {
      bestScore = score;
      best = { ref, titleScore, nameScore, urlExact: false };
    }
  }
  return best;
}

function noteFor(record, match) {
  if (!match?.ref) return "THIẾU: không tìm thấy trong 4 tab CSDL.";
  const r = match.ref;
  const strong = match.urlExact || match.titleScore >= 0.82 || (match.nameScore >= 0.78 && match.titleScore >= 0.55);
  if (!strong) {
    return `THIẾU/CHƯA CHẮC: chưa tìm thấy dòng khớp chắc trong 4 tab CSDL. Gần nhất: ${r.sheet}!R${r.row}; tên CSDL: ${r.owner}; kỷ lục CSDL: ${r.title}; điểm tên ${Math.round(match.nameScore * 100)}%, điểm kỷ lục ${Math.round(match.titleScore * 100)}%.`;
  }
  const issues = [];
  if (!match.urlExact && record.url && r.link) issues.push(`link khác: Data=${record.url}; CSDL=${r.link}`);
  else if (record.url && !r.link) issues.push("CSDL chưa có link để đối chiếu");
  else if (!record.url && r.link) issues.push(`thiếu link trong Data; CSDL có: ${r.link}`);
  if (match.nameScore < 0.78) issues.push(`tên cá nhân/đơn vị khác: Data='${record.name}'; CSDL='${r.owner}'`);
  if (match.titleScore < 0.82) issues.push(`tên kỷ lục khác/thiếu: CSDL='${r.title}'`);
  if (!issues.length) return `ĐÃ KHỚP với ${r.sheet}!R${r.row}.`;
  return `CẦN RÀ SOÁT với ${r.sheet}!R${r.row}: ${issues.join("; ")}`;
}

function setCell(rowXml, rowNumber, cellRef, text) {
  const cell = `<c r="${cellRef}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(text)}</t></is></c>`;
  const hRe = new RegExp(`<c\\b[^>]*\\br="${cellRef}"[^>]*(?:\\/>|>[\\s\\S]*?<\\/c>)`);
  if (hRe.test(rowXml)) return rowXml.replace(hRe, cell);

  const cells = [...rowXml.matchAll(/<c\b[^>]*\br="([A-Z]+\d+)"[^>]*>[\s\S]*?<\/c>/g)];
  for (const c of cells) {
    if (colIndex(c[1]) > colIndex(cellRef)) {
      return rowXml.slice(0, c.index) + cell + rowXml.slice(c.index);
    }
  }
  return rowXml.replace("</row>", `${cell}</row>`);
}

const refs = loadRefs();
const index = buildIndex(refs);
const targetSheets = sheetMap(targetDir);
const targetShared = sharedStrings(targetDir);
const dataSheet = [...targetSheets.values()].find((s) => s.name.startsWith("Data_record_01-01-2022_01-01-20"));
if (!dataSheet) throw new Error("Target Data_record sheet not found.");
const rows = rowsFromSheet(targetDir, dataSheet, targetShared);

const notes = new Map();
const summary = { total: 0, matched: 0, review: 0, missing: 0 };
for (const r of rows) {
  if (r.rowNumber <= 1) continue;
  const record = {
    row: r.rowNumber,
    recordId: r.values.get(1) || "",
    gabId: r.values.get(2) || "",
    name: r.values.get(3) || "",
    date: r.values.get(4) || "",
    title: r.values.get(5) || "",
    url: r.values.get(6) || "",
    description: r.values.get(7) || "",
  };
  if (!`${record.name}${record.title}${record.url}`.trim()) continue;
  const note = noteFor(record, findBest(record, refs, index));
  notes.set(r.rowNumber, note);
  summary.total++;
  if (note.startsWith("ĐÃ KHỚP")) summary.matched++;
  else if (note.startsWith("THIẾU")) summary.missing++;
  else summary.review++;
}

const sheetPath = path.join(targetDir, dataSheet.target);
let xml = read(sheetPath);
xml = xml.replace(/<dimension\b[^>]*ref="[^"]*"[^/]*\/>/, (m) => m.replace(/ref="[^"]*"/, `ref="A1:H${Math.max(...rows.map((r) => r.rowNumber))}"`));
xml = xml.replace(/<row\b[^>]*\br="1"[^>]*>[\s\S]*?<\/row>/, (m) => setCell(m, 1, "H1", "Note"));
for (const [rowNumber, note] of notes) {
  const re = new RegExp(`<row\\b[^>]*\\br="${rowNumber}"[^>]*>[\\s\\S]*?<\\/row>`);
  xml = xml.replace(re, (m) => setCell(m, rowNumber, `H${rowNumber}`, note));
}
write(sheetPath, xml);

fs.mkdirSync(outDir, { recursive: true });
fs.writeFileSync(path.join(outDir, "summary.json"), JSON.stringify(summary, null, 2), "utf8");
console.log(JSON.stringify({ outputDir: outDir, sheet: dataSheet.name, refs: refs.length, ...summary }, null, 2));
