import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const root = process.cwd();
const sourceXlsx = path.join(root, "Dulieutruyxuat", "Bản sao của DỮ LIỆU THÀNH TỰU GAB - 15_24, 13 tháng 9.xlsx");
const refXlsx = path.join(root, "Dulieutruyxuat", "MÔ TẢ YÊU CẦU NHẬP THÔNG TIN GAB.xlsx");
const outRoot = path.join(root, "outputs", "highlight_200_klg_duplicate");
const srcDir = path.join(outRoot, "src");
const refDir = path.join(outRoot, "ref");
const outputXlsx = path.join(root, "outputs", "200_KLG_CHUAN_HOA_duplicate_highlight.xlsx");

function ps(command) {
  execFileSync("powershell", ["-NoProfile", "-Command", command], { stdio: "pipe" });
}
function resetDir(dir) {
  if (fs.existsSync(dir)) fs.rmSync(dir, { recursive: true, force: true });
  fs.mkdirSync(dir, { recursive: true });
}
function read(p) {
  return fs.readFileSync(p, "utf8");
}
function write(p, s) {
  fs.writeFileSync(p, s, "utf8");
}
function esc(s = "") {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function dec(s = "") {
  return String(s)
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
}
function colIndex(ref) {
  let n = 0;
  for (const ch of (String(ref).match(/^[A-Z]+/)?.[0] || "")) n = n * 26 + ch.charCodeAt(0) - 64;
  return n;
}
function colLetters(n) {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
function sheetMap(dir) {
  const wb = read(path.join(dir, "xl", "workbook.xml"));
  const rels = read(path.join(dir, "xl", "_rels", "workbook.xml.rels"));
  const relMap = new Map();
  for (const m of rels.matchAll(/<Relationship\b([^>]*)\/>/g)) {
    const id = m[1].match(/\bId="([^"]+)"/)?.[1];
    const target = m[1].match(/\bTarget="([^"]+)"/)?.[1];
    if (id && target) relMap.set(id, target);
  }
  const map = new Map();
  for (const m of wb.matchAll(/<sheet\b([^>]*)\/>/g)) {
    const name = dec(m[1].match(/\bname="([^"]+)"/)?.[1] || "");
    const id = m[1].match(/\br:id="([^"]+)"/)?.[1] || m[1].match(/\bid="([^"]+)"/)?.[1];
    let target = relMap.get(id);
    if (!name || !target) continue;
    if (target.startsWith("/")) target = target.slice(1);
    if (!target.startsWith("xl/")) target = `xl/${target}`;
    map.set(name, { name, target });
  }
  return map;
}
function sharedStrings(dir) {
  const p = path.join(dir, "xl", "sharedStrings.xml");
  if (!fs.existsSync(p)) return [];
  const xml = read(p);
  return [...xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map((si) =>
    [...si[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((t) => dec(t[1])).join("")
  );
}
function cellValue(cellXml, attrs, shared) {
  const t = attrs.match(/\bt="([^"]+)"/)?.[1] || "";
  if (t === "s") return shared[Number(cellXml.match(/<v>([\s\S]*?)<\/v>/)?.[1])] ?? "";
  if (t === "inlineStr") return [...cellXml.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((m) => dec(m[1])).join("");
  return dec(cellXml.match(/<v>([\s\S]*?)<\/v>/)?.[1] || "");
}
function rowsFromSheet(dir, sheet, shared) {
  const xml = read(path.join(dir, sheet.target));
  const rows = [];
  for (const row of xml.matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>/g)) {
    const rowNumber = Number(row[1].match(/\br="(\d+)"/)?.[1]);
    if (!rowNumber) continue;
    const values = new Map();
    for (const cell of row[2].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>|<c\b([^>]*)\/>/g)) {
      const attrs = cell[1] || cell[3] || "";
      const ref = attrs.match(/\br="([^"]+)"/)?.[1];
      if (!ref) continue;
      values.set(colIndex(ref), cell[2] ? cellValue(cell[2], attrs, shared) : "");
    }
    rows.push({ rowNumber, values });
  }
  return rows;
}

function normalize(s = "") {
  return String(s).toLowerCase().normalize("NFD")
    .replace(/\p{Diacritic}/gu, "").replace(/đ/g, "d")
    .replace(/\b(ky|ki)\s*luc\s*gia\b/g, " ")
    .replace(/\b(klg|gs|pgs|ts|ths|bs|nsut|nsnd|ong|ba|anh|chi|hoa si|nha bao|nghe si|nghe nhan|giao su|vien si|tien si|thac si|danh du)\b/g, " ")
    .replace(/[^a-z0-9]+/g, " ").replace(/\s+/g, " ").trim();
}
const stop = new Set("xac lap ky luc viet nam the gioi chau a toan cau nguoi dau tien nhieu nhat lon nhat cao nhat co cua va voi trong tai nam ngay bo duoc bang bang nhan".split(" "));
function tokens(s = "", keepStop = false) {
  const n = normalize(s);
  return n ? [...new Set(n.split(" ").filter((x) => x.length > 1 && (keepStop || !stop.has(x))))] : [];
}
function score(a, b) {
  if (!a.length || !b.length) return 0;
  const as = new Set(a), bs = new Set(b);
  let common = 0;
  for (const t of as) if (bs.has(t)) common++;
  const ca = common / as.size, cb = common / bs.size;
  return Math.max(ca, cb) * 0.65 + Math.min(ca, cb) * 0.35;
}
function short(s = "", max = 160) {
  const t = String(s ?? "").replace(/\s+/g, " ").trim();
  return t.length > max ? `${t.slice(0, max - 3)}...` : t;
}
function excelDateText(value = "") {
  const s = String(value ?? "").trim();
  if (!s) return "";
  if (/^\d+(\.0+)?$/.test(s)) {
    const serial = Number(s);
    if (serial > 20000 && serial < 70000) {
      const d = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
      return `${String(d.getUTCDate()).padStart(2, "0")}/${String(d.getUTCMonth() + 1).padStart(2, "0")}/${d.getUTCFullYear()}`;
    }
  }
  return s.replace(/\s+/g, " ");
}
function parseAllDateParts(value = "") {
  if (/^\d+(\.0+)?$/.test(String(value ?? "").trim())) {
    const serial = Number(String(value).trim());
    if (serial <= 20000 || serial >= 70000) return [];
  }
  const s = excelDateText(value);
  const result = [];
  for (const m of s.matchAll(/\b(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})\b/g)) {
    const y = m[3].length === 2 ? `20${m[3]}` : m[3];
    result.push({ y, m: m[2].padStart(2, "0"), d: m[1].padStart(2, "0"), level: "day", text: m[0] });
  }
  if (result.length) return result;
  for (const m of s.matchAll(/\b(\d{1,2})[\/\-.](\d{4})\b/g)) result.push({ y: m[2], m: m[1].padStart(2, "0"), d: "", level: "month", text: m[0] });
  if (result.length) return result;
  for (const m of s.matchAll(/\b(19|20)\d{2}\b/g)) result.push({ y: m[0], m: "", d: "", level: "year", text: m[0] });
  return result;
}
function validDateText(value = "") {
  return parseAllDateParts(value).map((x) => x.text).join("; ");
}
function dateOk(sourceValues, refValue) {
  const refs = parseAllDateParts(refValue);
  const sources = sourceValues.flatMap(parseAllDateParts);
  if (!refs.length) return false;
  for (const src of sources) {
    for (const ref of refs) {
      if (src.y === ref.y && src.m && ref.m && src.m === ref.m && src.d && ref.d && src.d === ref.d) return true;
      if (src.y === ref.y && src.m && ref.m && src.m === ref.m && (src.level === "month" || ref.level === "month" || !src.d || !ref.d)) return true;
      if (src.y === ref.y && (src.level === "year" || ref.level === "year" || (!src.m && !ref.m))) return true;
    }
  }
  return false;
}
function typeFromText(s = "") {
  const n = normalize(s);
  if (n.includes("the gioi")) return "CSDL KLTG";
  if (n.includes("chau a")) return "CSDL KL CA";
  if (n.includes("toan cau")) return "CSDL KLVW";
  if (n.includes("viet nam")) return "CSDL KLVN";
  return "";
}
function refTypeFromSheet(sheet) {
  return sheet === "CSDL KLVN" ? "KỶ LỤC VIỆT NAM" : sheet === "CSDL KL CA" ? "KỶ LỤC CHÂU Á" : sheet === "CSDL KLTG" ? "KỶ LỤC THẾ GIỚI" : "KỶ LỤC NGƯỜI VIỆT TOÀN CẦU";
}
function refFromRow(sheetName, r) {
  let owner = "", title = "", date = "", id = "";
  if (sheetName === "CSDL KLVN") {
    id = r.values.get(3) || ""; date = r.values.get(4) || ""; title = r.values.get(5) || ""; owner = r.values.get(6) || "";
  } else if (sheetName === "CSDL KLVW") {
    id = r.values.get(4) || ""; date = r.values.get(5) || ""; title = r.values.get(6) || ""; owner = r.values.get(7) || "";
  } else {
    owner = r.values.get(2) || ""; title = r.values.get(3) || ""; id = r.values.get(4) || ""; date = r.values.get(5) || r.values.get(4) || "";
  }
  if (!String(owner + title).trim()) return null;
  return { sheet: sheetName, row: r.rowNumber, id: String(id).trim(), owner: String(owner).trim(), title: String(title).trim(), date: String(date).trim(), type: refTypeFromSheet(sheetName), ownerTokens: tokens(owner, true), titleTokens: tokens(title) };
}
function sourceFromRow(r) {
  const title = String(r.values.get(5) || "").trim();
  const gabTitle = String(r.values.get(14) || "").trim();
  const name = String(r.values.get(7) || "").trim();
  return {
    row: r.rowNumber,
    stt: String(r.values.get(1) || "").trim(),
    id: String(r.values.get(3) || "").trim(),
    dateOfficial: String(r.values.get(4) || "").trim(),
    title,
    name,
    type: String(r.values.get(11) || "").trim(),
    dateGab: String(r.values.get(13) || "").trim(),
    gabTitle,
    articleLink: String(r.values.get(15) || "").trim(),
    timeInContent: String(r.values.get(21) || "").trim(),
    nameTokens: tokens(name, true),
    titleTokens: tokens([title, gabTitle].filter(Boolean).join(" ")),
  };
}

function bestMatch(src, refs) {
  const wantedSheet = typeFromText(src.type || src.gabTitle || src.title);
  let best = null;
  for (const ref of refs) {
    const idHit = src.id && ref.id && normalize(src.id) === normalize(ref.id);
    const typeBonus = wantedSheet && ref.sheet === wantedSheet ? 0.12 : 0;
    const nameScore = score(src.nameTokens, ref.ownerTokens);
    const titleScore = score(src.titleTokens, ref.titleTokens);
    const total = (idHit ? 0.35 : 0) + nameScore * 0.25 + titleScore * 0.40 + typeBonus;
    if (!best || total > best.total) best = { ref, total, nameScore, titleScore, idHit };
  }
  return best;
}
function ensureInlineCell(rowXml, rowNumber, col, value, styleId = "") {
  const ref = `${colLetters(col)}${rowNumber}`;
  const style = styleId !== "" ? ` s="${styleId}"` : "";
  const cell = `<c r="${ref}"${style} t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`;
  const re = new RegExp(`<c\\b[^>]*\\br="${ref}"[^>]*(?:\\/>|>[\\s\\S]*?<\\/c>)`);
  if (re.test(rowXml)) return rowXml.replace(re, cell);
  return rowXml.replace("</row>", `${cell}</row>`);
}
function styleCell(rowXml, rowNumber, col, styleId) {
  const ref = `${colLetters(col)}${rowNumber}`;
  const selfClosing = new RegExp(`<c\\b([^>]*)\\br="${ref}"([^>]*)\\/>`);
  if (selfClosing.test(rowXml)) {
    return rowXml.replace(selfClosing, (_m, a, b) => {
      const attrs = `${a}r="${ref}"${b}`.replace(/\s+s="[^"]*"/, "").replace(/\s*\/\s*$/, "");
      return `<c${attrs} s="${styleId}"/>`;
    });
  }
  const paired = new RegExp(`<c\\b([^>]*)\\br="${ref}"([^>]*)>([\\s\\S]*?)<\\/c>`);
  if (paired.test(rowXml)) {
    return rowXml.replace(paired, (_m, a, b, body) => {
      const attrs = `${a}r="${ref}"${b}`.replace(/\s+s="[^"]*"/, "");
      return `<c${attrs} s="${styleId}">${body}</c>`;
    });
  }
  return rowXml.replace("</row>", `<c r="${ref}" s="${styleId}"/></row>`);
}
function addStyles(stylesPath) {
  let xml = read(stylesPath);
  const fillsCount = Number(xml.match(/<fills count="(\d+)"/)?.[1] || 0);
  const xfsCount = Number(xml.match(/<cellXfs count="(\d+)"/)?.[1] || 0);
  const fillXml = [
    `<fill><patternFill patternType="solid"><fgColor rgb="FFD9EAD3"/><bgColor indexed="64"/></patternFill></fill>`,
    `<fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>`,
    `<fill><patternFill patternType="solid"><fgColor rgb="FFF4CCCC"/><bgColor indexed="64"/></patternFill></fill>`,
    `<fill><patternFill patternType="solid"><fgColor rgb="FFCFE2F3"/><bgColor indexed="64"/></patternFill></fill>`,
  ].join("");
  xml = xml.replace(/<fills count="(\d+)">([\s\S]*?)<\/fills>/, `<fills count="${fillsCount + 4}">$2${fillXml}</fills>`);
  const newXfs = [0, 1, 2, 3].map((i) => `<xf numFmtId="0" fontId="0" fillId="${fillsCount + i}" borderId="0" xfId="0" applyFill="1"/>`).join("");
  xml = xml.replace(/<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/, `<cellXfs count="${xfsCount + 4}">$2${newXfs}</cellXfs>`);
  write(stylesPath, xml);
  return { green: xfsCount, yellow: xfsCount + 1, red: xfsCount + 2, blue: xfsCount + 3 };
}

resetDir(outRoot);
fs.mkdirSync(srcDir, { recursive: true });
fs.mkdirSync(refDir, { recursive: true });
fs.copyFileSync(sourceXlsx, path.join(outRoot, "source.zip"));
fs.copyFileSync(refXlsx, path.join(outRoot, "ref.zip"));
ps(`Expand-Archive -LiteralPath '${path.join(outRoot, "source.zip").replace(/'/g, "''")}' -DestinationPath '${srcDir.replace(/'/g, "''")}' -Force`);
ps(`Expand-Archive -LiteralPath '${path.join(outRoot, "ref.zip").replace(/'/g, "''")}' -DestinationPath '${refDir.replace(/'/g, "''")}' -Force`);

const srcSheets = sheetMap(srcDir);
const refSheets = sheetMap(refDir);
const sourceSheet = srcSheets.get("200 KLG CHUẨN HÓA");
if (!sourceSheet) throw new Error("Không thấy tab 200 KLG CHUẨN HÓA");
const sourceRows = rowsFromSheet(srcDir, sourceSheet, sharedStrings(srcDir)).filter((r) => r.rowNumber >= 3).map(sourceFromRow).filter((r) => `${r.stt}${r.name}${r.title}${r.gabTitle}`.trim());
const refs = [];
for (const name of ["CSDL KLVN", "CSDL KLVW", "CSDL KL CA", "CSDL KLTG"]) {
  for (const r of rowsFromSheet(refDir, refSheets.get(name), sharedStrings(refDir))) {
    if (r.rowNumber < 3) continue;
    const ref = refFromRow(name, r);
    if (ref) refs.push(ref);
  }
}
const checks = new Map();
for (const src of sourceRows) {
  const match = bestMatch(src, refs);
  const ref = match?.ref;
  if (!ref) continue;
  if (match.idHit || (match.nameScore >= 0.72 && match.titleScore >= 0.45) || (match.nameScore >= 0.55 && match.titleScore >= 0.68)) {
    const titleOk = match.titleScore >= 0.70 || match.idHit;
    const nameOk = match.nameScore >= 0.72 || match.idHit;
    const dOk = dateOk([src.dateOfficial, src.dateGab, src.timeInContent], ref.date);
    if (nameOk && titleOk && dOk) continue;
    const status = nameOk && titleOk ? "LỆCH THỜI GIAN" : (!nameOk ? "LỆCH TÊN/CHỦ THỂ" : "LỆCH TÊN KỶ LỤC");
    checks.set(src.row, {
      style: status === "LỆCH THỜI GIAN" ? "yellow" : "red",
      status,
      note: status === "LỆCH THỜI GIAN" ? `Thời gian CSDL: ${excelDateText(ref.date)}` : `So lại với CSDL: ${short(ref.title)}`,
      ref: `${ref.sheet}!${ref.row}`,
      refTime: excelDateText(ref.date),
      refTitle: ref.title,
    });
  } else {
    checks.set(src.row, {
      style: "green",
      status: "MỚI Ở TAB 200 / CSDL CHƯA THẤY CHẮC",
      note: "Nếu research 200 đúng thì bổ sung vào CSDL chuẩn.",
      ref: ref ? `${ref.sheet}!${ref.row} (gợi ý yếu)` : "",
      refTime: ref ? excelDateText(ref.date) : "",
      refTitle: ref?.title || "",
    });
  }
}

const styleIds = addStyles(path.join(srcDir, "xl", "styles.xml"));
const sheetPath = path.join(srcDir, sourceSheet.target);
let sheetXml = read(sheetPath);
sheetXml = sheetXml.replace(/<dimension\b[^>]*ref="[^"]*"[^/]*\/>/, (m) => m.replace(/ref="[^"]*"/, `ref="A1:AF1114"`));
const headers = [
  [28, "CHECK"],
  [29, "NOTE"],
  [30, "CSDL REF"],
  [31, "CSDL TIME"],
  [32, "CSDL TITLE"],
];
sheetXml = sheetXml.replace(/<row\b[^>]*\br="1"[^>]*>[\s\S]*?<\/row>/, (row) => {
  let out = row;
  for (const [col, value] of headers) out = ensureInlineCell(out, 1, col, value);
  return out;
});
for (const [rowNumber, check] of checks) {
  const styleId = styleIds[check.style];
  const re = new RegExp(`<row\\b[^>]*\\br="${rowNumber}"[^>]*>[\\s\\S]*?<\\/row>`);
  sheetXml = sheetXml.replace(re, (row) => {
    let out = row;
    for (let col = 1; col <= 32; col++) out = styleCell(out, rowNumber, col, styleId);
    out = ensureInlineCell(out, rowNumber, 28, check.status, styleId);
    out = ensureInlineCell(out, rowNumber, 29, check.note, styleId);
    out = ensureInlineCell(out, rowNumber, 30, check.ref, styleId);
    out = ensureInlineCell(out, rowNumber, 31, check.refTime, styleId);
    out = ensureInlineCell(out, rowNumber, 32, check.refTitle, styleId);
    return out;
  });
}
write(sheetPath, sheetXml);

const zipOut = path.join(outRoot, "highlight.zip");
if (fs.existsSync(zipOut)) fs.rmSync(zipOut, { force: true });
if (fs.existsSync(outputXlsx)) fs.rmSync(outputXlsx, { force: true });
ps(`Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::CreateFromDirectory('${srcDir.replace(/'/g, "''")}', '${zipOut.replace(/'/g, "''")}')`);
fs.copyFileSync(zipOut, outputXlsx);
console.log(JSON.stringify({ outputXlsx, highlightedRows: checks.size }, null, 2));
