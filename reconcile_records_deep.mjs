import fs from "node:fs";
import path from "node:path";

const [targetDir, refDir, outDir] = process.argv.slice(2);
if (!targetDir || !refDir || !outDir) {
  throw new Error("Usage: node reconcile_records_deep.mjs <target-xlsx-dir> <ref-xlsx-dir> <out-dir>");
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
    map.set(decodeXml(m[1]), { name: decodeXml(m[1]), id: m[2], target });
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

const stopWords = new Set([
  "viet", "nam", "nguoi", "dau", "tien", "nhieu", "nhat", "lon", "the", "gioi",
  "chau", "luc", "xac", "lap", "ky", "gia", "duoc", "co", "cua", "tai", "trong",
  "voi", "ve", "va", "cac", "cho", "mot", "bo", "suu", "tap", "danh", "hieu",
]);

function normalize(text = "") {
  return String(text)
    .toLowerCase()
    .replace(/^https?:\/\/(www\.)?/, "")
    .replace(/['"`]/g, " ")
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .replace(/đ/g, "d")
    .replace(/\b(ky|ki)\s*luc\s*gia\b/g, " ")
    .replace(/\b(klg|gs|ts|pgs|ths|nsut|nsnd|ong|ba|anh|chi|thay|hoa si|nha bao|nghe si|giao su|vien si|nghe nhan|nha thiet ke)\b/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function tokens(text, keepStop = false) {
  const n = normalize(text);
  if (!n) return [];
  return [...new Set(n.split(/\s+/).filter((t) => t.length > 1 && (keepStop || !stopWords.has(t))))];
}

function buildWeights(refs) {
  const docs = refs.map((r) => new Set([...r.ownerTokens, ...r.titleTokens]));
  const df = new Map();
  for (const doc of docs) for (const t of doc) df.set(t, (df.get(t) || 0) + 1);
  const total = Math.max(1, docs.length);
  return new Map([...df].map(([t, c]) => [t, Math.log(1 + total / c)]));
}

function weightedScore(a, b, weights) {
  if (!a.length || !b.length) return 0;
  const bs = new Set(b);
  let common = 0;
  let aw = 0;
  let bw = 0;
  for (const t of a) {
    const w = weights.get(t) || 1;
    aw += w;
    if (bs.has(t)) common += w;
  }
  for (const t of b) bw += weights.get(t) || 1;
  if (!aw || !bw) return 0;
  const ca = common / aw;
  const cb = common / bw;
  return Math.max(ca, cb) * 0.65 + Math.min(ca, cb) * 0.35;
}

function cleanUrl(url = "") {
  return String(url).trim().toLowerCase().replace(/^http:\/\//, "https://").replace(/\/+$/, "");
}

function shortText(s = "", max = 180) {
  s = String(s).replace(/\s+/g, " ").trim();
  return s.length > max ? `${s.slice(0, max - 3)}...` : s;
}

function formatValue(value = "", kind = "text") {
  const s = String(value ?? "").replace(/\s+/g, " ").trim();
  if (kind !== "date" || !/^\d+(\.0+)?$/.test(s)) return s;
  const serial = Number(s);
  if (!Number.isFinite(serial) || serial < 20000 || serial > 70000) return s;
  const date = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
  const dd = String(date.getUTCDate()).padStart(2, "0");
  const mm = String(date.getUTCMonth() + 1).padStart(2, "0");
  const yyyy = date.getUTCFullYear();
  return `${dd}/${mm}/${yyyy}`;
}

function quoteSheet(name) {
  return `'${String(name).replace(/'/g, "''")}'`;
}

function refColumns(sheet) {
  if (sheet === "CSDL KLVN") {
    return {
      owner: "F",
      ownerName: "KỶ LỤC GIA",
      title: "E",
      titleName: "TÊN KỶ LỤC",
      date: "D",
      dateName: "NGÀY XÁC LẬP",
      link: "",
      linkName: "",
      id: "C",
      idName: "SỐ XLKL",
    };
  }
  if (sheet === "CSDL KLVW") {
    return {
      owner: "G",
      ownerName: "Cá nhân/ Đơn vị sở hữu Kỷ lục",
      title: "F",
      titleName: "Tên kỷ lục",
      date: "E",
      dateName: "NGÀY CẤP",
      link: "",
      linkName: "",
      id: "D",
      idName: "Số Xác lập – Thời điểm xác lập",
    };
  }
  return {
    owner: "B",
    ownerName: "TÊN CÁ NHÂN/ĐƠN VỊ",
    title: "C",
    titleName: "TÊN KỶ LỤC",
    date: "E",
    dateName: "THỜI ĐIỂM XÁC LẬP",
    link: "F",
    linkName: "LINK",
    id: "D",
    idName: "NĂM XÁC LẬP",
  };
}

function dataAudit(record) {
  const sheet = quoteSheet(record.sheet);
  return [
    `${sheet}!C${record.row}(fullName)="${shortText(record.name, 90)}"`,
    `${sheet}!D${record.row}(time)="${shortText(formatValue(record.date, "date"), 30)}"`,
    `${sheet}!E${record.row}(title)="${shortText(record.title, 170)}"`,
    `${sheet}!F${record.row}(url)="${shortText(record.url, 170)}"`,
  ].join("; ");
}

function refAudit(ref) {
  const c = refColumns(ref.sheet);
  const sheet = quoteSheet(ref.sheet);
  const parts = [
    `${sheet}!${c.owner}${ref.row}(${c.ownerName})="${shortText(ref.owner, 100)}"`,
    `${sheet}!${c.title}${ref.row}(${c.titleName})="${shortText(ref.title, 170)}"`,
    `${sheet}!${c.date}${ref.row}(${c.dateName})="${shortText(formatValue(ref.date, "date"), 40)}"`,
  ];
  if (c.id) parts.push(`${sheet}!${c.id}${ref.row}(${c.idName})="${shortText(ref.id || "", 40)}"`);
  if (c.link) parts.push(`${sheet}!${c.link}${ref.row}(${c.linkName})="${shortText(ref.link, 170)}"`);
  else parts.push(`${quoteSheet(ref.sheet)} không có cột LINK trong vùng đối chiếu chính`);
  return parts.join("; ");
}

function commonAndMissing(aTokens, bTokens) {
  const bs = new Set(bTokens);
  const as = new Set(aTokens);
  return {
    common: aTokens.filter((t) => bs.has(t)).slice(0, 12).join(", "),
    missingInCsdl: aTokens.filter((t) => !bs.has(t)).slice(0, 12).join(", "),
    extraInCsdl: bTokens.filter((t) => !as.has(t)).slice(0, 12).join(", "),
  };
}

function detailedNote(status, record, ref, reason, fix, extra = "") {
  const lines = [
    `Kết luận: ${status}`,
    `Dòng Data cần kiểm tra: ${dataAudit(record)}`,
  ];
  if (ref) lines.push(`Dòng CSDL dùng để đối chiếu: ${refAudit(ref)}`);
  lines.push(`Nhận định: ${reason}`);
  if (extra) lines.push(`Căn cứ đối chiếu: ${extra}`);
  lines.push(`Đề xuất xử lý: ${fix}`);
  return lines.join("\n");
}

function refRecord(sheet, row, owner, title, date, link) {
  return {
    sheet,
    row,
    owner: owner || "",
    title: title || "",
    date: date || "",
    link: link || "",
    id: "",
    cleanLink: cleanUrl(link),
    ownerTokens: tokens(owner, true),
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
      let ref;
      if (name === "CSDL KLVN") {
        ref = refRecord(name, r.rowNumber, r.values.get(6), r.values.get(5), r.values.get(4), "");
        ref.id = r.values.get(3) || "";
      } else if (name === "CSDL KLVW") {
        ref = refRecord(name, r.rowNumber, r.values.get(7), r.values.get(6), r.values.get(5), "");
        ref.id = r.values.get(4) || "";
      } else {
        ref = refRecord(name, r.rowNumber, r.values.get(2), r.values.get(3), r.values.get(5), r.values.get(6));
        ref.id = r.values.get(4) || "";
      }
      refs.push(ref);
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
      if (t.length < 3) continue;
      if (!token.has(t)) token.set(t, []);
      token.get(t).push(r);
    }
  }
  return { url, token };
}

function ranked(record, refs, index, weights) {
  const nameTokens = tokens(record.name, true);
  const titleTokens = tokens(record.title);
  const candidateMap = new Map();
  for (const t of [...nameTokens, ...titleTokens]) {
    if (t.length < 3) continue;
    for (const r of index.token.get(t) || []) candidateMap.set(`${r.sheet}:${r.row}`, r);
  }
  const list = candidateMap.size ? [...candidateMap.values()] : refs;
  return list
    .map((ref) => {
      const titleScore = weightedScore(titleTokens, ref.titleTokens, weights);
      const nameScore = weightedScore(nameTokens, ref.ownerTokens, weights);
      return { ref, titleScore, nameScore, score: titleScore * 0.7 + nameScore * 0.3 };
    })
    .sort((a, b) => b.score - a.score);
}

function analyze(record, refs, index, weights) {
  const url = cleanUrl(record.url);
  if (url && index.url.has(url)) {
    const ref = index.url.get(url);
    return {
      kind: "url",
      ref,
      titleScore: 1,
      nameScore: weightedScore(tokens(record.name, true), ref.ownerTokens, weights),
      ranked: [],
    };
  }

  const allRanked = ranked(record, refs, index, weights);
  const best = allRanked[0];
  const sameOwner = allRanked.filter((m) => m.nameScore >= 0.82);
  const bestSameOwner = sameOwner[0];
  const bestTitle = allRanked.find((m) => m.titleScore >= 0.82);

  return {
    kind: "fuzzy",
    ref: best?.ref,
    titleScore: best?.titleScore || 0,
    nameScore: best?.nameScore || 0,
    best,
    bestSameOwner,
    bestTitle,
    sameOwner: sameOwner.slice(0, 3),
  };
}

function noteFor(record, result) {
  if (!result?.ref) {
    return detailedNote(
      "THIẾU/CHƯA CÓ DÒNG CSDL ĐỐI CHIẾU",
      record,
      null,
      "Không tìm thấy dòng nào trong 4 tab CSDL có đủ dấu hiệu giống tên cá nhân/đơn vị hoặc tên kỷ lục.",
      `Kiểm tra lại Data ${quoteSheet(record.sheet)}!C${record.row}, E${record.row}, F${record.row}; nếu đúng là thành tựu hợp lệ thì bổ sung bản ghi tương ứng vào một trong 4 tab CSDL, nếu sai thì sửa lại Data ở các cột C/E/F.`
    );
  }

  const exactUrl = result.kind === "url";
  const ref = exactUrl ? result.ref : (result.bestSameOwner?.ref || result.best?.ref || result.ref);
  const titleScore = exactUrl ? result.titleScore : (result.bestSameOwner?.titleScore ?? result.titleScore);
  const nameScore = exactUrl ? result.nameScore : (result.bestSameOwner?.nameScore ?? result.nameScore);
  const titlePct = Math.round(titleScore * 100);
  const namePct = Math.round(nameScore * 100);
  const refId = `${quoteSheet(ref.sheet)}!R${ref.row}`;
  const nameDiff = commonAndMissing(tokens(record.name, true), ref.ownerTokens);
  const titleDiff = commonAndMissing(tokens(record.title), ref.titleTokens);
  const scoreLine = `Mức độ khớp tên: ${namePct}%; mức độ khớp nội dung kỷ lục: ${titlePct}%. Phần tên giống nhau: ${nameDiff.common || "không có rõ"}. Phần tên Data có nhưng CSDL chưa thể hiện: ${nameDiff.missingInCsdl || "không có"}. Phần nội dung kỷ lục giống nhau: ${titleDiff.common || "không có rõ"}. Phần Data có nhưng CSDL chưa thể hiện rõ: ${titleDiff.missingInCsdl || "không có"}. Phần CSDL có thêm/khác so với Data: ${titleDiff.extraInCsdl || "không có"}.`;

  if (exactUrl) {
    const issues = [];
    if (nameScore < 0.78) issues.push(`tên ở Data khác tên chủ thể trong CSDL`);
    if (titleScore < 0.82) issues.push(`tên kỷ lục ở Data chưa khớp tên kỷ lục trong CSDL`);
    if (!issues.length) {
      return detailedNote(
        "ĐÃ KHỚP",
        record,
        ref,
        `Trùng link với ${refId}, tên và tên kỷ lục đủ khớp.`,
        "Không cần sửa. Có thể giữ nguyên Data; nếu muốn đầy đủ hơn thì chỉ bổ sung link cho những tab CSDL chưa có cột/ô link.",
        scoreLine
      );
    }
    return detailedNote(
      "CẦN RÀ SOÁT",
      record,
      ref,
      `Trùng link với ${refId} nhưng ${issues.join("; ")}.`,
      `Ưu tiên đối chiếu bằng link ở Data ${quoteSheet(record.sheet)}!F${record.row}. Nếu link đúng bài gốc thì sửa tên/chủ thể hoặc tên kỷ lục tại Data C/E hoặc tại ô CSDL tương ứng; nếu link sai thì sửa Data F${record.row}.`,
      scoreLine
    );
  }

  if (nameScore >= 0.82 && titleScore >= 0.82) {
    const linkReason = record.url && !ref.link
      ? "Tên và kỷ lục khớp; CSDL không có link trong vùng đối chiếu nên chưa xác nhận URL."
      : (!record.url && ref.link
        ? "Tên và kỷ lục khớp; Data đang thiếu URL nhưng CSDL có link."
        : "Tên cá nhân/đơn vị và tên kỷ lục khớp theo nội dung.");
    const fix = record.url && !ref.link
      ? `Không cần sửa nội dung chính. Nếu tab ${quoteSheet(ref.sheet)} có/được thêm cột link, nên bổ sung link từ Data ${quoteSheet(record.sheet)}!F${record.row}.`
      : (!record.url && ref.link
        ? `Bổ sung link từ ${quoteSheet(ref.sheet)}!${refColumns(ref.sheet).link}${ref.row} vào Data ${quoteSheet(record.sheet)}!F${record.row}.`
        : "Không cần sửa.");
    return detailedNote("ĐÃ KHỚP", record, ref, linkReason, fix, scoreLine);
  }

  if (nameScore >= 0.82 && titleScore >= 0.45) {
    return detailedNote(
      "CẦN RÀ SOÁT TÊN KỶ LỤC",
      record,
      ref,
      `CSDL có đúng/gần đúng cá nhân/đơn vị tại ${refId}, nhưng tên kỷ lục chỉ khớp một phần (${titlePct}%).`,
      `Mở Data ${quoteSheet(record.sheet)}!E${record.row} và ${quoteSheet(ref.sheet)}!${refColumns(ref.sheet).title}${ref.row}; nếu là cùng một thành tựu thì chuẩn hóa lại câu chữ ở một bên, nếu là thành tựu khác thì giữ Data và bổ sung thêm một dòng kỷ lục mới vào tab CSDL phù hợp.`,
      scoreLine
    );
  }

  if (nameScore >= 0.82) {
    const examples = result.sameOwner
      .filter((m) => m.ref !== ref)
      .map((m) => `${quoteSheet(m.ref.sheet)}!R${m.ref.row}: ${shortText(m.ref.title, 90)} (${Math.round(m.titleScore * 100)}%)`)
      .join(" | ");
    return detailedNote(
      "THIẾU KỶ LỤC TRONG CSDL",
      record,
      ref,
      `Đã tìm thấy cùng cá nhân/đơn vị tại ${refId}, nhưng các dòng CSDL của người/đơn vị này chưa có tên kỷ lục tương ứng với Data.`,
      `Nếu Data ${quoteSheet(record.sheet)}!E${record.row} là thành tựu thật cần quản lý, bổ sung một dòng mới vào tab CSDL phù hợp cho chủ thể này. Nếu Data đang ghi sai thành tựu, sửa Data E${record.row}; nếu chủ thể sai, sửa Data C${record.row}.`,
      `${scoreLine}${examples ? `; dòng cùng/chung chủ thể khác: ${examples}` : ""}`
    );
  }

  if (result.bestTitle && result.bestTitle.nameScore < 0.78) {
    const m = result.bestTitle;
    return detailedNote(
      "CẦN RÀ SOÁT CHỦ THỂ",
      record,
      m.ref,
      `Tên kỷ lục khá giống ${quoteSheet(m.ref.sheet)}!R${m.ref.row} (${Math.round(m.titleScore * 100)}%) nhưng tên cá nhân/đơn vị không khớp (${Math.round(m.nameScore * 100)}%).`,
      `Mở Data ${quoteSheet(record.sheet)}!C${record.row} và ${quoteSheet(m.ref.sheet)}!${refColumns(m.ref.sheet).owner}${m.ref.row}; nếu CSDL đúng chủ thể thì sửa Data C${record.row}, nếu Data đúng chủ thể thì cần bổ sung/sửa chủ thể ở tab CSDL.`,
      scoreLine
    );
  }

  return detailedNote(
    "THIẾU/CHƯA CHẮC",
    record,
    ref,
    `Chưa có dòng nào khớp chắc; dòng gần nhất chỉ là gợi ý tham khảo, chưa đủ căn cứ để xem là đúng.`,
    `Kiểm tra lại Data ${quoteSheet(record.sheet)}!C${record.row}, E${record.row}, F${record.row}. Nếu thông tin đúng và không có trong 4 tab CSDL thì bổ sung vào tab phù hợp; nếu dòng gần nhất mới là bản đúng thì sửa Data theo các ô CSDL đã nêu.`,
    scoreLine
  );
}

function setCell(rowXml, cellRef, text) {
  const cell = `<c r="${cellRef}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(text)}</t></is></c>`;
  const re = new RegExp(`<c\\b[^>]*\\br="${cellRef}"[^>]*(?:\\/>|>[\\s\\S]*?<\\/c>)`);
  if (re.test(rowXml)) return rowXml.replace(re, cell);
  const cells = [...rowXml.matchAll(/<c\b[^>]*\br="([A-Z]+\d+)"[^>]*(?:\/>|>[\s\S]*?<\/c>)/g)];
  for (const c of cells) {
    if (colIndex(c[1]) > colIndex(cellRef)) return rowXml.slice(0, c.index) + cell + rowXml.slice(c.index);
  }
  return rowXml.replace("</row>", `${cell}</row>`);
}

const refs = loadRefs();
const weights = buildWeights(refs);
const index = buildIndex(refs);
const targetSheets = sheetMap(targetDir);
const targetShared = sharedStrings(targetDir);
const dataSheet = [...targetSheets.values()].find((s) => s.name.startsWith("Data_record_01-01-2022_01-01-20"));
if (!dataSheet) throw new Error("Target Data_record sheet not found.");
const rows = rowsFromSheet(targetDir, dataSheet, targetShared);

const notes = new Map();
const summary = {
  total: 0,
  matched: 0,
  review: 0,
  missingRecord: 0,
  missingUncertain: 0,
};

for (const r of rows) {
  if (r.rowNumber <= 1) continue;
  const record = {
    row: r.rowNumber,
    sheet: dataSheet.name,
    recordId: r.values.get(1) || "",
    gabId: r.values.get(2) || "",
    name: r.values.get(3) || "",
    date: r.values.get(4) || "",
    title: r.values.get(5) || "",
    url: r.values.get(6) || "",
    description: r.values.get(7) || "",
  };
  if (!`${record.name}${record.title}${record.url}`.trim()) continue;
  const note = noteFor(record, analyze(record, refs, index, weights));
  notes.set(r.rowNumber, note);
  summary.total++;
  if (note.includes("Kết luận: ĐÃ KHỚP")) summary.matched++;
  else if (note.includes("Kết luận: THIẾU KỶ LỤC")) summary.missingRecord++;
  else if (note.includes("Kết luận: THIẾU")) summary.missingUncertain++;
  else summary.review++;
}

const sheetPath = path.join(targetDir, dataSheet.target);
let xml = read(sheetPath);
xml = xml.replace(/<dimension\b[^>]*ref="[^"]*"[^/]*\/>/, (m) => m.replace(/ref="[^"]*"/, `ref="A1:H${Math.max(...rows.map((r) => r.rowNumber))}"`));
xml = xml.replace(/<row\b[^>]*\br="1"[^>]*>[\s\S]*?<\/row>/, (m) => setCell(m, "H1", "Note"));
for (const [rowNumber, note] of notes) {
  const re = new RegExp(`<row\\b[^>]*\\br="${rowNumber}"[^>]*>[\\s\\S]*?<\\/row>`);
  xml = xml.replace(re, (m) => setCell(m, `H${rowNumber}`, note));
}
write(sheetPath, xml);

fs.mkdirSync(outDir, { recursive: true });
fs.writeFileSync(path.join(outDir, "summary_deep.json"), JSON.stringify(summary, null, 2), "utf8");
console.log(JSON.stringify({ outputDir: outDir, sheet: dataSheet.name, refs: refs.length, ...summary }, null, 2));
