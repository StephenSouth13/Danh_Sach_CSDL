import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const root = process.cwd();
const sourceXlsx = path.join(root, "Dulieutruyxuat", "Bản sao của DỮ LIỆU THÀNH TỰU GAB - 15_24, 13 tháng 9.xlsx");
const refXlsx = path.join(root, "Dulieutruyxuat", "MÔ TẢ YÊU CẦU NHẬP THÔNG TIN GAB.xlsx");
const outDir = path.join(root, "outputs", "csdl_authoritative_clean_work");
const srcDir = path.join(outDir, "src");
const refDir = path.join(outDir, "ref");
const outputXlsx = path.join(root, "outputs", "reconciled_gab_records_CSDL_authoritative_clean_notes.xlsx");

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
function dec(s = "") {
  return String(s)
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
}
function esc(s = "") {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
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
    .replace(/\([^)]*\)/g, " ")
    .replace(/\b(ky|ki)\s*luc\s*gia\b/g, " ")
    .replace(/\b(gs|pgs|ts|ths|thac sy|tien si|bac si|ong|ba|anh|chi|hoa si|nghe si|nghe nhan|nha tho|nha van|nha bao|doanh nhan|giao su|vien si|nsut|nsnd|uu tu|gd|gdnd)\b/g, " ")
    .replace(/[^a-z0-9]+/g, " ").replace(/\s+/g, " ").trim();
}
const stop = new Set("xac lap ky luc viet nam the gioi chau a dong duong nguoi dau tien nhieu nhat lon nhat cao nhat co cua va voi trong tai nam ngay bo tac pham duoc so huu".split(" "));
function tokens(s = "", keepStop = false) {
  const n = normalize(s);
  return n ? [...new Set(n.split(" ").filter((x) => x.length > 1 && (keepStop || !stop.has(x))))] : [];
}
function score(a, b) {
  if (!a.length || !b.length) return 0;
  const as = new Set(a), bs = new Set(b);
  let c = 0;
  for (const t of as) if (bs.has(t)) c++;
  return c / Math.max(as.size, bs.size);
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
  const js = s.match(/^(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+([A-Z][a-z]{2})\s+(\d{1,2})\s+(\d{4})/);
  if (js) {
    const mo = { Jan: "01", Feb: "02", Mar: "03", Apr: "04", May: "05", Jun: "06", Jul: "07", Aug: "08", Sep: "09", Oct: "10", Nov: "11", Dec: "12" };
    return `${String(js[2]).padStart(2, "0")}/${mo[js[1]]}/${js[3]}`;
  }
  return s.replace(/\s+/g, " ");
}
function dateKey(value = "") {
  const s = excelDateText(value);
  const dmy = s.match(/\b(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})\b/);
  if (dmy) return `${dmy[3].length === 2 ? "20" + dmy[3] : dmy[3]}-${dmy[2].padStart(2, "0")}-${dmy[1].padStart(2, "0")}`;
  const year = s.match(/\b(19|20)\d{2}\b/)?.[0];
  return year || normalize(s);
}
function short(s = "", max = 130) {
  const t = String(s ?? "").replace(/\s+/g, " ").trim();
  return t.length > max ? `${t.slice(0, max - 3)}...` : t;
}
function refMeta(ref) {
  return `${ref.sheet}!${ref.row}: ${short(ref.title, 100)} (${excelDateText(ref.date) || "chưa rõ ngày"})`;
}
function setCell(xml, rowNumber, col, value) {
  const ref = `${colLetters(col)}${rowNumber}`;
  const cellXml = `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`;
  const rowRe = new RegExp(`<row\\b([^>]*)\\br="${rowNumber}"([^>]*)>([\\s\\S]*?)<\\/row>`);
  const rowMatch = xml.match(rowRe);
  if (!rowMatch) return xml;
  let body = rowMatch[3];
  const cellRe = /<c\b[^>]*\br="([^"]+)"[^>]*(?:\/>|>[\s\S]*?<\/c>)/g;
  let replaced = false;
  body = body.replace(cellRe, (match, cellRef) => {
    if (cellRef === ref) {
      replaced = true;
      return cellXml;
    }
    return match;
  });
  if (!replaced) body += cellXml;
  return xml.replace(rowRe, `<row${rowMatch[1]}r="${rowNumber}"${rowMatch[2]}>${body}</row>`);
}
function updateDimension(xml, ref) {
  return xml.replace(/<dimension\b[^>]*\/>/, `<dimension ref="${ref}"/>`);
}
function refCols(sheetName, r) {
  if (sheetName === "CSDL KLVN") return { owner: r.values.get(6), title: r.values.get(5), date: r.values.get(4), province: r.values.get(7), field: r.values.get(8) };
  if (sheetName === "CSDL KLVW") return { owner: r.values.get(7), title: r.values.get(6), date: r.values.get(5), province: r.values.get(9) || r.values.get(8), field: r.values.get(10) || "" };
  return { owner: r.values.get(2), title: r.values.get(3), date: r.values.get(5) || r.values.get(4), province: r.values.get(7), field: "" };
}

resetDir(outDir);
fs.mkdirSync(srcDir, { recursive: true });
fs.mkdirSync(refDir, { recursive: true });
fs.copyFileSync(sourceXlsx, path.join(outDir, "source.zip"));
fs.copyFileSync(refXlsx, path.join(outDir, "ref.zip"));
ps(`Expand-Archive -LiteralPath '${path.join(outDir, "source.zip").replace(/'/g, "''")}' -DestinationPath '${srcDir.replace(/'/g, "''")}' -Force`);
ps(`Expand-Archive -LiteralPath '${path.join(outDir, "ref.zip").replace(/'/g, "''")}' -DestinationPath '${refDir.replace(/'/g, "''")}' -Force`);

const srcShared = sharedStrings(srcDir);
const refShared = sharedStrings(refDir);
const srcSheets = sheetMap(srcDir);
const refSheets = sheetMap(refDir);
const dataSheet = [...srcSheets.values()].find((s) => s.name.startsWith("Data_record"));
const hoiNgoSheet = [...srcSheets.values()].find((s) => normalize(s.name).includes("danh sach hoi ngo"));
if (!dataSheet || !hoiNgoSheet) throw new Error("Khong thay tab Data_record hoac Danh sach Hoi Ngo KLG");

const refRows = [];
for (const name of ["CSDL KLVN", "CSDL KLVW", "CSDL KL CA", "CSDL KLTG"]) {
  const sheet = refSheets.get(name);
  if (!sheet) continue;
  for (const r of rowsFromSheet(refDir, sheet, refShared)) {
    if (r.rowNumber < 3) continue;
    const raw = refCols(name, r);
    const owner = String(raw.owner ?? "").trim();
    const title = String(raw.title ?? "").trim();
    if (!owner || !title) continue;
    refRows.push({
      sheet: name, row: r.rowNumber, owner, title,
      date: String(raw.date ?? "").trim(),
      province: String(raw.province ?? "").trim(),
      field: String(raw.field ?? "").trim(),
      ownerKey: normalize(owner),
      ownerTokens: tokens(owner, true),
      titleTokens: tokens(title),
      dateKey: dateKey(raw.date ?? ""),
    });
  }
}
const seenRefs = new Set();
const refs = refRows.filter((r) => {
  const key = `${r.ownerKey}|${normalize(r.title)}|${r.dateKey}`;
  if (seenRefs.has(key)) return false;
  seenRefs.add(key);
  return true;
});
function ownerRefs(name) {
  const key = normalize(name);
  const nt = tokens(name, true);
  if (!key || nt.length < 2) return [];
  return refs.filter((r) => r.ownerKey === key || r.ownerKey.includes(key) || key.includes(r.ownerKey) || score(nt, r.ownerTokens) >= 0.82);
}
function bestRefForData(row, candidates) {
  let best = null;
  const dt = tokens(row.title);
  for (const ref of candidates) {
    const s = score(dt, ref.titleTokens);
    if (!best || s > best.score) best = { ref, score: s };
  }
  return best;
}

const dataRows = rowsFromSheet(srcDir, dataSheet, srcShared).filter((r) => r.rowNumber > 1).map((r) => ({
  rowNumber: r.rowNumber,
  recordId: String(r.values.get(1) ?? "").trim(),
  gabId: String(r.values.get(2) ?? "").trim(),
  fullName: String(r.values.get(3) ?? "").trim(),
  time: String(r.values.get(4) ?? "").trim(),
  title: String(r.values.get(5) ?? "").trim(),
  url: String(r.values.get(6) ?? "").trim(),
  desc: String(r.values.get(7) ?? "").trim(),
})).filter((r) => r.recordId || r.fullName || r.title);
const byName = new Map();
const rowResult = new Map();
for (const row of dataRows) {
  const candidates = ownerRefs(row.fullName);
  const best = bestRefForData(row, candidates);
  let check, note, matchedRef = null;
  if (!candidates.length) {
    check = "SAI/CẦN RÀ: Không tìm thấy kỷ lục gia trong CSDL chuẩn";
    note = `Tên "${row.fullName}" ở Data dòng ${row.rowNumber} chưa tìm thấy trong 4 tab CSDL chuẩn. Cần kiểm tra biến thể tên/tên nghệ danh hoặc xác nhận đây là dữ liệu ngoài CSDL.`;
  } else if (!best || best.score < 0.38) {
    check = "SAI/CẦN RÀ: Dư hoặc chưa khớp tên kỷ lục";
    note = `Có kỷ lục gia trong CSDL chuẩn nhưng tên thành tựu trên GAB chưa khớp rõ với CSDL. Data dòng ${row.rowNumber}: "${short(row.title)}". CSDL gần nhất: ${candidates.slice(0, 3).map(refMeta).join("; ")}.`;
  } else {
    matchedRef = best.ref;
    const dataDate = dateKey(row.time);
    const refDate = matchedRef.dateKey;
    const dateOk = dataDate && refDate && (dataDate === refDate || dataDate.includes(refDate) || refDate.includes(dataDate));
    if (dateOk) {
      check = "ĐÚNG: Khớp CSDL chuẩn";
      note = `Khớp theo kỷ lục gia, tên kỷ lục và thời gian xác lập với ${refMeta(matchedRef)}. Không cần xử lý thành tựu này.`;
    } else {
      check = "SAI: Sai thời gian xác lập";
      note = `Tên kỷ lục khớp CSDL nhưng ngày trên GAB/Data là "${excelDateText(row.time) || "trống"}", ngày chuẩn CSDL là "${excelDateText(matchedRef.date) || "trống"}". Cần sửa thời gian xác lập trên GAB theo ${refMeta(matchedRef)}.`;
    }
  }
  rowResult.set(row.rowNumber, { check, note, matchedRef });
  const key = normalize(row.fullName);
  if (!byName.has(key)) byName.set(key, []);
  byName.get(key).push({ row, result: rowResult.get(row.rowNumber) });
}
for (const [key, items] of byName) {
  const name = items[0].row.fullName;
  const candidates = ownerRefs(name);
  const matched = new Set(items.map((x) => x.result.matchedRef).filter(Boolean).map((r) => `${r.sheet}|${r.row}`));
  const missing = candidates.filter((r) => !matched.has(`${r.sheet}|${r.row}`));
  if (missing.length) {
    const missingText = `CSDL còn ${missing.length} thành tựu của kỷ lục gia này chưa thấy khớp trong Data_record: ${missing.slice(0, 6).map(refMeta).join("; ")}${missing.length > 6 ? "; ..." : ""}.`;
    for (const item of items) {
      item.result.note += `\n${missingText}`;
      if (item.result.check.startsWith("ĐÚNG")) item.result.check = "SAI/CẦN BỔ SUNG: GAB còn thiếu thành tựu khác";
    }
  }
}

let dataXml = read(path.join(srcDir, dataSheet.target));
dataXml = updateDimension(dataXml, `A1:I${Math.max(...dataRows.map((r) => r.rowNumber), 1)}`);
dataXml = setCell(dataXml, 1, 8, "CHECK ĐÚNG/SAI");
dataXml = setCell(dataXml, 1, 9, "NOTE THIẾU/DƯ/CẦN XỬ LÝ");
let dataStats = { ok: 0, wrongDate: 0, missing: 0, extra: 0 };
for (const row of dataRows) {
  const result = rowResult.get(row.rowNumber);
  if (result.check.startsWith("ĐÚNG")) dataStats.ok++;
  if (result.check.includes("Sai thời gian")) dataStats.wrongDate++;
  if (result.check.includes("BỔ SUNG")) dataStats.missing++;
  if (result.check.includes("Dư") || result.check.includes("Không tìm thấy")) dataStats.extra++;
  dataXml = setCell(dataXml, row.rowNumber, 8, result.check);
  dataXml = setCell(dataXml, row.rowNumber, 9, result.note);
}
write(path.join(srcDir, dataSheet.target), dataXml.replace(/<\/c>>+/g, "</c>"));

const hoiRows = rowsFromSheet(srcDir, hoiNgoSheet, srcShared).filter((r) => r.rowNumber >= 4 && r.rowNumber <= 303);
let hoiXml = read(path.join(srcDir, hoiNgoSheet.target));
hoiXml = updateDimension(hoiXml, "A1:Z303");
hoiXml = setCell(hoiXml, 1, 25, "CHECK ĐÚNG/SAI");
hoiXml = setCell(hoiXml, 1, 26, "NOTE THIẾU/DƯ/CẦN XỬ LÝ");
let hoiStats = { rows: hoiRows.length, withName: 0, blankName: 0, hasIssue: 0 };
for (const r of hoiRows) {
  const name = String(r.values.get(7) || r.values.get(9) || "").trim();
  if (!name) {
    hoiStats.blankName++;
    hoiXml = setCell(hoiXml, r.rowNumber, 25, "CẦN BỔ SUNG TÊN");
    hoiXml = setCell(hoiXml, r.rowNumber, 26, "Dòng này chưa có họ tên ở cột G/I nên chưa đủ căn cứ đối chiếu với CSDL chuẩn.");
    continue;
  }
  hoiStats.withName++;
  const candidates = ownerRefs(name);
  const gabItems = byName.get(normalize(name)) || dataRows
    .filter((d) => normalize(d.fullName).includes(normalize(name)) || normalize(name).includes(normalize(d.fullName)))
    .map((row) => ({ row, result: rowResult.get(row.rowNumber) }));
  const wrong = gabItems.filter((x) => x.result?.check.includes("Sai thời gian"));
  const extra = gabItems.filter((x) => x.result?.check.includes("Dư") || x.result?.check.includes("Không tìm thấy"));
  const matchedRefs = new Set(gabItems.map((x) => x.result?.matchedRef).filter(Boolean).map((m) => `${m.sheet}|${m.row}`));
  const missing = candidates.filter((c) => !matchedRefs.has(`${c.sheet}|${c.row}`));
  let check = "ĐÚNG/CHƯA THẤY LỖI THÀNH TỰU";
  if (!candidates.length) check = "CẦN RÀ: Không thấy tên trong CSDL chuẩn";
  else if (missing.length || wrong.length || extra.length) check = "SAI/CẦN XỬ LÝ";
  if (check !== "ĐÚNG/CHƯA THẤY LỖI THÀNH TỰU") hoiStats.hasIssue++;
  const note = [
    `Đối chiếu "${name}" với 4 tab CSDL chuẩn: CSDL có ${candidates.length} thành tựu, Data/GAB tìm thấy ${gabItems.length} dòng.`,
    missing.length ? `Thiếu trên GAB/Data: ${missing.length} thành tựu. Cần bổ sung/kiểm tra các mục: ${missing.slice(0, 6).map(refMeta).join("; ")}${missing.length > 6 ? "; ..." : ""}.` : "Chưa ghi nhận thiếu thành tựu so với CSDL chuẩn.",
    wrong.length ? `Sai thời gian xác lập: ${wrong.length} dòng Data cần sửa ngày theo CSDL, gồm ${wrong.slice(0, 4).map((x) => `Data dòng ${x.row.rowNumber}`).join(", ")}.` : "Chưa ghi nhận sai thời gian xác lập.",
    extra.length ? `Dư/chưa khớp CSDL: ${extra.length} dòng Data cần xác minh có phải dữ liệu ngoài CSDL hay cần chuẩn hóa.` : "Chưa ghi nhận dòng dư/chưa khớp CSDL."
  ].join("\n");
  hoiXml = setCell(hoiXml, r.rowNumber, 25, check);
  hoiXml = setCell(hoiXml, r.rowNumber, 26, note);
}
write(path.join(srcDir, hoiNgoSheet.target), hoiXml.replace(/<\/c>>+/g, "</c>"));

const outZip = path.join(outDir, "out.zip");
if (fs.existsSync(outZip)) fs.rmSync(outZip, { force: true });
ps(`Compress-Archive -Path '${path.join(srcDir, "*").replace(/'/g, "''")}' -DestinationPath '${outZip.replace(/'/g, "''")}' -Force`);
fs.copyFileSync(outZip, outputXlsx);

console.log(JSON.stringify({
  outputXlsx,
  refRowsRaw: refRows.length,
  refRowsAfterDedupe: refs.length,
  dataRows: dataRows.length,
  dataStats,
  hoiNgoStats: hoiStats
}, null, 2));
