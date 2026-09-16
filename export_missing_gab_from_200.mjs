import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const root = process.cwd();
const dbListXlsx = path.join(root, "DỮ LIỆU THÀNH TỰU GAB đối chiếu CSDL ĐÃ CHUẨN HÓA (1).xlsx");
const source200Xlsx = path.join(root, "Dulieutruyxuat", "Bản sao của DỮ LIỆU THÀNH TỰU GAB - 15_24, 13 tháng 9.xlsx");
const outRoot = path.join(root, "outputs", "export_missing_gab_from_200");
const dbDir = path.join(outRoot, "db_list");
const srcDir = path.join(outRoot, "src200");
const buildDir = path.join(outRoot, "xlsx");
const outputXlsx = path.join(root, "outputs", "data_record_import_KLG_chua_co_thanh_tuu_tu_200KLG.xlsx");

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
    .replace(/\([^)]*\)/g, " ")
    .replace(/\b(klg|ky luc gia|gs|pgs|ts|ths|bs|ong|ba|anh|chi|hoa si|nha bao|nghe si|nghe nhan|giao su|vien si|tien si|thac si|danh du)\b/g, " ")
    .replace(/[^a-z0-9]+/g, " ").replace(/\s+/g, " ").trim();
}
function gabKey(url = "") {
  const m = String(url).trim().match(/gab\.world\/vi\/bank\/([^/?#]+)/i);
  return m ? m[1].toLowerCase() : "";
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
function validTime(value = "") {
  const s = String(value ?? "").trim();
  if (!s) return "";
  if (/^\d+(\.0+)?$/.test(s)) {
    const serial = Number(s);
    if (serial <= 20000 || serial >= 70000) {
      return /^(19|20)\d{2}$/.test(s) ? s : "";
    }
  }
  const text = excelDateText(s);
  return /\b(\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}|\d{1,2}[\/\-.]\d{4}|(19|20)\d{2})\b/.test(text) ? text : "";
}
function cleanUrl(url = "") {
  const s = String(url || "").trim();
  if (!/^https?:\/\//i.test(s)) return "";
  if (/không tìm|chưa có/i.test(s)) return "";
  return s;
}
function pickTime(row) {
  return [row.dateGab, row.dateOfficial, row.timeInContent].map(validTime).find(Boolean) || "";
}
function titleFor(row) {
  return row.gabTitle || (row.type ? `Xác lập ${row.type} "${row.title}"` : row.title);
}
function row200(r) {
  return {
    row: r.rowNumber,
    stt: String(r.values.get(1) || "").trim(),
    id: String(r.values.get(3) || "").trim(),
    dateOfficial: String(r.values.get(4) || "").trim(),
    title: String(r.values.get(5) || "").trim(),
    gabLink: String(r.values.get(6) || "").trim(),
    fullName: String(r.values.get(7) || "").trim(),
    type: String(r.values.get(11) || "").trim(),
    dateGab: String(r.values.get(13) || "").trim(),
    gabTitle: String(r.values.get(14) || "").trim(),
    articleLink: String(r.values.get(15) || "").trim(),
    desc: String(r.values.get(16) || "").trim(),
    contentGab: String(r.values.get(20) || "").trim(),
    timeInContent: String(r.values.get(21) || "").trim(),
  };
}

resetDir(outRoot);
resetDir(dbDir);
resetDir(srcDir);
resetDir(buildDir);
fs.copyFileSync(dbListXlsx, path.join(outRoot, "db_list.zip"));
fs.copyFileSync(source200Xlsx, path.join(outRoot, "src200.zip"));
ps(`Expand-Archive -LiteralPath '${path.join(outRoot, "db_list.zip").replace(/'/g, "''")}' -DestinationPath '${dbDir.replace(/'/g, "''")}' -Force`);
ps(`Expand-Archive -LiteralPath '${path.join(outRoot, "src200.zip").replace(/'/g, "''")}' -DestinationPath '${srcDir.replace(/'/g, "''")}' -Force`);

const dbRows = rowsFromSheet(dbDir, sheetMap(dbDir).get("Trang tính1"), sharedStrings(dbDir))
  .filter((r) => r.rowNumber > 1)
  .map((r) => ({
    row: r.rowNumber,
    name: String(r.values.get(2) || "").trim(),
    title: String(r.values.get(3) || "").trim(),
    status: String(r.values.get(5) || "").trim(),
    source: String(r.values.get(7) || "").trim(),
    nameKey: normalize(r.values.get(2) || ""),
    gab: gabKey(r.values.get(7) || ""),
  }))
  .filter((r) => /chưa có trên gab/i.test(r.status));

const dbGabKeys = new Set(dbRows.map((r) => r.gab).filter(Boolean));
const dbNames = new Set(dbRows.map((r) => r.nameKey).filter(Boolean));

const sourceRows = rowsFromSheet(srcDir, sheetMap(srcDir).get("200 KLG CHUẨN HÓA"), sharedStrings(srcDir))
  .filter((r) => r.rowNumber >= 3)
  .map(row200)
  .filter((r) => `${r.fullName}${r.title}${r.gabTitle}`.trim());

const exported = [];
const seen = new Set();
for (const row of sourceRows) {
  const gKey = gabKey(row.gabLink);
  const nKey = normalize(row.fullName);
  const byLink = gKey && dbGabKeys.has(gKey);
  const byName = nKey && dbNames.has(nKey);
  if (!byLink && !byName) continue;
  const title = titleFor(row);
  if (!title.trim()) continue;
  const key = [gKey, nKey, normalize(title), pickTime(row), cleanUrl(row.articleLink)].join("|");
  if (seen.has(key)) continue;
  seen.add(key);
  const matchedDb = dbRows.find((r) => (gKey && r.gab === gKey) || (nKey && r.nameKey === nKey));
  exported.push([
    "",
    row.gabLink,
    row.fullName,
    pickTime(row),
    title,
    cleanUrl(row.articleLink),
    row.desc || row.contentGab,
    byLink ? "Khớp bằng LINK GAB" : "Khớp bằng TÊN",
    matchedDb ? matchedDb.name : "",
    matchedDb ? matchedDb.title : "",
    matchedDb ? matchedDb.status : "",
    matchedDb ? matchedDb.source : "",
    matchedDb ? matchedDb.row : "",
    row.row,
    byLink ? "Ưu tiên cao: cùng link GAB với danh sách khách báo chưa có thành tựu." : "Cần rà nhanh: không khớp link GAB, đang lấy theo tên.",
  ]);
}

function cell(value, ref, style = 0) {
  const s = style ? ` s="${style}"` : "";
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`;
}
function sheetXml(rows) {
  const widths = [18, 52, 26, 16, 70, 52, 90, 20, 28, 60, 24, 52, 18, 18, 70];
  const cols = `<cols>${widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("")}</cols>`;
  const sheetData = rows.map((row, i) => {
    const r = i + 1;
    return `<row r="${r}">${row.map((v, j) => cell(v, `${colLetters(j + 1)}${r}`, i === 0 ? 1 : 0)).join("")}</row>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><dimension ref="A1:O${rows.length}"/>${cols}<sheetData>${sheetData}</sheetData><autoFilter ref="A1:O${rows.length}"/></worksheet>`;
}
function writeWorkbook(rows) {
  fs.mkdirSync(path.join(buildDir, "_rels"), { recursive: true });
  fs.mkdirSync(path.join(buildDir, "xl", "_rels"), { recursive: true });
  fs.mkdirSync(path.join(buildDir, "xl", "worksheets"), { recursive: true });
  fs.mkdirSync(path.join(buildDir, "docProps"), { recursive: true });
  fs.writeFileSync(path.join(buildDir, "[Content_Types].xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`);
  fs.writeFileSync(path.join(buildDir, "_rels", ".rels"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  fs.writeFileSync(path.join(buildDir, "xl", "workbook.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data_record_import" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  fs.writeFileSync(path.join(buildDir, "xl", "_rels", "workbook.xml.rels"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  fs.writeFileSync(path.join(buildDir, "xl", "styles.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>`);
  fs.writeFileSync(path.join(buildDir, "xl", "worksheets", "sheet1.xml"), sheetXml(rows));
  fs.writeFileSync(path.join(buildDir, "docProps", "core.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Data record import</dc:title></cp:coreProperties>`);
  fs.writeFileSync(path.join(buildDir, "docProps", "app.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Codex</Application></Properties>`);
}

const headers = [
  "recordId",
  "gabId / Link hồ sơ GAB",
  "fullName / Tên KLG",
  "time / Thời gian import",
  "title / Tiêu đề thành tựu import",
  "url / Link bài viết nguồn",
  "description / Nội dung mô tả import",
  "Cách đối chiếu",
  "Tên trong DS khách",
  "Tiêu đề trong DS khách",
  "Trạng thái khách ghi",
  "Nguồn/link trong DS khách",
  "Dòng DS khách",
  "Dòng tab 200 KLG",
  "Ghi chú rà soát",
];
writeWorkbook([headers, ...exported]);
const zipOut = path.join(outRoot, "data_record_import.zip");
if (fs.existsSync(zipOut)) fs.rmSync(zipOut, { force: true });
if (fs.existsSync(outputXlsx)) fs.rmSync(outputXlsx, { force: true });
ps(`Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::CreateFromDirectory('${buildDir.replace(/'/g, "''")}', '${zipOut.replace(/'/g, "''")}')`);
fs.copyFileSync(zipOut, outputXlsx);
console.log(JSON.stringify({ outputXlsx, dbMissingRows: dbRows.length, exportedRows: exported.length }, null, 2));
