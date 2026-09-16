import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const root = process.cwd();
const inputXlsx = path.join(root, "outputs", "data_record_import_KLG_chua_co_thanh_tuu_tu_200KLG.xlsx");
const workRoot = path.join(root, "outputs", "pretty_import_workbook");
const unzipDir = path.join(workRoot, "input");
const buildDir = path.join(workRoot, "xlsx");
const zipOut = path.join(workRoot, "pretty.zip");
const outputXlsx = inputXlsx;

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
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
function dec(s = "") {
  return String(s)
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}
function colIndex(ref) {
  let n = 0;
  for (const ch of (String(ref).match(/^[A-Z]+/)?.[0] || "")) {
    n = n * 26 + ch.charCodeAt(0) - 64;
  }
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
function rowsFromFirstSheet(dir, shared) {
  const xml = read(path.join(dir, "xl", "worksheets", "sheet1.xml"));
  const rows = [];
  for (const row of xml.matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>/g)) {
    const rowNumber = Number(row[1].match(/\br="(\d+)"/)?.[1]);
    if (!rowNumber) continue;
    const values = [];
    for (const cell of row[2].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>|<c\b([^>]*)\/>/g)) {
      const attrs = cell[1] || cell[3] || "";
      const ref = attrs.match(/\br="([^"]+)"/)?.[1];
      if (!ref) continue;
      values[colIndex(ref) - 1] = cell[2] ? cellValue(cell[2], attrs, shared) : "";
    }
    rows.push(values);
  }
  return rows;
}
function cell(value, ref, style = 0) {
  const s = style ? ` s="${style}"` : "";
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`;
}
function rowXml(values, rowNumber, styleForCell = () => 0, height = "") {
  const ht = height ? ` ht="${height}" customHeight="1"` : "";
  return `<row r="${rowNumber}"${ht}>${values.map((v, j) => cell(v, `${colLetters(j + 1)}${rowNumber}`, styleForCell(j, v))).join("")}</row>`;
}
function stylesXml() {
  const fonts = [
    `<font><sz val="11"/><name val="Calibri"/></font>`,
    `<font><b/><color rgb="FFFFFFFF"/><sz val="15"/><name val="Calibri"/></font>`,
    `<font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font>`,
    `<font><b/><color rgb="FF1F2937"/><sz val="11"/><name val="Calibri"/></font>`,
  ].join("");
  const fills = [
    `<fill><patternFill patternType="none"/></fill>`,
    `<fill><patternFill patternType="gray125"/></fill>`,
    `<fill><patternFill patternType="solid"><fgColor rgb="FF17365D"/></patternFill></fill>`,
    `<fill><patternFill patternType="solid"><fgColor rgb="FF404040"/></patternFill></fill>`,
    `<fill><patternFill patternType="solid"><fgColor rgb="FFD9EAF7"/></patternFill></fill>`,
    `<fill><patternFill patternType="solid"><fgColor rgb="FFE7F3E8"/></patternFill></fill>`,
    `<fill><patternFill patternType="solid"><fgColor rgb="FFFFF8E1"/></patternFill></fill>`,
    `<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/></patternFill></fill>`,
  ].join("");
  const border = `<border><left style="thin"><color rgb="FFD9E2EC"/></left><right style="thin"><color rgb="FFD9E2EC"/></right><top style="thin"><color rgb="FFD9E2EC"/></top><bottom style="thin"><color rgb="FFD9E2EC"/></bottom><diagonal/></border>`;
  const xf = (fontId, fillId, vertical = "top") => `<xf numFmtId="0" fontId="${fontId}" fillId="${fillId}" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="${vertical}" wrapText="1"/></xf>`;
  const xfs = [
    `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`,
    xf(1, 2, "center"),
    xf(3, 7),
    xf(2, 3, "center"),
    xf(2, 2, "center"),
    xf(3, 4, "center"),
    xf(3, 5),
    xf(3, 6),
    xf(0, 7),
  ].join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="4">${fonts}</fonts><fills count="8">${fills}</fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>${border}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="9">${xfs}</cellXfs></styleSheet>`;
}
function buildSheet(records) {
  const widths = [30, 28, 54, 18, 52, 28, 16, 72, 54, 90, 22, 28, 24, 18, 18, 52];
  const cols = `<cols>${widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("")}</cols>`;
  const rows = [];
  const empty = new Array(16).fill("");
  rows.push(rowXml(["BẢNG RÀ SOÁT 200 KLG - THÀNH TỰU CẦN XUẤT THEO FORM DATABASE", ...empty.slice(1)], 1, () => 1, 28));
  rows.push(rowXml([`Kết quả: ${records.length} dòng thành tựu lấy từ tab 200 KLG CHUẨN HÓA, vì KLG vẫn nằm trong danh sách khách báo chưa có thành tựu trên GAB.`, ...empty.slice(1)], 2, () => 2, 24));
  rows.push(rowXml(["Cách đọc: A-C là kết luận; D-J là form database để import; K-P là vết đối chiếu để kiểm tra lại khi cần.", ...empty.slice(1)], 3, () => 2, 24));
  rows.push(rowXml(empty, 4, () => 0, 6));
  const headers = [
    "KẾT LUẬN",
    "MỨC TIN CẬY",
    "VÌ SAO CÓ DÒNG NÀY",
    "recordId",
    "gabId / Link hồ sơ GAB",
    "fullName / Tên KLG",
    "time / Thời gian",
    "title / Tiêu đề thành tựu",
    "url / Link bài viết nguồn",
    "description / Nội dung mô tả",
    "Cách khớp",
    "Tên trong DS khách",
    "Trạng thái khách ghi",
    "Dòng DS khách",
    "Dòng tab 200 KLG",
    "Nguồn/link trong DS khách",
  ];
  rows.push(rowXml(headers, 5, (j) => (j <= 2 ? 3 : j <= 9 ? 4 : 5), 38));
  records.forEach((r, idx) => {
    const matchMethod = r[7] || "";
    const byLink = /LINK/i.test(matchMethod);
    const conclusion = byLink ? "CẦN IMPORT THÀNH TỰU" : "CẦN KIỂM TRA TRƯỚC KHI IMPORT";
    const confidence = byLink ? "CHẮC - cùng link GAB" : "CẦN RÀ NHANH - chỉ khớp tên";
    const reason = byLink
      ? "KLG có trong DS khách báo chưa có thành tựu trên GAB và khớp đúng link hồ sơ GAB với tab 200 KLG."
      : "KLG có trong DS khách báo chưa có thành tựu trên GAB nhưng không khớp link; đang ghép theo tên để tránh sót.";
    const values = [
      conclusion,
      confidence,
      reason,
      r[0] || "",
      r[1] || "",
      r[2] || "",
      r[3] || "",
      r[4] || "",
      r[5] || "",
      r[6] || "",
      matchMethod,
      r[8] || "",
      r[10] || "",
      r[12] || "",
      r[13] || "",
      r[11] || "",
    ];
    rows.push(rowXml(values, idx + 6, (j) => (j <= 2 ? (byLink ? 6 : 7) : 8), 60));
  });
  const maxRow = records.length + 5;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><dimension ref="A1:P${maxRow}"/>${cols}<sheetData>${rows.join("")}</sheetData><autoFilter ref="A5:P${maxRow}"/><mergeCells count="3"><mergeCell ref="A1:P1"/><mergeCell ref="A2:P2"/><mergeCell ref="A3:P3"/></mergeCells></worksheet>`;
}
function writeWorkbook(records) {
  fs.mkdirSync(path.join(buildDir, "_rels"), { recursive: true });
  fs.mkdirSync(path.join(buildDir, "xl", "_rels"), { recursive: true });
  fs.mkdirSync(path.join(buildDir, "xl", "worksheets"), { recursive: true });
  fs.mkdirSync(path.join(buildDir, "docProps"), { recursive: true });
  fs.writeFileSync(path.join(buildDir, "[Content_Types].xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`);
  fs.writeFileSync(path.join(buildDir, "_rels", ".rels"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  fs.writeFileSync(path.join(buildDir, "xl", "workbook.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Ra soat import" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  fs.writeFileSync(path.join(buildDir, "xl", "_rels", "workbook.xml.rels"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  fs.writeFileSync(path.join(buildDir, "xl", "styles.xml"), stylesXml());
  fs.writeFileSync(path.join(buildDir, "xl", "worksheets", "sheet1.xml"), buildSheet(records));
  fs.writeFileSync(path.join(buildDir, "docProps", "core.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Ra soat import 200 KLG</dc:title></cp:coreProperties>`);
  fs.writeFileSync(path.join(buildDir, "docProps", "app.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Codex</Application></Properties>`);
}

resetDir(workRoot);
resetDir(unzipDir);
resetDir(buildDir);
fs.copyFileSync(inputXlsx, path.join(workRoot, "input.zip"));
ps(`Expand-Archive -LiteralPath '${path.join(workRoot, "input.zip").replace(/'/g, "''")}' -DestinationPath '${unzipDir.replace(/'/g, "''")}' -Force`);
const records = rowsFromFirstSheet(unzipDir, sharedStrings(unzipDir)).slice(1);
writeWorkbook(records);
if (fs.existsSync(zipOut)) fs.rmSync(zipOut, { force: true });
ps(`Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::CreateFromDirectory('${buildDir.replace(/'/g, "''")}', '${zipOut.replace(/'/g, "''")}')`);
fs.copyFileSync(zipOut, outputXlsx);
console.log(JSON.stringify({ outputXlsx, rows: records.length }, null, 2));
