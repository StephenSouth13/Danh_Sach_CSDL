import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const sourceDir = "outputs\\reconcile_grouped_check_notes_work\\target";
const outDir = path.resolve("outputs", "hoi_ngo_priority_work");
const xlsxDir = path.join(outDir, "xlsx");
const outputXlsx = path.resolve("outputs", "reconciled_gab_records_customer_feedback_hoi_ngo_priority.xlsx");

function resetDir(dir) {
  if (fs.existsSync(dir)) fs.rmSync(dir, { recursive: true, force: true });
  fs.mkdirSync(dir, { recursive: true });
}
function esc(s = "") {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function dec(s = "") {
  return String(s).replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
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
function colIndex(ref) {
  let n = 0;
  for (const ch of (String(ref).match(/^[A-Z]+/)?.[0] || "")) n = n * 26 + ch.charCodeAt(0) - 64;
  return n;
}
function sharedStrings(dir) {
  const p = path.join(dir, "xl", "sharedStrings.xml");
  if (!fs.existsSync(p)) return [];
  const xml = fs.readFileSync(p, "utf8");
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

const shared = sharedStrings(sourceDir);
const sheetXml = fs.readFileSync(path.join(sourceDir, "xl", "worksheets", "sheet1.xml"), "utf8");
const rows = [];
for (const row of sheetXml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)) {
  const rowNumber = Number(row[1]);
  const values = new Map();
  for (const cell of row[2].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
    const ref = cell[1].match(/\br="([^"]+)"/)?.[1];
    if (!ref) continue;
    values.set(colIndex(ref), cellValue(cell[2], cell[1], shared));
  }
  if (rowNumber === 1 || values.get(12) === "Ưu tiên hội ngộ") {
    rows.push({ rowNumber, values });
  }
}

const headers = [
  "Dòng trong file đối chiếu",
  "recordId",
  "gabId",
  "fullName",
  "time",
  "title",
  "url",
  "description",
  "Check",
  "Note",
  "Tỉnh",
  "LĨNH VỰC XÁC LẬP",
  "Ưu tiên",
];
const data = [headers];
for (const r of rows.filter((r) => r.rowNumber > 1)) {
  data.push([
    r.rowNumber,
    r.values.get(1) || "",
    r.values.get(2) || "",
    r.values.get(3) || "",
    r.values.get(4) || "",
    r.values.get(5) || "",
    r.values.get(6) || "",
    r.values.get(7) || "",
    r.values.get(8) || "",
    r.values.get(9) || "",
    r.values.get(10) || "",
    r.values.get(11) || "",
    r.values.get(12) || "",
  ]);
}

resetDir(outDir);
resetDir(xlsxDir);
fs.mkdirSync(path.join(xlsxDir, "_rels"), { recursive: true });
fs.mkdirSync(path.join(xlsxDir, "xl", "_rels"), { recursive: true });
fs.mkdirSync(path.join(xlsxDir, "xl", "worksheets"), { recursive: true });
fs.mkdirSync(path.join(xlsxDir, "docProps"), { recursive: true });

let sheetData = "";
data.forEach((row, i) => {
  const r = i + 1;
  let cells = "";
  row.forEach((value, j) => {
    const ref = `${colLetters(j + 1)}${r}`;
    cells += `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`;
  });
  sheetData += `<row r="${r}">${cells}</row>`;
});

fs.writeFileSync(path.join(xlsxDir, "[Content_Types].xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`);
fs.writeFileSync(path.join(xlsxDir, "_rels", ".rels"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
fs.writeFileSync(path.join(xlsxDir, "xl", "workbook.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Uu tien hoi ngo" sheetId="1" r:id="rId1"/></sheets></workbook>`);
fs.writeFileSync(path.join(xlsxDir, "xl", "_rels", "workbook.xml.rels"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
fs.writeFileSync(path.join(xlsxDir, "xl", "styles.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9EAD3"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>`);
fs.writeFileSync(path.join(xlsxDir, "xl", "worksheets", "sheet1.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:M${data.length}"/><cols><col min="1" max="1" width="16" customWidth="1"/><col min="2" max="3" width="36" customWidth="1"/><col min="4" max="4" width="28" customWidth="1"/><col min="5" max="5" width="14" customWidth="1"/><col min="6" max="6" width="60" customWidth="1"/><col min="9" max="10" width="42" customWidth="1"/><col min="11" max="13" width="24" customWidth="1"/></cols><sheetData>${sheetData}</sheetData><autoFilter ref="A1:M${data.length}"/></worksheet>`);
fs.writeFileSync(path.join(xlsxDir, "docProps", "core.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Ưu tiên hội ngộ</dc:title></cp:coreProperties>`);
fs.writeFileSync(path.join(xlsxDir, "docProps", "app.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Codex</Application></Properties>`);

const zipOut = path.join(outDir, "priority.zip");
if (fs.existsSync(outputXlsx)) fs.rmSync(outputXlsx, { force: true });
execFileSync("powershell", ["-NoProfile", "-Command", `Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::CreateFromDirectory('${xlsxDir.replace(/'/g, "''")}', '${zipOut.replace(/'/g, "''")}')`]);
fs.copyFileSync(zipOut, outputXlsx);
console.log(JSON.stringify({ rows: data.length - 1, outputXlsx }, null, 2));
