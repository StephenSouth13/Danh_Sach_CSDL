import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const inputXlsx = path.resolve("outputs", "reconciled_gab_records_customer_feedback_hoi_ngo_priority.xlsx");
const outputXlsx = path.resolve("outputs", "reconciled_gab_records_customer_feedback_hoi_ngo_priority_dedup.xlsx");
const workDir = path.resolve("outputs", "hoi_ngo_priority_dedup_work");
const zipPath = path.join(workDir, "book.zip");
const outDir = path.join(workDir, "xlsx");
const outZip = path.join(workDir, "out.zip");

function ps(command) {
  execFileSync("powershell", ["-NoProfile", "-Command", command], { stdio: "pipe" });
}
function resetDir(dir) {
  if (fs.existsSync(dir)) fs.rmSync(dir, { recursive: true, force: true });
  fs.mkdirSync(dir, { recursive: true });
}
function esc(s = "") {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
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
function normalize(s = "") {
  return String(s)
    .toLowerCase()
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .replace(/đ/g, "d")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}
function achievementKey(row) {
  const recordId = String(row[1] ?? "").trim();
  const gabId = String(row[2] ?? "").trim();
  const name = normalize(row[3] ?? "");
  const time = normalize(row[4] ?? "");
  const title = normalize(row[5] ?? "");
  const check = normalize(row[8] ?? "");
  const url = normalize(row[6] ?? "");
  return [recordId, gabId, name, time, title, check, url].join("|");
}

if (!fs.existsSync(inputXlsx)) throw new Error(`Missing ${inputXlsx}`);
resetDir(workDir);
fs.copyFileSync(inputXlsx, zipPath);
ps(`Expand-Archive -LiteralPath '${zipPath.replace(/'/g, "''")}' -DestinationPath '${workDir.replace(/'/g, "''")}' -Force`);

const shared = sharedStrings(workDir);
const xml = fs.readFileSync(path.join(workDir, "xl", "worksheets", "sheet1.xml"), "utf8");
const parsedRows = [];
for (const row of xml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)) {
  const values = [];
  for (const cell of row[2].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
    const ref = cell[1].match(/\br="([^"]+)"/)?.[1];
    if (!ref) continue;
    values[colIndex(ref) - 1] = cellValue(cell[2], cell[1], shared);
  }
  parsedRows.push(values);
}

const headers = parsedRows[0] ?? [];
const seen = new Set();
const kept = [];
const dupes = [];
for (const row of parsedRows.slice(1)) {
  const key = achievementKey(row);
  if (seen.has(key)) dupes.push(row);
  else {
    seen.add(key);
    kept.push(row);
  }
}

const data = [headers, ...kept];
resetDir(outDir);
fs.mkdirSync(path.join(outDir, "_rels"), { recursive: true });
fs.mkdirSync(path.join(outDir, "xl", "_rels"), { recursive: true });
fs.mkdirSync(path.join(outDir, "xl", "worksheets"), { recursive: true });
fs.mkdirSync(path.join(outDir, "docProps"), { recursive: true });

let sheetData = "";
data.forEach((row, i) => {
  const r = i + 1;
  let cells = "";
  for (let j = 0; j < 13; j++) {
    const ref = `${colLetters(j + 1)}${r}`;
    cells += `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${esc(row[j] ?? "")}</t></is></c>`;
  }
  sheetData += `<row r="${r}">${cells}</row>`;
});

fs.writeFileSync(path.join(outDir, "[Content_Types].xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`);
fs.writeFileSync(path.join(outDir, "_rels", ".rels"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
fs.writeFileSync(path.join(outDir, "xl", "workbook.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Uu tien hoi ngo dedup" sheetId="1" r:id="rId1"/></sheets></workbook>`);
fs.writeFileSync(path.join(outDir, "xl", "_rels", "workbook.xml.rels"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
fs.writeFileSync(path.join(outDir, "xl", "styles.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9EAD3"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>`);
fs.writeFileSync(path.join(outDir, "xl", "worksheets", "sheet1.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:M${data.length}"/><cols><col min="1" max="1" width="16" customWidth="1"/><col min="2" max="3" width="36" customWidth="1"/><col min="4" max="4" width="28" customWidth="1"/><col min="5" max="5" width="14" customWidth="1"/><col min="6" max="6" width="60" customWidth="1"/><col min="7" max="8" width="60" customWidth="1"/><col min="9" max="10" width="42" customWidth="1"/><col min="11" max="13" width="24" customWidth="1"/></cols><sheetData>${sheetData}</sheetData><autoFilter ref="A1:M${data.length}"/></worksheet>`);
fs.writeFileSync(path.join(outDir, "docProps", "core.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Uu tien hoi ngo dedup</dc:title></cp:coreProperties>`);
fs.writeFileSync(path.join(outDir, "docProps", "app.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Codex</Application></Properties>`);

if (fs.existsSync(outZip)) fs.rmSync(outZip, { force: true });
execFileSync("powershell", ["-NoProfile", "-Command", `Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::CreateFromDirectory('${outDir.replace(/'/g, "''")}', '${outZip.replace(/'/g, "''")}')`]);
fs.copyFileSync(outZip, outputXlsx);

console.log(JSON.stringify({
  inputRows: parsedRows.length - 1,
  keptRows: kept.length,
  duplicateRowsRemoved: dupes.length,
  outputXlsx,
}, null, 2));
