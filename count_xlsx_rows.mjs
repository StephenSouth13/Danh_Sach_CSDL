import fs from "node:fs";
import path from "node:path";

const [dir] = process.argv.slice(2);
if (!dir) throw new Error("Usage: node count_xlsx_rows.mjs <xlsx-expanded-dir>");

function read(p) {
  return fs.readFileSync(p, "utf8");
}

function decodeXml(s = "") {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
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
  for (const m of rels.matchAll(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)) relMap.set(m[1], m[2]);
  const sheets = [];
  for (const m of wb.matchAll(/<sheet\b[^>]*name="([^"]+)"[^>]*sheetId="([^"]+)"[^>]*(?:r:id|id)="([^"]+)"/g)) {
    let target = relMap.get(m[3]);
    if (!target) continue;
    if (!target.startsWith("xl/")) target = `xl/${target}`;
    sheets.push({ name: decodeXml(m[1]), id: m[2], target });
  }
  return sheets;
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

function cellValue(cellXml, attrs, shared) {
  const t = attrs.match(/\bt="([^"]+)"/)?.[1] || "";
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

const shared = sharedStrings(dir);
for (const sheet of sheetMap(dir)) {
  const xml = read(path.join(dir, sheet.target));
  let rowCount = 0;
  let nonEmptyRows = 0;
  let firstDataRow = null;
  let lastDataRow = null;
  const headers = new Map();
  for (const row of xml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)) {
    rowCount++;
    const rowNumber = Number(row[1]);
    let nonEmpty = false;
    for (const cell of row[2].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>|<c\b([^>]*)\/>/g)) {
      const attrs = cell[1] || cell[3] || "";
      const ref = attrs.match(/\br="([^"]+)"/)?.[1];
      if (!ref) continue;
      const value = cell[2] ? cellValue(cell[2], attrs, shared) : "";
      if (rowNumber === 1 && value) headers.set(colIndex(ref), value);
      if (value.trim()) nonEmpty = true;
    }
    if (nonEmpty) {
      nonEmptyRows++;
      if (firstDataRow === null) firstDataRow = rowNumber;
      lastDataRow = rowNumber;
    }
  }
  console.log(JSON.stringify({
    sheet: sheet.name,
    xmlRows: rowCount,
    nonEmptyRows,
    dataRowsAfterHeader: Math.max(0, nonEmptyRows - 1),
    firstNonEmptyRow: firstDataRow,
    lastNonEmptyRow: lastDataRow,
    headers: Object.fromEntries([...headers].slice(0, 10)),
  }));
}
