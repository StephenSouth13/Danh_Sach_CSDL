import fs from "node:fs";
import path from "node:path";

const [dir, sheetNamePrefix = "Data_record", start = "930", end = "1170"] = process.argv.slice(2);
if (!dir) throw new Error("Usage: node inspect_row_range.mjs <xlsx-expanded-dir> [sheet-prefix] [start] [end]");

const read = (p) => fs.readFileSync(p, "utf8");
const decodeXml = (s = "") => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
function colIndex(ref) {
  let n = 0;
  for (const ch of (String(ref).match(/^[A-Z]+/)?.[0] || "")) n = n * 26 + ch.charCodeAt(0) - 64;
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
    sheets.push({ name: decodeXml(m[1]), target });
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
  if (t === "s") return shared[Number(cellXml.match(/<v>([\s\S]*?)<\/v>/)?.[1])] ?? "";
  if (t === "inlineStr") return [...cellXml.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((m) => decodeXml(m[1])).join("");
  return decodeXml(cellXml.match(/<v>([\s\S]*?)<\/v>/)?.[1] || "");
}

const sheet = sheetMap(dir).find((s) => s.name.startsWith(sheetNamePrefix));
const shared = sharedStrings(dir);
const xml = read(path.join(dir, sheet.target));
const s = Number(start), e = Number(end);
for (const row of xml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)) {
  const rowNumber = Number(row[1]);
  if (rowNumber < s || rowNumber > e) continue;
  const cells = [];
  let hasValue = false;
  for (const cell of row[2].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>|<c\b([^>]*)\/>/g)) {
    const attrs = cell[1] || cell[3] || "";
    const ref = attrs.match(/\br="([^"]+)"/)?.[1];
    if (!ref) continue;
    const value = cell[2] ? cellValue(cell[2], attrs, shared) : "";
    if (value.trim()) hasValue = true;
    if (colIndex(ref) <= 8) cells.push(`${ref}:${value.slice(0, 60)}`);
  }
  console.log(`${rowNumber}\t${hasValue ? "HAS_VALUE" : "blank"}\t${cells.join(" | ")}`);
}
