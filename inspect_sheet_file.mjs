import fs from "node:fs";
import path from "node:path";

const [dir, sheetFile = "sheet1.xml", start = "1", end = "20", maxCol = "12"] = process.argv.slice(2);
if (!dir) throw new Error("Usage: node inspect_sheet_file.mjs <xlsx-expanded-dir> <sheet-file> <start> <end> [max-col]");

function decodeXml(s = "") {
  return s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
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
  return [...xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)]
    .map((si) => [...si[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((m) => decodeXml(m[1])).join(""));
}
function cellValue(cellXml, attrs, shared) {
  const t = attrs.match(/\bt="([^"]+)"/)?.[1] || "";
  if (t === "s") return shared[Number(cellXml.match(/<v>([\s\S]*?)<\/v>/)?.[1])] ?? "";
  if (t === "inlineStr") return [...cellXml.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((m) => decodeXml(m[1])).join("");
  return decodeXml(cellXml.match(/<v>([\s\S]*?)<\/v>/)?.[1] || "");
}

const xml = fs.readFileSync(path.join(dir, "xl", "worksheets", sheetFile), "utf8");
const shared = sharedStrings(dir);
const s = Number(start), e = Number(end), mc = Number(maxCol);
let nonEmpty = 0;
for (const row of xml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)) {
  const rowNumber = Number(row[1]);
  let has = false;
  const cells = [];
  for (const cell of row[2].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>|<c\b([^>]*)\/>/g)) {
    const attrs = cell[1] || cell[3] || "";
    const ref = attrs.match(/\br="([^"]+)"/)?.[1];
    if (!ref) continue;
    const value = cell[2] ? cellValue(cell[2], attrs, shared) : "";
    if (value.trim()) has = true;
    if (rowNumber >= s && rowNumber <= e && colIndex(ref) <= mc) cells.push(`${ref}:${value.slice(0, 80)}`);
  }
  if (has) nonEmpty++;
  if (rowNumber >= s && rowNumber <= e) console.log(`${rowNumber}\t${has ? "HAS_VALUE" : "blank"}\t${cells.join(" | ")}`);
}
console.error(JSON.stringify({ sheetFile, nonEmpty }));
