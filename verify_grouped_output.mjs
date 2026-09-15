import fs from "node:fs";
import path from "node:path";

const dir = "outputs\\reconcile_grouped_check_notes_work\\target";
const sheetPath = path.join(dir, "xl", "worksheets", "sheet1.xml");
const sharedPath = path.join(dir, "xl", "sharedStrings.xml");
const xml = fs.readFileSync(sheetPath, "utf8");
let shared = [];
if (fs.existsSync(sharedPath)) {
  const ss = fs.readFileSync(sharedPath, "utf8");
  shared = [...ss.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map((si) =>
    [...si[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)]
      .map((t) => t[1].replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&"))
      .join("")
  );
}
function val(cell, attrs) {
  const t = attrs.match(/\bt="([^"]+)"/)?.[1] || "";
  if (t === "s") return shared[Number(cell.match(/<v>([\s\S]*?)<\/v>/)?.[1])] || "";
  if (t === "inlineStr") return [...cell.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((m) => m[1].replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&")).join("");
  return cell.match(/<v>([\s\S]*?)<\/v>/)?.[1] || "";
}
function rowValues(rowNo) {
  const row = xml.match(new RegExp(`<row\\b[^>]*\\br="${rowNo}"[^>]*>([\\s\\S]*?)<\\/row>`))?.[1] || "";
  const out = {};
  for (const m of row.matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
    const ref = m[1].match(/\br="([^"]+)"/)?.[1];
    if (ref && /^[ABCHIJKL]\d+$/.test(ref)) out[ref] = val(m[2], m[1]).slice(0, 500);
  }
  return out;
}

let blankRecordIdMissingRows = 0;
let badMissingRows = 0;
for (const row of xml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)) {
  const rowNo = Number(row[1]);
  if (rowNo <= 1) continue;
  const values = {};
  for (const m of row[2].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
    const ref = m[1].match(/\br="([^"]+)"/)?.[1];
    if (ref && /^[AH]\d+$/.test(ref)) values[ref[0]] = val(m[2], m[1]);
  }
  if (values.H === "GAB thiếu thành tựu") {
    if (!values.A) blankRecordIdMissingRows++;
    else badMissingRows++;
  }
}

console.log(JSON.stringify({
  dimension: xml.match(/<dimension[^>]+>/)?.[0],
  mergeCountI: (xml.match(/<mergeCell ref="I\d+:I\d+"/g) || []).length,
  hasTokenWord: /token/i.test(xml),
  blankRecordIdMissingRows,
  badMissingRows,
  rows: [1, 2, 3, 4, 5, 6, 7, 8, 939, 940].map((n) => ({ row: n, values: rowValues(n) })),
}, null, 2));
