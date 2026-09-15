import fs from "node:fs";
import path from "node:path";

const dir = "outputs/reconcile_grouped_check_notes_work/priority_source";
const wb = fs.readFileSync(path.join(dir, "xl/workbook.xml"), "utf8");
const rels = fs.readFileSync(path.join(dir, "xl/_rels/workbook.xml.rels"), "utf8");
const relMap = new Map([...rels.matchAll(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)].map((m) => [m[1], m[2]]));
for (const m of wb.matchAll(/<sheet\b[^>]*name="([^"]+)"[^>]*sheetId="([^"]+)"[^>]*(?:r:id|id)="([^"]+)"/g)) {
  console.log(JSON.stringify({ name: m[1], sheetId: m[2], rid: m[3], target: relMap.get(m[3]) }));
}
