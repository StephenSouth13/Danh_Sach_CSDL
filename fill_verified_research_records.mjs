import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const root = process.cwd();
const work = path.join(root, "outputs", "complete_standard_work");
const expanded = path.join(work, "xlsx");
const sheetPath = path.join(expanded, "xl", "worksheets", "sheet2.xml");
const source = path.join(root, "Dulieutruyxuat", "So sánh.xlsx");
const zip = path.join(work, "verified_research_updates.zip");

const updates = [{
  row: 356,
  expectedName: "LÊ THỊ ÚT",
  values: {
    D: "Năm 2018 (chưa rõ ngày)",
    E: "Người vẽ tranh bằng nhiều chất liệu độc đáo nhất",
    K: "KỶ LỤC GIA VIỆT NAM",
    M: "2018",
    N: "Xác lập Kỷ lục Việt Nam “Người vẽ tranh bằng nhiều chất liệu độc đáo nhất”",
    O: "https://kyluc.vn/tin-tuc/ky-luc-viet-nam/vietkings-ky-luc-gia-9x-ve-tranh-bang-ca-phe-va-tuong-ot",
    S: "Đã xác minh Lê Thị Út (nghệ danh Kim Út) và thành tựu từ Kyluc.vn; nguồn công khai chỉ nêu năm 2018, chưa nêu ngày trao cụ thể.",
  },
}];

const escapeXml = (value = "") => String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;")
  .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const setCell = (rowXml, row, column, value) => {
  const ref = `${column}${row}`;
  const pattern = new RegExp(`<c\\b[^>]*\\br="${ref}"[^>]*>[\\s\\S]*?<\\/c>|<c\\b[^>]*\\br="${ref}"[^>]*/>`);
  const existing = rowXml.match(pattern)?.[0];
  const style = existing?.match(/\bs="(\d+)"/)?.[1] ?? "1";
  const replacement = `<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
  return existing ? rowXml.replace(pattern, replacement) : rowXml.replace("</row>", `${replacement}</row>`);
};

let xml = fs.readFileSync(sheetPath, "utf8");
for (const update of updates) {
  const pattern = new RegExp(`<row\\b[^>]*\\br="${update.row}"[^>]*>[\\s\\S]*?<\\/row>`);
  const original = xml.match(pattern)?.[0];
  if (!original || !original.includes(update.expectedName)) throw new Error(`Sai dòng hoặc không tìm thấy ${update.expectedName}`);
  let changed = original;
  for (const [column, value] of Object.entries(update.values)) changed = setCell(changed, update.row, column, value);
  xml = xml.replace(pattern, changed);
}
fs.writeFileSync(sheetPath, xml, "utf8");

if (fs.existsSync(zip)) fs.rmSync(zip, { force: true });
execFileSync("powershell", ["-NoProfile", "-Command",
  `Add-Type -AssemblyName System.IO.Compression.FileSystem; [IO.Compression.ZipFile]::CreateFromDirectory('${expanded.replace(/'/g, "''")}','${zip.replace(/'/g, "''")}')`,
]);
fs.copyFileSync(zip, source);
console.log(JSON.stringify({ updated: updates.map((item) => item.expectedName), output: source }, null, 2));
