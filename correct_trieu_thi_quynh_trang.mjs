import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const root = process.cwd();
const work = path.join(root, "outputs", "complete_standard_work");
const expanded = path.join(work, "xlsx");
const sheetPath = path.join(expanded, "xl", "worksheets", "sheet2.xml");
const workbookPath = path.join(root, "Dulieutruyxuat", "So sánh.xlsx");
const zipPath = path.join(work, "corrected_trieu_thi_quynh_trang.zip");

const targetRow = 349;
const gabUrl = "https://gab.world/vi/bank/nOBW6Q2RveR4n5UvdLMSsU84dKm2";
const achievementDate = "25/07/2020";
const achievement = "Ca sĩ ra mắt bộ DVD nhạc trữ tình nhiều nhất trong cùng một ngày";
const status = "KỶ LỤC GIA VIỆT NAM";
const sourceUrl = "https://kyluc.vn/tin-tuc/ky-luc/ca-si-trieu-trang-xac-lap-ky-luc-viet-nam-voi-bo-album-nhac-thanh-xuan-dung-so-sanh-em-voi-ai";
const note = "Đã xác minh link GAB do người dùng cung cấp và thành tựu từ bài chính thức của Kyluc.vn.";

const escapeXml = (value) => String(value)
  .replace(/&/g, "&amp;")
  .replace(/</g, "&lt;")
  .replace(/>/g, "&gt;")
  .replace(/"/g, "&quot;");

const setInlineCell = (rowXml, column, value) => {
  const ref = `${column}${targetRow}`;
  const cellPattern = new RegExp(`<c\\b[^>]*\\br="${ref}"[^>]*>[\\s\\S]*?<\\/c>`);
  const existing = rowXml.match(cellPattern)?.[0];
  const style = existing?.match(/\\bs="(\\d+)"/)?.[1] ?? "1";
  const replacement = `<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
  if (existing) return rowXml.replace(cellPattern, replacement);
  return rowXml.replace("</row>", `${replacement}</row>`);
};

let xml = fs.readFileSync(sheetPath, "utf8");
const rowPattern = new RegExp(`<row\\b[^>]*\\br="${targetRow}"[^>]*>[\\s\\S]*?<\\/row>`);
const originalRow = xml.match(rowPattern)?.[0];
if (!originalRow || !originalRow.includes("TRIỆU THỊ QUỲNH TRANG")) {
  throw new Error(`Không tìm thấy TRIỆU THỊ QUỲNH TRANG tại dòng ${targetRow}.`);
}

let updatedRow = setInlineCell(originalRow, "F", gabUrl);
updatedRow = setInlineCell(updatedRow, "D", achievementDate);
updatedRow = setInlineCell(updatedRow, "E", achievement);
updatedRow = setInlineCell(updatedRow, "K", status);
updatedRow = setInlineCell(updatedRow, "M", achievementDate);
updatedRow = setInlineCell(updatedRow, "N", `Xác lập Kỷ lục Việt Nam “${achievement}”`);
updatedRow = setInlineCell(updatedRow, "O", sourceUrl);
updatedRow = setInlineCell(updatedRow, "S", note);
xml = xml.replace(rowPattern, updatedRow);
fs.writeFileSync(sheetPath, xml, "utf8");

if (fs.existsSync(zipPath)) fs.rmSync(zipPath, { force: true });
const sourceEscaped = expanded.replace(/'/g, "''");
const zipEscaped = zipPath.replace(/'/g, "''");
execFileSync("powershell", [
  "-NoProfile",
  "-Command",
  `Add-Type -AssemblyName System.IO.Compression.FileSystem; [IO.Compression.ZipFile]::CreateFromDirectory('${sourceEscaped}','${zipEscaped}')`,
]);
fs.copyFileSync(zipPath, workbookPath);

console.log(JSON.stringify({ targetRow, gabUrl, achievementDate, achievement, status, sourceUrl, workbookPath }, null, 2));
