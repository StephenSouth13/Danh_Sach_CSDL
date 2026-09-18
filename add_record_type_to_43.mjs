import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const root = process.cwd();
const work = path.join(root, "outputs", "complete_standard_work");
const expanded = path.join(work, "xlsx");
const source = path.join(root, "Dulieutruyxuat", "So sánh.xlsx");
const zip = path.join(work, "with_record_type_43.zip");
const restoreZip = path.join(work, "corrected_trieu_thi_quynh_trang.zip");
const sheet2Path = path.join(expanded, "xl", "worksheets", "sheet2.xml");
const sheet5Path = path.join(expanded, "xl", "worksheets", "sheet5.xml");

// Restore the untouched comparison sheet before adding the genuinely new column.
execFileSync("powershell", ["-NoProfile", "-Command",
  `Add-Type -AssemblyName System.IO.Compression.FileSystem; $z=[IO.Compression.ZipFile]::OpenRead('${restoreZip.replace(/'/g, "''")}'); $e=$z.GetEntry('xl\\worksheets\\sheet5.xml'); $i=$e.Open(); $o=[IO.File]::Create('${sheet5Path.replace(/'/g, "''")}'); $i.CopyTo($o); $o.Close(); $i.Close(); $z.Dispose()`,
]);

const decode = (value = "") => String(value)
  .replace(/&#x([0-9a-f]+);/gi, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
  .replace(/&#(\d+);/g, (_, num) => String.fromCodePoint(Number(num)))
  .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"')
  .replace(/&apos;/g, "'").replace(/&amp;/g, "&");
const escapeXml = (value = "") => String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;")
  .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const norm = (value = "") => String(value).toLowerCase().normalize("NFD")
  .replace(/\p{Diacritic}/gu, "").replace(/[^a-z0-9]+/g, " ").trim();
const sharedXml = fs.readFileSync(path.join(expanded, "xl", "sharedStrings.xml"), "utf8");
const shared = [...sharedXml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map((item) =>
  [...item[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((text) => decode(text[1])).join(""),
);
const cellValue = (rowXml, column) => {
  const cell = rowXml.match(new RegExp(`<c\\b([^>]*)\\br="${column}\\d+"([^>]*)>([\\s\\S]*?)<\\/c>`));
  if (!cell) return "";
  const attrs = `${cell[1]} ${cell[2]}`;
  if (/\bt="s"/.test(attrs)) return shared[Number(cell[3].match(/<v>(\d+)<\/v>/)?.[1])] ?? "";
  return decode([...cell[3].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((m) => m[1]).join(""));
};
const rows = (xml) => [...xml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*\/>|<row\b[^>]*\br="(\d+)"[^>]*>[\s\S]*?<\/row>/g)]
  .map((match) => ({ 0: match[0], 1: match[1] || match[2] }));

const typeByCertificate = new Map();
for (const match of rows(fs.readFileSync(sheet2Path, "utf8"))) {
  const certificate = norm(cellValue(match[0], "C"));
  const type = cellValue(match[0], "K").trim();
  if (certificate && type && !type.includes("CHƯA")) typeByCertificate.set(certificate, type);
}

const inferType = (rowXml) => {
  const certificate = norm(cellValue(rowXml, "G"));
  const evidence = norm([cellValue(rowXml, "G"), cellValue(rowXml, "I"), cellValue(rowXml, "K"), cellValue(rowXml, "L")].join(" "));
  if (evidence.includes("ky luc the gioi") || evidence.includes("world record")) return "KỶ LỤC GIA THẾ GIỚI";
  if (evidence.includes("ky luc chau a") || evidence.includes("asian record")) return "KỶ LỤC GIA CHÂU Á";
  const mapped = typeByCertificate.get(certificate);
  const mappedNorm = norm(mapped);
  if (mappedNorm.includes("the gioi")) return "KỶ LỤC GIA THẾ GIỚI";
  if (mappedNorm.includes("chau a")) return "KỶ LỤC GIA CHÂU Á";
  return "KỶ LỤC GIA VIỆT NAM";
};

let xml = fs.readFileSync(sheet5Path, "utf8");
xml = xml.replace(/<mergeCell ref="A1:N1"\/>/g, '<mergeCell ref="A1:O1"/>')
  .replace(/<mergeCell ref="A2:N2"\/>/g, '<mergeCell ref="A2:O2"/>')
  .replace(/<mergeCell ref="A3:N3"\/>/g, '<mergeCell ref="A3:O3"/>')
  .replace(/<autoFilter ref="A5:N57"\/>/g, '<autoFilter ref="A5:O57"/>');

xml = xml.replace(/<row\b[^>]*\br="(\d+)"[^>]*\/>|<row\b[^>]*\br="(\d+)"[^>]*>[\s\S]*?<\/row>/g, (rowXml, selfClosingNumber, normalNumber) => {
  const rowNumberText = selfClosingNumber || normalNumber;
  const rowNumber = Number(rowNumberText);
  if (rowNumber < 5 || rowNumber > 57 || selfClosingNumber) return rowXml;
  const value = rowNumber === 5 ? "LOẠI KỶ LỤC" : inferType(rowXml);
  const nCell = rowXml.match(new RegExp(`<c\\b[^>]*\\br="N${rowNumber}"[^>]*>[\\s\\S]*?<\\/c>`))?.[0];
  const style = nCell?.match(/\bs="(\d+)"/)?.[1] ?? (rowNumber === 5 ? "64" : "66");
  const oCell = `<c r="O${rowNumber}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
  const existing = new RegExp(`<c\\b[^>]*\\br="O${rowNumber}"[^>]*>[\\s\\S]*?<\\/c>`);
  return existing.test(rowXml) ? rowXml.replace(existing, oCell) : rowXml.replace("</row>", `${oCell}</row>`);
});
xml = xml.replace(/<dimension ref="A1:[A-Z]+\d+"\/>/, '<dimension ref="A1:O57"/>');
fs.writeFileSync(sheet5Path, xml, "utf8");

if (fs.existsSync(zip)) fs.rmSync(zip, { force: true });
execFileSync("powershell", ["-NoProfile", "-Command",
  `Add-Type -AssemblyName System.IO.Compression.FileSystem; [IO.Compression.ZipFile]::CreateFromDirectory('${expanded.replace(/'/g, "''")}','${zip.replace(/'/g, "''")}')`,
]);
fs.copyFileSync(zip, source);
console.log(JSON.stringify({ rowsFilled: 52, newColumn: "O - LOẠI KỶ LỤC", output: source }, null, 2));
