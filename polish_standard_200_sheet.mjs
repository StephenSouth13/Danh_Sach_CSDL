import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const root = process.cwd();
const work = path.join(root, "outputs", "complete_standard_work");
const expanded = path.join(work, "xlsx");
const sheetPath = path.join(expanded, "xl", "worksheets", "sheet2.xml");
const stylesPath = path.join(expanded, "xl", "styles.xml");
const source = path.join(root, "Dulieutruyxuat", "So sánh.xlsx");
const zip = path.join(work, "polished_standard_200.zip");

let styles = fs.readFileSync(stylesPath, "utf8");
const marker = '<!-- polished-200-styles -->';
if (!styles.includes(marker)) {
  const additions = [
    '<xf numFmtId="0" fontId="24" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>',
    '<xf numFmtId="0" fontId="27" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>',
    '<xf numFmtId="0" fontId="19" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>',
    '<xf numFmtId="0" fontId="27" fillId="12" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>',
  ].join("");
  styles = styles.replace(/<cellXfs count="(\d+)">/, (_, count) => `<cellXfs count="${Number(count) + 4}">${marker}`)
    .replace('</cellXfs>', `${additions}</cellXfs>`);
  fs.writeFileSync(stylesPath, styles, "utf8");
}

const styleMatch = styles.match(/<cellXfs count="(\d+)">/);
const styleCount = Number(styleMatch?.[1] ?? 0);
const headerStyle = styleCount - 4;
const bodyStyle = styleCount - 3;
const linkStyle = styleCount - 2;
const warningStyle = styleCount - 1;

const cellColumn = (attrs) => attrs.match(/\br="([A-Z]+)\d+"/)?.[1] ?? "";
const isBeyondT = (column) => column.length > 1 || column > "T";
let xml = fs.readFileSync(sheetPath, "utf8");

xml = xml.replace(/<sheetViews>[\s\S]*?<\/sheetViews>/, "")
  .replace(/(<worksheet\b[^>]*>)/, '$1<sheetViews><sheetView workbookViewId="0"><pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="G3" sqref="G3"/></sheetView></sheetViews>')
  .replace(/<cols>[\s\S]*?<\/cols>/, '<cols><col min="1" max="1" width="7" customWidth="1"/><col min="2" max="2" width="18" customWidth="1"/><col min="3" max="4" width="15" customWidth="1"/><col min="5" max="5" width="42" customWidth="1"/><col min="6" max="6" width="38" customWidth="1"/><col min="7" max="7" width="28" customWidth="1"/><col min="8" max="8" width="11" customWidth="1"/><col min="9" max="9" width="20" customWidth="1"/><col min="10" max="10" width="16" customWidth="1"/><col min="11" max="12" width="24" customWidth="1"/><col min="13" max="13" width="15" customWidth="1"/><col min="14" max="14" width="48" customWidth="1"/><col min="15" max="15" width="44" customWidth="1"/><col min="16" max="18" width="18" customWidth="1"/><col min="19" max="19" width="46" customWidth="1"/><col min="20" max="20" width="16" customWidth="1"/></cols>')
  .replace(/<mergeCells[^>]*>[\s\S]*?<\/mergeCells>/, '<mergeCells count="9"><mergeCell ref="F1:J1"/><mergeCell ref="R1:S1"/><mergeCell ref="A1:A2"/><mergeCell ref="B1:B2"/><mergeCell ref="C1:C2"/><mergeCell ref="D1:D2"/><mergeCell ref="E1:E2"/><mergeCell ref="O1:O2"/><mergeCell ref="P1:P2"/></mergeCells>')
  .replace(/<autoFilter[^>]*\/>/g, "")
  .replace(/<dimension ref="[^"]+"\/>/, '<dimension ref="A1:T373"/>');

xml = xml.replace(/<row\b([^>]*)>([\s\S]*?)<\/row>/g, (whole, attrs, content) => {
  const row = Number(attrs.match(/\br="(\d+)"/)?.[1] ?? 0);
  const unresolved = /CHƯA XÁC MINH/.test(content);
  const height = row <= 2 ? (row === 1 ? 32 : 58) : 48;
  const cleanedAttrs = attrs.replace(/\sht="[^"]*"/g, "").replace(/\scustomHeight="[^"]*"/g, "");
  const cells = content.replace(/<c\b([^>]*)\/>|<c\b([^>]*)>[\s\S]*?<\/c>/g, (cell, emptyAttrs, fullAttrs) => {
    const allAttrs = emptyAttrs || fullAttrs || "";
    const column = cellColumn(allAttrs);
    if (isBeyondT(column)) return "";
    const style = row <= 2 ? headerStyle : unresolved ? warningStyle : (["F", "O"].includes(column) ? linkStyle : bodyStyle);
    return cell.replace(/\bs="\d+"/, `s="${style}"`).replace(/(<c\b(?![^>]*\bs=)[^>]*)(>)/, `$1 s="${style}"$2`);
  });
  return `<row${cleanedAttrs} ht="${height}" customHeight="1">${cells}</row>`;
});

fs.writeFileSync(sheetPath, xml, "utf8");
if (fs.existsSync(zip)) fs.rmSync(zip, { force: true });
execFileSync("powershell", ["-NoProfile", "-Command",
  `Add-Type -AssemblyName System.IO.Compression.FileSystem; [IO.Compression.ZipFile]::CreateFromDirectory('${expanded.replace(/'/g, "''")}','${zip.replace(/'/g, "''")}')`,
]);
fs.copyFileSync(zip, source);
console.log(JSON.stringify({ headerStyle, bodyStyle, linkStyle, warningStyle, output: source }, null, 2));
