import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const root = process.cwd();
const source = path.join(root, "Dulieutruyxuat", "So sánh.xlsx");
const work = path.join(root, "outputs", "verify_current_compare");
const xlsx = path.join(work, "xlsx");

function read(file) { return fs.readFileSync(file, "utf8"); }
function write(file, value) { fs.writeFileSync(file, value); }
function esc(value = "") { return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
function dec(value = "") { return String(value).replace(/&#x([0-9a-f]+);/gi, (_, x) => String.fromCodePoint(parseInt(x, 16))).replace(/&#(\d+);/g, (_, x) => String.fromCodePoint(Number(x))).replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&"); }
function colNo(ref) { let n = 0; for (const c of (String(ref).match(/^[A-Z]+/)?.[0] || "")) n = n * 26 + c.charCodeAt(0) - 64; return n; }
function colName(n) { let s = ""; while (n) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
function sharedStrings() {
  const file = path.join(xlsx, "xl", "sharedStrings.xml");
  if (!fs.existsSync(file)) return [];
  return [...read(file).matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map((m) => [...m[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((x) => dec(x[1])).join(""));
}
function rows(sheetFile, strings) {
  const result = [];
  for (const m of read(path.join(xlsx, "xl", "worksheets", sheetFile)).matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>/g)) {
    const row = Number(m[1].match(/\br="(\d+)"/)?.[1]); const values = new Map();
    for (const c of m[2].matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const attrs = c[1] || ""; const ref = attrs.match(/\br="([^"]+)"/)?.[1]; if (!ref) continue;
      const type = attrs.match(/\bt="([^"]+)"/)?.[1] || ""; const body = c[2] || ""; let value = "";
      if (type === "s") value = strings[Number(body.match(/<v>([\s\S]*?)<\/v>/)?.[1])] || "";
      else if (type === "inlineStr") value = [...body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((x) => dec(x[1])).join("");
      else value = dec(body.match(/<v>([\s\S]*?)<\/v>/)?.[1] || "");
      values.set(colNo(ref), value);
    }
    result.push({ row, values });
  }
  return result;
}
function norm(value = "") {
  return String(value).toLowerCase().normalize("NFD").replace(/\p{Diacritic}/gu, "").replace(/đ/g, "d")
    .replace(/\b(ky luc gia|klg|ong|ba|anh|chi|gs|pgs|ts|ths|bs|nghe nhan|nghe si|hoa si|nha suu tap|anh hung lao dong|giao su|bac si|tien si|thac si|phap danh|chau a|viet nam|the gioi)\b/g, " ")
    .replace(/[^a-z0-9]+/g, " ").replace(/\s+/g, " ").trim();
}
function gabKey(value = "") { return String(value).match(/gab\.world\/vi\/bank\/([^/?#]+)/i)?.[1]?.toLowerCase() || ""; }
function nameVariants(value = "") {
  const raw = String(value).trim(); const result = new Set([norm(raw)]);
  for (const m of raw.matchAll(/\(([^)]+)\)/g)) result.add(norm(m[1]));
  result.add(norm(raw.replace(/\([^)]*\)/g, " ")));
  return [...result].filter(Boolean);
}
function tokenScore(a, b) {
  const x = new Set(norm(a).split(" ").filter(Boolean)), y = new Set(norm(b).split(" ").filter(Boolean));
  if (!x.size || !y.size) return 0; let hit = 0; for (const t of x) if (y.has(t)) hit++;
  return (2 * hit) / (x.size + y.size);
}
function levenshtein(a, b) {
  const x = norm(a), y = norm(b); const row = Array.from({length:y.length+1},(_,i)=>i);
  for (let i=1;i<=x.length;i++) { let prev=row[0]; row[0]=i; for (let j=1;j<=y.length;j++) { const old=row[j]; row[j]=Math.min(row[j]+1,row[j-1]+1,prev+(x[i-1]===y[j-1]?0:1)); prev=old; } }
  return row[y.length];
}
function similarity(a,b) {
  const x=norm(a).split(" ").filter(Boolean), y=norm(b).split(" ").filter(Boolean);
  const token=tokenScore(a,b);
  if (x.length!==y.length || x.length<2) return token;
  let changed=0;
  for(let i=0;i<x.length;i++) if(x[i]!==y[i]) { if(levenshtein(x[i],y[i])>1) return token; changed++; }
  return changed===1 ? Math.max(token,0.9) : token;
}

const strings = sharedStrings();
const list560 = rows("sheet1.xml", strings).filter((r)=>r.row>1).map((r)=>({ row:r.row, name:String(r.values.get(4)||"").trim(), link:String(r.values.get(3)||"").trim(), gab:gabKey(r.values.get(3)||""), variants:nameVariants(r.values.get(4)||"") })).filter((r)=>r.name);
const raw200 = rows("sheet3.xml", strings).filter((r)=>r.row>2).map((r)=>({ row:r.row, name:String(r.values.get(3)||"").trim(), link:String(r.values.get(2)||"").trim(), gab:gabKey(r.values.get(2)||""), variants:nameVariants(r.values.get(3)||"") })).filter((r)=>r.name);

const peopleMap = new Map();
for (const r of raw200) {
  const key = r.gab || r.variants[0];
  if (!peopleMap.has(key)) peopleMap.set(key, { ...r, sourceRows:[] });
  peopleMap.get(key).sourceRows.push(r.row);
}
const people = [...peopleMap.values()];
const results = [];
for (const p of people) {
  let match = p.gab ? list560.find((r)=>r.gab===p.gab) : null;
  let status = "", basis = "", note = "", bestScore = 0;
  if (match) { status="CÓ TRONG DANH SÁCH 560"; basis="Trùng chính xác link GAB"; note="Đúng cùng một hồ sơ GAB."; }
  if (!match) {
    match = list560.find((r)=>p.variants.some((a)=>r.variants.includes(a)));
    if (match) { status="CÓ TRONG DANH SÁCH 560"; basis="Trùng chính xác tên đã chuẩn hóa"; note=p.gab&&match.gab&&p.gab!==match.gab?"Tên trùng nhưng link GAB khác; cần xác nhận hồ sơ đúng.":"Tên trùng sau khi bỏ hoa/thường, dấu và danh xưng."; }
  }
  if (!match) {
    let best=null;
    for (const r of list560) for (const a of p.variants) for (const b of r.variants) { const s=similarity(a,b); if (!best||s>best.score) best={r,score:s}; }
    bestScore=best?.score||0;
    if (bestScore>=0.78) { match=best.r; status="CẦN KIỂM TRA"; basis=`Tên gần giống ${Math.round(bestScore*100)}%`; note="Không tự kết luận là cùng người; cần kiểm tra thủ công."; }
  }
  if (!match) { status="KHÔNG TÌM THẤY TRONG 560"; basis="Không có link hoặc tên khớp đủ chắc"; note="Chưa thấy người tương ứng trong danh sách DEV xuất về."; }
  results.push({ ...p, status, matchName:match?.name||"", matchLink:match?.link||"", matchRow:match?.row||"", basis, note });
}
results.sort((a,b)=>a.status.localeCompare(b.status,"vi")||a.name.localeCompare(b.name,"vi"));
const counts = { yes:results.filter((r)=>r.status.startsWith("CÓ ")).length, review:results.filter((r)=>r.status==="CẦN KIỂM TRA").length, no:results.filter((r)=>r.status.startsWith("KHÔNG")).length };
const displayResults = results.filter((r)=>r.status.startsWith("CÓ ")).sort((a,b)=>a.name.localeCompare(b.name,"vi"));

let styles = read(path.join(xlsx,"xl","styles.xml"));
function appendCollection(xml, tag, item) {
  const re = new RegExp(`<${tag}\\b([^>]*)count="(\\d+)"([^>]*)>([\\s\\S]*?)<\\/${tag}>`); const m=xml.match(re); if(!m) throw new Error(`Không tìm thấy ${tag}`);
  return xml.replace(re, `<${tag}${m[1]}count="${Number(m[2])+1}"${m[3]}>${m[4]}${item}</${tag}>`);
}
const fontCount=Number(styles.match(/<fonts count="(\d+)"/)?.[1]||0);
styles=appendCollection(styles,"fonts",`<font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Arial"/></font>`);
const fillsStart=Number(styles.match(/<fills count="(\d+)"/)?.[1]||0);
for (const color of ["FF17365D","FFD9EAD3","FFFFF2CC","FFF4CCCC","FFD9E2F3"]) styles=appendCollection(styles,"fills",`<fill><patternFill patternType="solid"><fgColor rgb="${color}"/><bgColor indexed="64"/></patternFill></fill>`);
const xfsStart=Number(styles.match(/<cellXfs count="(\d+)"/)?.[1]||0);
const xfs=[
  `<xf numFmtId="0" fontId="${fontCount}" fillId="${fillsStart}" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>`,
  `<xf numFmtId="0" fontId="0" fillId="${fillsStart+4}" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>`,
  `<xf numFmtId="0" fontId="0" fillId="${fillsStart+1}" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>`,
  `<xf numFmtId="0" fontId="0" fillId="${fillsStart+2}" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>`,
  `<xf numFmtId="0" fontId="0" fillId="${fillsStart+3}" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>`,
  `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>`
];
for (const xf of xfs) styles=appendCollection(styles,"cellXfs",xf);
write(path.join(xlsx,"xl","styles.xml"),styles);

const cell=(v,ref,s)=>`<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
const row=(values,n,style,height=34)=>`<row r="${n}" ht="${height}" customHeight="1">${values.map((v,i)=>cell(v,`${colName(i+1)}${n}`,Array.isArray(style)?style[i]:style)).join("")}</row>`;
const table=[];
table.push(row(["KẾT QUẢ ĐỐI CHIẾU KLG BAN ĐẦU VỚI DANH SÁCH 560"],1,xfsStart,28));
table.push(row([`Tab KLG ban đầu có ${results.length} người; tìm thấy ${counts.yes} người trùng trong danh sách 560.`],2,xfsStart+1,24));
table.push(row(["Bảng này chỉ liệt kê các trường hợp khớp chắc bằng link GAB hoặc tên sau khi chuẩn hóa."],3,xfsStart+1,24));
table.push(`<row r="4" ht="8" customHeight="1"/>`);
const headers=["STT","TÊN TRONG KLG BAN ĐẦU","KẾT QUẢ","TÊN TƯƠNG ỨNG TRONG 560","LINK GAB BAN ĐẦU","LINK GAB TRONG 560","CĂN CỨ ĐỐI CHIẾU","GHI CHÚ"];
table.push(row(headers,5,xfsStart,42));
displayResults.forEach((r,i)=>{
  const statusStyle=xfsStart+2;
  table.push(row([i+1,r.name,r.status,r.matchName,r.link,r.matchLink,r.basis,r.note],i+6,[xfsStart+5,xfsStart+5,statusStyle,xfsStart+5,xfsStart+5,xfsStart+5,xfsStart+5,xfsStart+5],40));
});
const last=displayResults.length+5, widths=[7,30,28,34,48,48,31,52];
const sheet=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><dimension ref="A1:H${last}"/><cols>${widths.map((w,i)=>`<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`).join("")}</cols><sheetData>${table.join("")}</sheetData><autoFilter ref="A5:H${last}"/><mergeCells count="3"><mergeCell ref="A1:H1"/><mergeCell ref="A2:H2"/><mergeCell ref="A3:H3"/></mergeCells></worksheet>`;
write(path.join(xlsx,"xl","worksheets","sheet4.xml"),sheet);

let wb=read(path.join(xlsx,"xl","workbook.xml"));
wb=wb.replace(/<sheet[^>]*name="KẾT QUẢ 200 vs 560"[^>]*\/>/g,"").replace(/<sheet[^>]*name="KẾT QUẢ BAN ĐẦU vs 560"[^>]*\/>/g,"").replace(/<sheet[^>]*name="TRÙNG BAN ĐẦU - 560"[^>]*\/>/g,"");
wb=wb.replace("</sheets>",`<sheet sheetId="4" name="TRÙNG BAN ĐẦU - 560" state="visible" r:id="rId7"/></sheets>`);
write(path.join(xlsx,"xl","workbook.xml"),wb);
let rel=read(path.join(xlsx,"xl","_rels","workbook.xml.rels"));
rel=rel.replace(/<Relationship[^>]*Id="rId7"[^>]*\/>/g,"").replace("</Relationships>",`<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet4.xml"/></Relationships>`);
write(path.join(xlsx,"xl","_rels","workbook.xml.rels"),rel);
let types=read(path.join(xlsx,"[Content_Types].xml"));
if(!types.includes('/xl/worksheets/sheet4.xml')) types=types.replace("</Types>",`<Override PartName="/xl/worksheets/sheet4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`);
write(path.join(xlsx,"[Content_Types].xml"),types);

const zip=path.join(work,"updated.zip"), temp=path.join(work,"So sánh.updated.xlsx");
if(fs.existsSync(zip)) fs.rmSync(zip,{force:true}); if(fs.existsSync(temp)) fs.rmSync(temp,{force:true});
execFileSync("powershell",["-NoProfile","-Command",`Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::CreateFromDirectory('${xlsx.replace(/'/g,"''")}', '${zip.replace(/'/g,"''")}')`],{stdio:"pipe"});
fs.copyFileSync(zip,temp); fs.copyFileSync(temp,source);
console.log(JSON.stringify({source, sharedStrings:strings.length, sample2626:strings[2626]||"", rows560:list560.length, rawRows200:raw200.length, uniquePeople200:results.length, ...counts},null,2));
