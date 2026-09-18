import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const root=process.cwd();
const source=path.join(root,"Dulieutruyxuat","So sánh.xlsx");
const work=path.join(root,"outputs","pull_43_achievements_refresh");
const xlsx=path.join(work,"xlsx");
const read=(p)=>fs.readFileSync(p,"utf8");
const write=(p,v)=>fs.writeFileSync(p,v);
const esc=(v="")=>String(v).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
const dec=(v="")=>String(v).replace(/&#x([0-9a-f]+);/gi,(_,x)=>String.fromCodePoint(parseInt(x,16))).replace(/&#(\d+);/g,(_,x)=>String.fromCodePoint(Number(x))).replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&amp;/g,"&");
function colNo(ref){let n=0;for(const c of(String(ref).match(/^[A-Z]+/)?.[0]||""))n=n*26+c.charCodeAt(0)-64;return n;}
function colName(n){let s="";while(n){const r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26);}return s;}
function shared(){const p=path.join(xlsx,"xl","sharedStrings.xml");if(!fs.existsSync(p))return[];return[...read(p).matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map(m=>[...m[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(x=>dec(x[1])).join(""));}
function rows(file,strings){const out=[];for(const m of read(path.join(xlsx,"xl","worksheets",file)).matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>/g)){const row=Number(m[1].match(/\br="(\d+)"/)?.[1]);const values=new Map();for(const c of m[2].matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)){const attrs=c[1]||"",ref=attrs.match(/\br="([^"]+)"/)?.[1];if(!ref)continue;const type=attrs.match(/\bt="([^"]+)"/)?.[1]||"",body=c[2]||"";let value="";if(type==="s")value=strings[Number(body.match(/<v>([\s\S]*?)<\/v>/)?.[1])]||"";else if(type==="inlineStr")value=[...body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(x=>dec(x[1])).join("");else value=dec(body.match(/<v>([\s\S]*?)<\/v>/)?.[1]||"");values.set(colNo(ref),value);}out.push({row,values});}return out;}
function norm(v=""){return String(v).toLowerCase().normalize("NFD").replace(/\p{Diacritic}/gu,"").replace(/đ/g,"d").replace(/\b(ky luc gia|klg|ong|ba|anh|chi|gs|pgs|ts|ths|bs|nghe nhan|nghe si|hoa si|nha suu tap|anh hung lao dong|giao su|bac si|tien si|thac si|phap danh|chau a|viet nam|the gioi)\b/g," ").replace(/[^a-z0-9]+/g," ").replace(/\s+/g," ").trim();}
function gabKey(v=""){return String(v).match(/gab\.world\/vi\/bank\/([^/?#]+)/i)?.[1]?.toLowerCase()||"";}
function dateText(v=""){const s=String(v).trim();if(/^\d+(\.0+)?$/.test(s)){const n=Number(s);if(n>20000&&n<70000){const d=new Date(Date.UTC(1899,11,30)+n*86400000);return`${String(d.getUTCDate()).padStart(2,"0")}/${String(d.getUTCMonth()+1).padStart(2,"0")}/${d.getUTCFullYear()}`;}return /^(19|20)\d{2}$/.test(s)?s:"";}return s;}

const strings=shared();
const overlaps=rows("sheet4.xml",strings).filter(r=>r.row>=6).map(r=>({initial:String(r.values.get(2)||"").trim(),dev:String(r.values.get(4)||"").trim(),initialLink:String(r.values.get(5)||"").trim(),devLink:String(r.values.get(6)||"").trim()})).filter(r=>r.initial&&r.dev);
const overlapGroupMap=new Map();
for(const r of overlaps){const key=gabKey(r.devLink)||norm(r.dev);if(!overlapGroupMap.has(key))overlapGroupMap.set(key,{...r});else{const g=overlapGroupMap.get(key);if(!g.initial.split(" | ").includes(r.initial))g.initial+=` | ${r.initial}`;if(!g.initialLink&&r.initialLink)g.initialLink=r.initialLink;}}
const overlapPeople=[...overlapGroupMap.values()];
const overlapNames=new Map(),overlapGab=new Map();
for(const r of overlapPeople){for(const n of r.initial.split(" | "))overlapNames.set(norm(n),r);overlapNames.set(norm(r.dev),r);for(const link of[r.initialLink,r.devLink]){const key=gabKey(link);if(key)overlapGab.set(key,r);}}
const achievements=rows("sheet2.xml",strings).filter(r=>r.row>=3).map(r=>({row:r.row,stt:String(r.values.get(1)||"").trim(),profile:String(r.values.get(2)||"").trim(),cert:String(r.values.get(3)||"").trim(),officialDate:dateText(r.values.get(4)||""),recordTitle:String(r.values.get(5)||"").trim(),gab:String(r.values.get(6)||"").trim(),name:String(r.values.get(7)||"").trim(),otherTitle:String(r.values.get(12)||"").trim(),gabTime:dateText(r.values.get(13)||""),gabTitle:String(r.values.get(14)||"").trim(),article:String(r.values.get(15)||"").trim(),description:String(r.values.get(16)||"").trim(),note:String(r.values.get(19)||"").trim()})).filter(r=>r.name&&(r.recordTitle||r.gabTitle));
const pulled=[];
for(const a of achievements){const g=gabKey(a.gab);let hit=g?overlapGab.get(g):null,basis="";if(hit)basis="Trùng link GAB";if(!hit){hit=overlapNames.get(norm(a.name));if(hit)basis="Trùng tên đã chuẩn hóa";}if(!hit){hit=overlapPeople.find(o=>o.initial.split(" | ").some(n=>norm(n).endsWith(norm(a.name)))||norm(o.dev)===norm(a.name));if(hit)basis="Trùng tên; tab 200 chưa có link GAB";}if(hit)pulled.push({...a,initial:hit.initial,dev:hit.dev,basis});}
const people=new Set(pulled.map(r=>gabKey(r.gab)||norm(r.name))).size;
const matchedOverlapNames=new Set(pulled.flatMap(r=>r.initial.split(" | ").map(norm)));
const missingOverlaps=overlapPeople.filter(r=>!r.initial.split(" | ").some(n=>matchedOverlapNames.has(norm(n)))).map(r=>({initial:r.initial,dev:r.dev,initialLink:r.initialLink,devLink:r.devLink}));
pulled.sort((a,b)=>norm(a.name).localeCompare(norm(b.name))||a.row-b.row);

let styles=read(path.join(xlsx,"xl","styles.xml"));
function append(xml,tag,item){const re=new RegExp(`<${tag}\\b([^>]*)count="(\\d+)"([^>]*)>([\\s\\S]*?)<\\/${tag}>`),m=xml.match(re);if(!m)throw new Error(`Không tìm thấy ${tag}`);return xml.replace(re,`<${tag}${m[1]}count="${Number(m[2])+1}"${m[3]}>${m[4]}${item}</${tag}>`);}
const fontStart=Number(styles.match(/<fonts\b[^>]*count="(\d+)"/)?.[1]||0);styles=append(styles,"fonts",`<font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Arial"/></font>`);
const fillStart=Number(styles.match(/<fills\b[^>]*count="(\d+)"/)?.[1]||0);for(const c of["FF17365D","FFD9E2F3","FFD9EAD3"])styles=append(styles,"fills",`<fill><patternFill patternType="solid"><fgColor rgb="${c}"/><bgColor indexed="64"/></patternFill></fill>`);
const xfStart=Number(styles.match(/<cellXfs\b[^>]*count="(\d+)"/)?.[1]||0);for(const xf of[`<xf numFmtId="0" fontId="${fontStart}" fillId="${fillStart}" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>`,`<xf numFmtId="0" fontId="0" fillId="${fillStart+1}" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>`,`<xf numFmtId="0" fontId="0" fillId="${fillStart+2}" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>`,`<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>`])styles=append(styles,"cellXfs",xf);
write(path.join(xlsx,"xl","styles.xml"),styles);

const cell=(v,ref,s)=>`<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
const row=(vals,n,style,height=38)=>`<row r="${n}" ht="${height}" customHeight="1">${vals.map((v,i)=>cell(v,`${colName(i+1)}${n}`,Array.isArray(style)?style[i]:style)).join("")}</row>`;
const table=[];table.push(row(["THÀNH TỰU CỦA NHÓM TRÙNG BAN ĐẦU - 560 TRONG TAB 200 NGƯỜI ĐÃ CHUẨN"],1,xfStart,30));table.push(row([`Có ${overlaps.length} dòng trùng Ban đầu - 560, tương ứng ${overlapPeople.length} người duy nhất; tìm thấy đủ ${people} người trong tab 200, gồm ${pulled.length} dòng thành tựu.`],2,xfStart+1,24));table.push(row(["Lê Minh Phú có 2 cách ghi trong danh sách trùng nên 43 dòng tương ứng 42 người. Mỗi dòng dưới đây là một thành tựu từ tab 200."],3,xfStart+1,24));table.push(`<row r="4" ht="8" customHeight="1"/>`);
const headers=["STT","TÊN TRONG TAB 200","TÊN Ở BAN ĐẦU","TÊN TRONG DS 560","CĂN CỨ KHỚP","LINK HỒ SƠ GAB","SỐ XÁC LẬP","NGÀY XÁC LẬP","TÊN KỶ LỤC GỐC","THỜI GIAN TRÊN GAB","TIÊU ĐỀ THÀNH TỰU","LINK BÀI VIẾT","MÔ TẢ THÀNH TỰU","GHI CHÚ TAB 200"];
table.push(row(headers,5,xfStart,48));pulled.forEach((r,i)=>table.push(row([i+1,r.name,r.initial,r.dev,r.basis,r.gab,r.cert,r.officialDate,r.recordTitle,r.gabTime,r.gabTitle||r.recordTitle,r.article,r.description,r.note],i+6,[xfStart+3,xfStart+2,xfStart+3,xfStart+3,xfStart+2,xfStart+3,xfStart+3,xfStart+3,xfStart+3,xfStart+3,xfStart+3,xfStart+3,xfStart+3,xfStart+3],52)));
const last=pulled.length+5,widths=[7,27,31,27,23,47,19,16,58,17,66,50,75,34];
const sheet=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><dimension ref="A1:N${last}"/><cols>${widths.map((w,i)=>`<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`).join("")}</cols><sheetData>${table.join("")}</sheetData><autoFilter ref="A5:N${last}"/><mergeCells count="3"><mergeCell ref="A1:N1"/><mergeCell ref="A2:N2"/><mergeCell ref="A3:N3"/></mergeCells></worksheet>`;
write(path.join(xlsx,"xl","worksheets","sheet5.xml"),sheet);
let wb=read(path.join(xlsx,"xl","workbook.xml"));wb=wb.replace(/<sheet[^>]*name="THÀNH TỰU TRÙNG 43 KLG"[^>]*\/>/g,"").replace("</sheets>",`<sheet sheetId="5" name="THÀNH TỰU TRÙNG 43 KLG" state="visible" r:id="rId8"/></sheets>`);write(path.join(xlsx,"xl","workbook.xml"),wb);
let rel=read(path.join(xlsx,"xl","_rels","workbook.xml.rels"));rel=rel.replace(/<Relationship[^>]*Id="rId8"[^>]*\/>/g,"").replace("</Relationships>",`<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet5.xml"/></Relationships>`);write(path.join(xlsx,"xl","_rels","workbook.xml.rels"),rel);
let types=read(path.join(xlsx,"[Content_Types].xml"));if(!types.includes('/xl/worksheets/sheet5.xml'))types=types.replace("</Types>",`<Override PartName="/xl/worksheets/sheet5.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`);write(path.join(xlsx,"[Content_Types].xml"),types);
const zip=path.join(work,"updated.zip"),temp=path.join(work,"So sánh.updated.xlsx");if(fs.existsSync(zip))fs.rmSync(zip,{force:true});if(fs.existsSync(temp))fs.rmSync(temp,{force:true});execFileSync("powershell",["-NoProfile","-Command",`Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::CreateFromDirectory('${xlsx.replace(/'/g,"''")}', '${zip.replace(/'/g,"''")}')`],{stdio:"pipe"});fs.copyFileSync(zip,temp);fs.copyFileSync(temp,source);
console.log(JSON.stringify({source,overlapRows:overlaps.length,uniqueOverlapPeople:overlapPeople.length,matchedPeople:people,achievementRows:pulled.length,missingOverlaps},null,2));
