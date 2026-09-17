import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const root = process.cwd();
const feedback = path.join(root, "DỮ LIỆU THÀNH TỰU GAB đối chiếu CSDL ĐÃ CHUẨN HÓA (1).xlsx");
const source200 = path.join(root, "Dulieutruyxuat", "Bản sao của DỮ LIỆU THÀNH TỰU GAB - 15_24, 13 tháng 9.xlsx");
const work = path.join(root, "outputs", "clear_achievement_gap_work");
const out = path.join(root, "outputs", "DANH_SACH_THANH_TUU_CON_THIEU_CUA_200_KLG.xlsx");

function ps(command) { execFileSync("powershell", ["-NoProfile", "-Command", command], { stdio: "pipe" }); }
function reset(dir) { if (fs.existsSync(dir)) fs.rmSync(dir, { recursive: true, force: true }); fs.mkdirSync(dir, { recursive: true }); }
function read(file) { return fs.readFileSync(file, "utf8"); }
function esc(value = "") { return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/\"/g, "&quot;"); }
function dec(value = "") { return String(value).replace(/&#x([0-9a-f]+);/gi, (_, x) => String.fromCodePoint(parseInt(x, 16))).replace(/&#(\d+);/g, (_, x) => String.fromCodePoint(Number(x))).replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&"); }
function colNo(ref) { let n = 0; for (const c of (String(ref).match(/^[A-Z]+/)?.[0] || "")) n = n * 26 + c.charCodeAt(0) - 64; return n; }
function colName(n) { let s = ""; while (n) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
function mapSheets(dir) {
  const wb = read(path.join(dir, "xl", "workbook.xml"));
  const rel = read(path.join(dir, "xl", "_rels", "workbook.xml.rels"));
  const links = new Map([...rel.matchAll(/<Relationship\b([^>]*)\/>/g)].map((m) => [m[1].match(/\bId="([^"]+)"/)?.[1], m[1].match(/\bTarget="([^"]+)"/)?.[1]]));
  const result = new Map();
  for (const m of wb.matchAll(/<sheet\b([^>]*)\/>/g)) {
    const attrs = m[1]; const name = dec(attrs.match(/\bname="([^"]+)"/)?.[1] || ""); const id = attrs.match(/\br:id="([^"]+)"/)?.[1]; let target = links.get(id);
    if (!name || !target) continue; if (target.startsWith("/")) target = target.slice(1); if (!target.startsWith("xl/")) target = `xl/${target}`; result.set(name, { target });
  }
  return result;
}
function shared(dir) {
  const file = path.join(dir, "xl", "sharedStrings.xml"); if (!fs.existsSync(file)) return [];
  return [...read(file).matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map((m) => [...m[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((x) => dec(x[1])).join(""));
}
function rows(dir, sheet, strings) {
  const result = [];
  for (const m of read(path.join(dir, sheet.target)).matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>/g)) {
    const row = Number(m[1].match(/\br="(\d+)"/)?.[1]); const values = new Map();
    for (const c of m[2].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>|<c\b([^>]*)\/>/g)) {
      const attrs = c[1] || c[3] || ""; const ref = attrs.match(/\br="([^"]+)"/)?.[1]; if (!ref) continue;
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
function norm(value = "") { return String(value).toLowerCase().normalize("NFD").replace(/\p{Diacritic}/gu, "").replace(/đ/g, "d").replace(/\([^)]*\)/g, " ").replace(/\b(ky luc gia|klg|ong|ba|anh|chi|gs|pgs|ts|ths|bs|nghe nhan|nghe si|tien si|thac si)\b/g, " ").replace(/[^a-z0-9]+/g, " ").replace(/\s+/g, " ").trim(); }
function tokens(value = "") { return new Set(norm(value).split(" ").filter((x) => x.length > 1 && !["va", "cua", "tai", "duoc", "xac", "lap", "ky", "luc", "viet", "nam"].includes(x))); }
function score(a, b) { const x = tokens(a), y = tokens(b); if (!x.size || !y.size) return 0; let hit = 0; for (const t of x) if (y.has(t)) hit++; return (2 * hit) / (x.size + y.size); }
function gabKey(value = "") { return String(value).match(/gab\.world\/vi\/bank\/([^/?#]+)/i)?.[1]?.toLowerCase() || ""; }
function dateText(value = "") { const s = String(value).trim(); if (/^\d+(\.0+)?$/.test(s)) { const n = Number(s); if (n > 20000 && n < 70000) { const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000); return `${String(d.getUTCDate()).padStart(2, "0")}/${String(d.getUTCMonth() + 1).padStart(2, "0")}/${d.getUTCFullYear()}`; } return /^(19|20)\d{2}$/.test(s) ? s : ""; } const js = s.match(/^(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+([A-Z][a-z]{2})\s+(\d{1,2})\s+(\d{4})/); if (js) { const months = { Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12 }; return `${String(js[2]).padStart(2,"0")}/${String(months[js[1]]).padStart(2,"0")}/${js[3]}`; } return s.replace(/\s+/g, " "); }
function cleanUrl(value = "") { const s = String(value).trim(); return /^https?:\/\//i.test(s) ? s : ""; }

reset(work); const fDir = path.join(work, "feedback"), sDir = path.join(work, "source"), build = path.join(work, "xlsx"); reset(fDir); reset(sDir); reset(build);
fs.copyFileSync(feedback, path.join(work, "feedback.zip")); fs.copyFileSync(source200, path.join(work, "source.zip"));
ps(`Expand-Archive -LiteralPath '${path.join(work, "feedback.zip").replace(/'/g, "''")}' -DestinationPath '${fDir.replace(/'/g, "''")}' -Force`);
ps(`Expand-Archive -LiteralPath '${path.join(work, "source.zip").replace(/'/g, "''")}' -DestinationPath '${sDir.replace(/'/g, "''")}' -Force`);

const fm = mapSheets(fDir), sm = mapSheets(sDir), fsx = shared(fDir), ssx = shared(sDir);
const requested = rows(fDir, fm.get("Trang tính1"), fsx).filter((r) => r.row > 1).map((r) => ({ name: String(r.values.get(2)||"").trim(), status: String(r.values.get(5)||"").trim(), source: String(r.values.get(7)||"").trim(), nameKey: norm(r.values.get(2)||""), gab: gabKey(r.values.get(7)||"") })).filter((r) => norm(r.status).includes("chua co tren gab"));
const actual = rows(fDir, fm.get("Data_record_01-01-2022_01-01-20"), fsx).filter((r) => r.row > 1).map((r) => ({ row:r.row, id:String(r.values.get(1)||"").trim(), gab:String(r.values.get(2)||"").trim(), name:String(r.values.get(3)||"").trim(), time:dateText(r.values.get(4)||""), title:String(r.values.get(5)||"").trim(), url:String(r.values.get(6)||"").trim(), nameKey:norm(r.values.get(3)||""), gabKey:gabKey(r.values.get(2)||"") }));
const src = rows(sDir, sm.get("200 KLG CHUẨN HÓA"), ssx).filter((r) => r.row >= 3).map((r) => ({ row:r.row, time:dateText(r.values.get(13)||r.values.get(4)||r.values.get(21)||""), rawTitle:String(r.values.get(5)||"").trim(), title:String(r.values.get(14)||r.values.get(5)||"").trim(), gab:cleanUrl(r.values.get(6)||""), name:String(r.values.get(7)||"").trim(), url:cleanUrl(r.values.get(15)||""), desc:String(r.values.get(16)||r.values.get(20)||"").trim(), nameKey:norm(r.values.get(7)||""), gabKey:gabKey(r.values.get(6)||"") })).filter((r) => r.name && r.title);

const wantedGab = new Set(requested.map((r) => r.gab).filter(Boolean)); const wantedNames = new Set(requested.map((r) => r.nameKey).filter(Boolean));
const scoped = src.filter((r) => (r.gabKey && wantedGab.has(r.gabKey)) || wantedNames.has(r.nameKey));
const duplicateCount = new Map(); for (const r of scoped) { const key = `${r.gabKey||r.nameKey}|${norm(r.title)}|${r.time}`; duplicateCount.set(key, (duplicateCount.get(key)||0)+1); }
const emittedDup = new Set(); const issues = [];
for (const r of scoped) {
  const ownerActual = actual.filter((a) => (r.gabKey && a.gabKey === r.gabKey) || (!r.gabKey && a.nameKey === r.nameKey) || (a.nameKey && a.nameKey === r.nameKey));
  let best = null; for (const a of ownerActual) { const s = Math.max(score(r.title, a.title), score(r.rawTitle, a.title)); const urlHit = r.url && a.url && r.url.toLowerCase() === a.url.toLowerCase(); const rank = urlHit ? 1 : s; if (!best || rank > best.rank) best = { ...a, rank, titleScore:s, urlHit }; }
  const dupKey = `${r.gabKey||r.nameKey}|${norm(r.title)}|${r.time}`;
  if ((duplicateCount.get(dupKey)||0) > 1) {
    if (emittedDup.has(dupKey)) continue; emittedDup.add(dupKey);
    issues.push({ ...r, conclusion:"TRÙNG TRONG TAB 200", action:`Giữ 1 dòng, kiểm tra và loại ${duplicateCount.get(dupKey)-1} dòng trùng trước khi import.`, nearest:best?.title||"", gabTime:best?.time||"", style:3 }); continue;
  }
  if (best && (best.urlHit || best.titleScore >= 0.72)) continue;
  if (best && best.titleScore >= 0.46) issues.push({ ...r, conclusion:"CẦN RÀ TRƯỚC KHI IMPORT", action:"GAB có một thành tựu gần giống; so lại tiêu đề và thời gian để tránh nhập trùng.", nearest:best.title, gabTime:best.time, style:2 });
  else issues.push({ ...r, conclusion:"THIẾU TRÊN GAB", action:"Bổ sung thành tựu này vào GAB theo thông tin từ tab 200 KLG chuẩn hóa.", nearest:best?.title||"", gabTime:best?.time||"", style:1 });
}
issues.sort((a,b) => norm(a.name).localeCompare(norm(b.name)) || a.conclusion.localeCompare(b.conclusion) || a.title.localeCompare(b.title));

const uniquePeople = new Set(issues.map((r) => r.gabKey || r.nameKey)).size; const missing = issues.filter((r) => r.style===1).length, review = issues.filter((r) => r.style===2).length, dup = issues.filter((r) => r.style===3).length;
const headers = ["STT","TÊN KỶ LỤC GIA","KẾT LUẬN","CẦN LÀM GÌ","LINK HỒ SƠ GAB","THỜI GIAN CHUẨN","THÀNH TỰU THEO TAB 200","LINK BÀI VIẾT NGUỒN","MÔ TẢ ĐỂ IMPORT","THÀNH TỰU ĐANG CÓ GẦN NHẤT TRÊN GAB","THỜI GIAN TRÊN GAB"];
const data = issues.map((r,i) => [i+1,r.name,r.conclusion,r.action,r.gab,r.time,r.title,r.url,r.desc,r.nearest,r.gabTime]);

function cell(value, ref, style=0) { return `<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`; }
function rowXml(values, row, styles=[]) { return `<row r="${row}"${row===1?' ht="28" customHeight="1"':row<=3?' ht="22" customHeight="1"':row===5?' ht="44" customHeight="1"':' ht="42" customHeight="1"'}>${values.map((v,i)=>cell(v,`${colName(i+1)}${row}`,styles[i]??0)).join("")}</row>`; }
const sheetRows = [];
sheetRows.push(rowXml(["DANH SÁCH CHÊNH LỆCH THÀNH TỰU CỦA 200 KLG VỚI GAB"],1,[4]));
sheetRows.push(rowXml([`Kết quả: ${issues.length} dòng cần xử lý của ${uniquePeople} KLG | Thiếu rõ: ${missing} | Cần rà gần trùng: ${review} | Trùng trong tab 200: ${dup}`],2,[5]));
sheetRows.push(rowXml(["Cách đọc: xanh = có thể bổ sung; vàng = kiểm tra thành tựu gần giống; đỏ = dữ liệu trùng ngay trong tab 200. Các thành tựu đã khớp đủ không xuất hiện."],3,[6]));
sheetRows.push(`<row r="4" ht="8" customHeight="1"/>`);
sheetRows.push(rowXml(headers,5,headers.map(()=>7)));
for (let i=0;i<data.length;i++) { const base = issues[i].style===1?8:issues[i].style===2?9:10; const styles = data[i].map((_,c)=>c===2||c===3?base:11); sheetRows.push(rowXml(data[i],i+6,styles)); }
const last = data.length+5; const widths=[7,27,25,38,46,16,68,48,74,62,16];
const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><dimension ref="A1:K${last}"/><cols>${widths.map((w,i)=>`<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`).join("")}</cols><sheetData>${sheetRows.join("")}</sheetData><autoFilter ref="A5:K${last}"/><mergeCells count="3"><mergeCell ref="A1:K1"/><mergeCell ref="A2:K2"/><mergeCell ref="A3:K3"/></mergeCells></worksheet>`;

fs.mkdirSync(path.join(build,"_rels"),{recursive:true}); fs.mkdirSync(path.join(build,"xl","_rels"),{recursive:true}); fs.mkdirSync(path.join(build,"xl","worksheets"),{recursive:true}); fs.mkdirSync(path.join(build,"docProps"),{recursive:true});
fs.writeFileSync(path.join(build,"[Content_Types].xml"),`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`);
fs.writeFileSync(path.join(build,"_rels",".rels"),`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
fs.writeFileSync(path.join(build,"xl","workbook.xml"),`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="THÀNH TỰU CÒN THIẾU" sheetId="1" r:id="rId1"/></sheets></workbook>`);
fs.writeFileSync(path.join(build,"xl","_rels","workbook.xml.rels"),`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
fs.writeFileSync(path.join(build,"xl","styles.xml"),`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="4"><font><sz val="10"/><name val="Arial"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="16"/><name val="Arial"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="10"/><name val="Arial"/></font><font><i/><color rgb="FF44546A"/><sz val="10"/><name val="Arial"/></font></fonts><fills count="7"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF17365D"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9EAD3"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF4CCCC"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9E2F3"/></patternFill></fill></fills><borders count="2"><border/><border><left style="thin"><color rgb="FFD9E1F2"/></left><right style="thin"><color rgb="FFD9E1F2"/></right><top style="thin"><color rgb="FFD9E1F2"/></top><bottom style="thin"><color rgb="FFD9E1F2"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="12"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="3" borderId="1" xfId="0"/><xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf><xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="3" fillId="6" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf></cellXfs></styleSheet>`);
fs.writeFileSync(path.join(build,"xl","worksheets","sheet1.xml"),sheetXml);
fs.writeFileSync(path.join(build,"docProps","core.xml"),`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Danh sách thành tựu còn thiếu của 200 KLG</dc:title></cp:coreProperties>`);
fs.writeFileSync(path.join(build,"docProps","app.xml"),`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Codex</Application></Properties>`);
const zip = path.join(work,"result.zip"); if (fs.existsSync(zip)) fs.rmSync(zip,{force:true}); if (fs.existsSync(out)) fs.rmSync(out,{force:true}); ps(`Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::CreateFromDirectory('${build.replace(/'/g,"''")}', '${zip.replace(/'/g,"''")}')`); fs.copyFileSync(zip,out);
console.log(JSON.stringify({out, requestedRows:requested.length, scoped200Rows:scoped.length, actualRows:actual.length, issues:issues.length, people:uniquePeople, missing, review, duplicate:dup},null,2));
