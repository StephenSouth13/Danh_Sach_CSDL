import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const workbook = path.resolve("outputs", "reconciled_gab_records_CSDL_authoritative_clean_notes.xlsx");
const workDir = path.resolve("outputs", "enhance_hoi_ngo_notes_detail_work");

function ps(command) {
  execFileSync("powershell", ["-NoProfile", "-Command", command], { stdio: "pipe" });
}
function resetDir(dir) {
  if (fs.existsSync(dir)) fs.rmSync(dir, { recursive: true, force: true });
  fs.mkdirSync(dir, { recursive: true });
}
function read(p) { return fs.readFileSync(p, "utf8"); }
function write(p, s) { fs.writeFileSync(p, s, "utf8"); }
function dec(s = "") {
  return String(s).replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
}
function esc(s = "") {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function colIndex(ref) {
  let n = 0;
  for (const ch of (String(ref).match(/^[A-Z]+/)?.[0] || "")) n = n * 26 + ch.charCodeAt(0) - 64;
  return n;
}
function colLetters(n) {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
function sharedStrings(dir) {
  const p = path.join(dir, "xl", "sharedStrings.xml");
  if (!fs.existsSync(p)) return [];
  const xml = read(p);
  return [...xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map((si) =>
    [...si[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((t) => dec(t[1])).join(""));
}
function cellValue(cellXml, attrs, shared) {
  const t = attrs.match(/\bt="([^"]+)"/)?.[1] || "";
  if (t === "s") return shared[Number(cellXml.match(/<v>([\s\S]*?)<\/v>/)?.[1])] ?? "";
  if (t === "inlineStr") return [...cellXml.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((m) => dec(m[1])).join("");
  return dec(cellXml.match(/<v>([\s\S]*?)<\/v>/)?.[1] || "");
}
function rowsFromSheet(dir, sheetFile, shared) {
  const xml = read(path.join(dir, "xl", "worksheets", sheetFile));
  const rows = [];
  for (const row of xml.matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>/g)) {
    const rowNumber = Number(row[1].match(/\br="(\d+)"/)?.[1]);
    if (!rowNumber) continue;
    const values = new Map();
    for (const cell of row[2].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>|<c\b([^>]*)\/>/g)) {
      const attrs = cell[1] || cell[3] || "";
      const ref = attrs.match(/\br="([^"]+)"/)?.[1];
      if (!ref) continue;
      values.set(colIndex(ref), cell[2] ? cellValue(cell[2], attrs, shared) : "");
    }
    rows.push({ rowNumber, values });
  }
  return rows;
}
function normalize(s = "") {
  return String(s).toLowerCase().normalize("NFD").replace(/\p{Diacritic}/gu, "").replace(/đ/g, "d")
    .replace(/\([^)]*\)/g, " ").replace(/\b(ky|ki)\s*luc\s*gia\b/g, " ")
    .replace(/\b(gs|pgs|ts|ths|ong|ba|nghe si|nghe nhan|giao su|tien si|thac si|bac si|danh du)\b/g, " ")
    .replace(/[^a-z0-9]+/g, " ").replace(/\s+/g, " ").trim();
}
function short(s = "", max = 130) {
  const t = String(s ?? "").replace(/\s+/g, " ").trim();
  return t.length > max ? `${t.slice(0, max - 3)}...` : t;
}
function setCell(xml, rowNumber, col, value) {
  const ref = `${colLetters(col)}${rowNumber}`;
  const cellXml = `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`;
  const rowRe = new RegExp(`<row\\b([^>]*)\\br="${rowNumber}"([^>]*)>([\\s\\S]*?)<\\/row>`);
  const rowMatch = xml.match(rowRe);
  if (!rowMatch) return xml;
  let body = rowMatch[3];
  const cellRe = /<c\b[^>]*\br="([^"]+)"[^>]*(?:\/>|>[\s\S]*?<\/c>)/g;
  let replaced = false;
  body = body.replace(cellRe, (match, cellRef) => {
    if (cellRef === ref) {
      replaced = true;
      return cellXml;
    }
    return match;
  });
  if (!replaced) body += cellXml;
  body = [...body.matchAll(/<c\b[^>]*\br="([^"]+)"[^>]*(?:\/>|>[\s\S]*?<\/c>)/g)]
    .map((m) => ({ ref: m[1], xml: m[0] }))
    .sort((a, b) => colIndex(a.ref) - colIndex(b.ref))
    .map((c) => c.xml)
    .join("");
  return xml.replace(rowRe, `<row${rowMatch[1]}r="${rowNumber}"${rowMatch[2]}>${body}</row>`);
}

resetDir(workDir);
fs.copyFileSync(workbook, path.join(workDir, "book.zip"));
ps(`Expand-Archive -LiteralPath '${path.join(workDir, "book.zip").replace(/'/g, "''")}' -DestinationPath '${workDir.replace(/'/g, "''")}' -Force`);

const shared = sharedStrings(workDir);
const data = rowsFromSheet(workDir, "sheet1.xml", shared).filter((r) => r.rowNumber > 1).map((r) => ({
  rowNumber: r.rowNumber,
  recordId: String(r.values.get(1) ?? ""),
  gabId: String(r.values.get(2) ?? ""),
  name: String(r.values.get(3) ?? "").trim(),
  date: String(r.values.get(4) ?? "").trim(),
  title: String(r.values.get(5) ?? "").trim(),
  check: String(r.values.get(8) ?? "").trim(),
  note: String(r.values.get(9) ?? "").trim(),
})).filter((r) => r.name || r.title);
const byName = new Map();
for (const r of data) {
  const key = normalize(r.name);
  if (!byName.has(key)) byName.set(key, []);
  byName.get(key).push(r);
}
function rowsForName(name) {
  const key = normalize(name);
  return byName.get(key) || data.filter((r) => normalize(r.name).includes(key) || key.includes(normalize(r.name)));
}
function line(row, prefix) {
  const fix = row.note.replace(/\s+/g, " ").trim();
  return `${prefix} Data dòng ${row.rowNumber}: "${short(row.title, 140)}"; ngày trong GAB/Data="${row.date || "trống"}"; kết luận="${row.check}". Chi tiết lỗi/cách sửa/nguồn chuẩn: ${short(fix, 760)}`;
}

let hoiXml = read(path.join(workDir, "xl", "worksheets", "sheet2.xml"));
const hoiRows = rowsFromSheet(workDir, "sheet2.xml", shared).filter((r) => r.rowNumber >= 4 && r.rowNumber <= 303);
let written = 0;
for (const r of hoiRows) {
  const name = String(r.values.get(7) || r.values.get(9) || "").trim();
  if (!name) {
    hoiXml = setCell(hoiXml, r.rowNumber, 25, "CẦN BỔ SUNG TÊN");
    hoiXml = setCell(hoiXml, r.rowNumber, 26, "Dòng này chưa có họ tên ở cột G/I nên chưa đủ căn cứ đối chiếu với CSDL chuẩn.");
    written++;
    continue;
  }
  const items = rowsForName(name);
  const wrong = items.filter((x) => normalize(x.check).includes("sai thoi gian"));
  const extra = items.filter((x) => normalize(x.check).includes("du") || normalize(x.check).includes("khong tim thay"));
  const dupes = items.filter((x) => normalize(x.check).includes("trung dong gab"));
  const missingNotes = [...new Set(items.flatMap((x) => x.note.split(/\n/).filter((p) => normalize(p).includes("csdl con") || normalize(p).includes("thieu"))))];
  const ok = items.filter((x) => normalize(x.check).startsWith("dung"));
  const check = wrong.length || extra.length || dupes.length || missingNotes.length ? "SAI/CẦN XỬ LÝ" : "ĐÚNG/CHƯA THẤY LỖI THÀNH TỰU";
  const note = [
    `Đối chiếu "${name}": Data/GAB có ${items.length} dòng liên quan. CSDL chuẩn là căn cứ; nếu GAB/Data khác CSDL thì cần sửa trên GAB hoặc xác minh dữ liệu ngoài CSDL. Khi có lỗi, ghi rõ Data dòng/cột cần sửa và nguồn chuẩn tại file MÔ TẢ YÊU CẦU NHẬP THÔNG TIN GAB.xlsx, tab/dòng/cột tương ứng.`,
    ok.length ? `ĐÃ KHỚP: ${ok.length} dòng đã khớp CSDL chuẩn, không cần sửa thành tựu ở các dòng này.` : "ĐÃ KHỚP: Chưa có dòng nào khớp hoàn toàn.",
    wrong.length ? `SAI THỜI GIAN: ${wrong.length} dòng. ${wrong.slice(0, 10).map((x, i) => line(x, `${i + 1}.`)).join(" ")}` : "SAI THỜI GIAN: Chưa ghi nhận.",
    extra.length ? `DƯ/CHƯA KHỚP CSDL: ${extra.length} dòng. ${extra.slice(0, 10).map((x, i) => line(x, `${i + 1}.`)).join(" ")}` : "DƯ/CHƯA KHỚP CSDL: Chưa ghi nhận.",
    missingNotes.length ? `THIẾU SO VỚI CSDL: ${missingNotes.slice(0, 6).join(" ")}` : "THIẾU SO VỚI CSDL: Chưa ghi nhận.",
    dupes.length ? `TRÙNG DÒNG: ${dupes.length} dòng trùng kỹ thuật, không tính là thành tựu khác nhau. ${dupes.slice(0, 10).map((x) => `Data dòng ${x.rowNumber}: ${short(x.note, 360)}`).join(" ")}` : "TRÙNG DÒNG: Chưa ghi nhận."
  ].join("\n");
  hoiXml = setCell(hoiXml, r.rowNumber, 25, check);
  hoiXml = setCell(hoiXml, r.rowNumber, 26, note);
  written++;
}
write(path.join(workDir, "xl", "worksheets", "sheet2.xml"), hoiXml.replace(/<\/c>>+/g, "</c>"));
const outZip = path.join(workDir, "out.zip");
if (fs.existsSync(outZip)) fs.rmSync(outZip, { force: true });
fs.rmSync(path.join(workDir, "book.zip"), { force: true });
const packageItems = [
  path.join(workDir, "[Content_Types].xml"),
  path.join(workDir, "_rels"),
  path.join(workDir, "docProps"),
  path.join(workDir, "xl"),
].map((p) => `'${p.replace(/'/g, "''")}'`).join(",");
ps(`Compress-Archive -LiteralPath ${packageItems} -DestinationPath '${outZip.replace(/'/g, "''")}' -Force`);
fs.copyFileSync(outZip, workbook);
console.log(JSON.stringify({ workbook, hoiRows: hoiRows.length, written }, null, 2));
