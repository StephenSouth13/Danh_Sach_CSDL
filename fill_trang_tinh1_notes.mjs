import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const workbook = path.resolve("outputs", "reconciled_gab_records_customer_feedback.xlsx");
const backup = path.resolve("outputs", "reconciled_gab_records_customer_feedback_before_trang_tinh1_notes.xlsx");
const workDir = path.resolve("outputs", "fill_trang_tinh1_notes_work");
const zipPath = path.join(workDir, "book.zip");
const outZip = path.join(workDir, "out.zip");

function ps(command) {
  execFileSync("powershell", ["-NoProfile", "-Command", command], { stdio: "pipe" });
}

function resetDir(dir) {
  if (fs.existsSync(dir)) fs.rmSync(dir, { recursive: true, force: true });
  fs.mkdirSync(dir, { recursive: true });
}

function read(p) {
  return fs.readFileSync(p, "utf8");
}

function write(p, s) {
  fs.writeFileSync(p, s, "utf8");
}

function decodeXml(s = "") {
  return String(s)
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function escapeXml(s = "") {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function colIndex(ref) {
  const letters = String(ref).match(/^[A-Z]+/)?.[0] || "";
  let n = 0;
  for (const ch of letters) n = n * 26 + ch.charCodeAt(0) - 64;
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

function sheetMap(dir) {
  const wb = read(path.join(dir, "xl", "workbook.xml"));
  const rels = read(path.join(dir, "xl", "_rels", "workbook.xml.rels"));
  const relMap = new Map();
  for (const m of rels.matchAll(/<Relationship\b([^>]*)\/>/g)) {
    const attrs = m[1];
    const id = attrs.match(/\bId="([^"]+)"/)?.[1];
    const target = attrs.match(/\bTarget="([^"]+)"/)?.[1];
    if (id && target) relMap.set(id, target);
  }
  const map = new Map();
  for (const m of wb.matchAll(/<sheet\b([^>]*)\/>/g)) {
    const attrs = m[1];
    const name = attrs.match(/\bname="([^"]+)"/)?.[1];
    const id = attrs.match(/\br:id="([^"]+)"/)?.[1] || attrs.match(/\bid="([^"]+)"/)?.[1];
    if (!name || !id) continue;
    let target = relMap.get(id);
    if (!target) continue;
    if (target.startsWith("/")) target = target.slice(1);
    if (!target.startsWith("xl/")) target = `xl/${target}`;
    map.set(decodeXml(name), { name: decodeXml(name), target });
  }
  return map;
}

function sharedStrings(dir) {
  const p = path.join(dir, "xl", "sharedStrings.xml");
  if (!fs.existsSync(p)) return [];
  const xml = read(p);
  return [...xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map((si) =>
    [...si[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((m) => decodeXml(m[1])).join("")
  );
}

function cellValue(cellXml, attrs, shared) {
  const t = attrs.match(/\bt="([^"]+)"/)?.[1] || "";
  if (t === "s") return shared[Number(cellXml.match(/<v>([\s\S]*?)<\/v>/)?.[1])] ?? "";
  if (t === "inlineStr") return [...cellXml.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((m) => decodeXml(m[1])).join("");
  return decodeXml(cellXml.match(/<v>([\s\S]*?)<\/v>/)?.[1] || "");
}

function rowsFromSheet(dir, sheet, shared) {
  const xml = read(path.join(dir, sheet.target));
  const rows = [];
  for (const row of xml.matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>/g)) {
    const attrs = row[1];
    const rowNumber = Number(attrs.match(/\br="(\d+)"/)?.[1]);
    if (!rowNumber) continue;
    const values = new Map();
    for (const cell of row[2].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>|<c\b([^>]*)\/>/g)) {
      const cellAttrs = cell[1] || cell[3] || "";
      const ref = cellAttrs.match(/\br="([^"]+)"/)?.[1];
      if (!ref) continue;
      values.set(colIndex(ref), cell[2] ? cellValue(cell[2], cellAttrs, shared) : "");
    }
    rows.push({ rowNumber, values });
  }
  return rows;
}

function normalize(text = "") {
  return String(text)
    .toLowerCase()
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .replace(/đ/g, "d")
    .replace(/\([^)]*\)/g, " ")
    .replace(/\b(ky|ki)\s*luc\s*gia\b/g, " ")
    .replace(/\b(gs|pgs|ts|ths|thac sy|tien si|bac si|ong|ba|anh|chi|hoa si|nghe si|nghe nhan|nha tho|nha van|nha bao|doanh nhan|giao su|vien si|nsut|nsnd|uu tu)\b/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function nameKeys(text = "") {
  const base = normalize(text);
  const raw = String(text);
  const keys = new Set([base]);
  for (const m of raw.matchAll(/\(([^)]{2,})\)/g)) keys.add(normalize(m[1]));
  return [...keys].filter(Boolean);
}

function titleShort(s = "", max = 115) {
  const t = String(s ?? "").replace(/\s+/g, " ").trim();
  return t.length > max ? `${t.slice(0, max - 3)}...` : t;
}

function excelDateText(value = "") {
  const s = String(value ?? "").trim();
  if (!s) return "";
  if (/^\d+(\.0+)?$/.test(s)) {
    const serial = Number(s);
    if (serial > 20000 && serial < 70000) {
      const date = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
      return `${String(date.getUTCDate()).padStart(2, "0")}/${String(date.getUTCMonth() + 1).padStart(2, "0")}/${date.getUTCFullYear()}`;
    }
  }
  const js = s.match(/^(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+([A-Z][a-z]{2})\s+(\d{1,2})\s+(\d{4})/);
  if (js) {
    const months = { Jan: "01", Feb: "02", Mar: "03", Apr: "04", May: "05", Jun: "06", Jul: "07", Aug: "08", Sep: "09", Oct: "10", Nov: "11", Dec: "12" };
    return `${String(js[2]).padStart(2, "0")}/${months[js[1]]}/${js[3]}`;
  }
  return s.replace(/\s+/g, " ");
}

function rowRefList(rows, max = 4) {
  return rows.slice(0, max).map((r) => {
    const date = excelDateText(r.time);
    return `Data dòng ${r.rowNumber}: "${titleShort(r.title)}"${date ? ` - ngày ${date}` : ""}`;
  });
}

function findMatches(index, priorityName) {
  const keys = nameKeys(priorityName);
  const found = new Map();
  for (const key of keys) {
    for (const row of index) {
      if (!row.nameKey || !key) continue;
      const ok = row.nameKey === key || row.nameKey.includes(key) || key.includes(row.nameKey);
      if (ok) found.set(row.rowNumber, row);
    }
  }
  return [...found.values()].sort((a, b) => a.rowNumber - b.rowNumber);
}

function makeNote(name, matches) {
  if (!String(name ?? "").trim()) {
    return "Chưa có họ tên ở cột G/I nên chưa đủ căn cứ đối chiếu với Data_record/CSDL. Cần bổ sung họ tên kỷ lục gia trước khi rà thành tựu.";
  }
  if (!matches.length) {
    return `Chưa tìm thấy thành tựu tương ứng theo tên "${name}" trong Data_record đã đối chiếu. Cần kiểm tra lại cách ghi họ tên/tên nghệ danh hoặc bổ sung thành tựu từ CSDL chuẩn nếu 4 sheet có dữ liệu.`;
  }

  const gabRows = matches.filter((r) => r.recordId);
  const missing = matches.filter((r) => normalize(r.check) === "gab thieu thanh tuu");
  const wrongDate = matches.filter((r) => normalize(r.check) === "sai thoi gian xac lap");
  const extra = matches.filter((r) => normalize(r.check) === "khong khop csdl chuan");
  const ok = matches.filter((r) => normalize(r.check).includes("khop") && normalize(r.check) !== "khong khop csdl chuan");

  const parts = [];
  parts.push(`Đối chiếu theo "${name}": tìm thấy ${matches.length} dòng liên quan trong Data_record (${gabRows.length} dòng GAB có recordId, ${missing.length} dòng bổ sung từ CSDL đang trống recordId).`);
  if (missing.length) {
    parts.push(`GAB còn thiếu ${missing.length} thành tựu theo CSDL chuẩn; cần nhập mới lên GAB các dòng: ${rowRefList(missing, 5).join("; ")}.`);
  }
  if (wrongDate.length) {
    parts.push(`Có ${wrongDate.length} thành tựu đã có trên GAB nhưng sai/thừa thiếu thời gian xác lập; cần sửa ngày trên GAB theo ngày chuẩn đã ghi ở cột Note của Data_record: ${rowRefList(wrongDate, 4).join("; ")}.`);
  }
  if (extra.length) {
    parts.push(`Có ${extra.length} dòng GAB chưa khớp thành tựu trong 4 sheet CSDL chuẩn; cần xác minh có phải dữ liệu ngoài CSDL hay cần chuẩn hóa/bỏ khỏi GAB: ${rowRefList(extra, 4).join("; ")}.`);
  }
  if (!missing.length && !wrongDate.length && !extra.length) {
    parts.push(`Chưa ghi nhận thiếu thành tựu hoặc sai thời gian so với phần Data_record đã đối chiếu; các dòng đang có recordId có thể ưu tiên kiểm tra hình/link/thông tin thẻ nếu khách yêu cầu.`);
  } else if (ok.length) {
    parts.push(`Ngoài các điểm trên, có ${ok.length} dòng đã khớp CSDL chuẩn, không cần xử lý thành tựu.`);
  }
  return parts.join("\n");
}

function setCellInline(xml, rowNumber, col, value) {
  const colName = colLetters(col);
  const ref = `${colName}${rowNumber}`;
  const cellXml = `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
  const rowRe = new RegExp(`<row\\b([^>]*)\\br="${rowNumber}"([^>]*)>([\\s\\S]*?)<\\/row>`);
  const rowMatch = xml.match(rowRe);
  if (!rowMatch) return xml;
  let body = rowMatch[3];
  const cellRe = new RegExp(`<c\\b[^>]*\\br="${ref}"[^>]*(?:>[\\s\\S]*?<\\/c>|\\/)?>`);
  if (cellRe.test(body)) {
    body = body.replace(cellRe, cellXml);
  } else {
    const cells = [...body.matchAll(/<c\b[^>]*\br="([^"]+)"[^>]*(?:>[\s\S]*?<\/c>|\/>)/g)];
    let inserted = false;
    let nextBody = "";
    let cursor = 0;
    for (const c of cells) {
      const idx = c.index ?? 0;
      if (!inserted && colIndex(c[1]) > col) {
        nextBody += body.slice(cursor, idx) + cellXml;
        cursor = idx;
        inserted = true;
      }
    }
    body = inserted ? nextBody + body.slice(cursor) : body + cellXml;
  }
  return xml.replace(rowRe, `<row${rowMatch[1]}r="${rowNumber}"${rowMatch[2]}>${body}</row>`);
}

if (!fs.existsSync(workbook)) throw new Error(`Missing workbook: ${workbook}`);
fs.copyFileSync(workbook, backup);
resetDir(workDir);
fs.copyFileSync(workbook, zipPath);
ps(`Expand-Archive -LiteralPath '${zipPath.replace(/'/g, "''")}' -DestinationPath '${workDir.replace(/'/g, "''")}' -Force`);

const shared = sharedStrings(workDir);
const sheets = sheetMap(workDir);
const dataSheet = [...sheets.values()].find((s) => s.name.startsWith("Data_record"));
const prioritySheet = [...sheets.values()].find((s) => normalize(s.name) === "trang tinh1");
if (!dataSheet || !prioritySheet) throw new Error("Khong thay Data_record hoac Trang tinh1");

const dataRows = rowsFromSheet(workDir, dataSheet, shared)
  .filter((r) => r.rowNumber > 1)
  .map((r) => ({
    rowNumber: r.rowNumber,
    recordId: String(r.values.get(1) ?? "").trim(),
    name: String(r.values.get(3) ?? "").trim(),
    time: String(r.values.get(4) ?? "").trim(),
    title: String(r.values.get(5) ?? "").trim(),
    check: String(r.values.get(8) ?? "").trim(),
    note: String(r.values.get(9) ?? "").trim(),
    nameKey: normalize(r.values.get(3) ?? ""),
  }))
  .filter((r) => r.nameKey);

const priorityRows = rowsFromSheet(workDir, prioritySheet, shared).filter((r) => r.rowNumber >= 4 && r.rowNumber <= 303);
let sheetXml = read(path.join(workDir, prioritySheet.target));
let withName = 0;
let blankName = 0;
let notesWritten = 0;
let missingPeople = 0;
let wrongDatePeople = 0;
let extraPeople = 0;

for (const row of priorityRows) {
  const name = String(row.values.get(7) ?? row.values.get(9) ?? "").trim();
  if (name) withName++;
  else blankName++;
  const matches = findMatches(dataRows, name);
  if (matches.some((m) => normalize(m.check) === "gab thieu thanh tuu")) missingPeople++;
  if (matches.some((m) => normalize(m.check) === "sai thoi gian xac lap")) wrongDatePeople++;
  if (matches.some((m) => normalize(m.check) === "khong khop csdl chuan")) extraPeople++;
  sheetXml = setCellInline(sheetXml, row.rowNumber, 18, makeNote(name, matches));
  notesWritten++;
}

sheetXml = sheetXml.replace(/<\/c>>+/g, "</c>");
write(path.join(workDir, prioritySheet.target), sheetXml);
if (fs.existsSync(outZip)) fs.rmSync(outZip, { force: true });
ps(`Compress-Archive -Path '${path.join(workDir, "*").replace(/'/g, "''")}' -DestinationPath '${outZip.replace(/'/g, "''")}' -Force`);
fs.copyFileSync(outZip, workbook);

console.log(JSON.stringify({
  workbook,
  backup,
  prioritySheet: prioritySheet.name,
  dataRows: dataRows.length,
  priorityRows: priorityRows.length,
  withName,
  blankName,
  notesWritten,
  missingPeople,
  wrongDatePeople,
  extraPeople,
}, null, 2));
