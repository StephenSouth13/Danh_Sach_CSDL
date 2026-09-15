import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const targetXlsx = "D:\\KLG\\Danh_Sach_CSDL\\Dulieutruyxuat\\Bản sao của DỮ LIỆU THÀNH TỰU GAB - 15_24, 13 tháng 9.xlsx";
const refXlsx = "D:\\KLG\\Danh_Sach_CSDL\\Dulieutruyxuat\\MÔ TẢ YÊU CẦU NHẬP THÔNG TIN GAB.xlsx";
const outDir = path.resolve("outputs", "reconcile_grouped_check_notes_work");
const targetDir = path.join(outDir, "target");
const refDir = path.join(outDir, "ref");
const outputXlsx = path.resolve("outputs", "reconciled_gab_records_customer_feedback.xlsx");

function ps(command) {
  execFileSync("powershell", ["-NoProfile", "-Command", command], { stdio: "pipe" });
}

function resetDir(dir) {
  if (fs.existsSync(dir)) fs.rmSync(dir, { recursive: true, force: true });
  fs.mkdirSync(dir, { recursive: true });
}

resetDir(outDir);
fs.mkdirSync(targetDir, { recursive: true });
fs.mkdirSync(refDir, { recursive: true });
const targetZip = path.join(outDir, "target.zip");
const refZip = path.join(outDir, "ref.zip");
fs.copyFileSync(targetXlsx, targetZip);
fs.copyFileSync(refXlsx, refZip);
ps(`Expand-Archive -LiteralPath '${targetZip.replace(/'/g, "''")}' -DestinationPath '${targetDir.replace(/'/g, "''")}' -Force`);
ps(`Expand-Archive -LiteralPath '${refZip.replace(/'/g, "''")}' -DestinationPath '${refDir.replace(/'/g, "''")}' -Force`);

const read = (p) => fs.readFileSync(p, "utf8");
const write = (p, s) => fs.writeFileSync(p, s, "utf8");

function decodeXml(s = "") {
  return s
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
  for (const m of rels.matchAll(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)) {
    relMap.set(m[1], m[2]);
  }
  const map = new Map();
  for (const m of wb.matchAll(/<sheet\b[^>]*name="([^"]+)"[^>]*sheetId="([^"]+)"[^>]*(?:r:id|id)="([^"]+)"/g)) {
    let target = relMap.get(m[3]);
    if (!target) continue;
    if (!target.startsWith("xl/")) target = `xl/${target}`;
    map.set(decodeXml(m[1]), { name: decodeXml(m[1]), id: m[2], target });
  }
  return map;
}

function sharedStrings(dir) {
  const p = path.join(dir, "xl", "sharedStrings.xml");
  if (!fs.existsSync(p)) return [];
  const xml = read(p);
  const result = [];
  for (const si of xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)) {
    let text = "";
    for (const t of si[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)) text += decodeXml(t[1]);
    result.push(text);
  }
  return result;
}

function cellValue(cellXml, attrs, shared) {
  const t = attrs.match(/\bt="([^"]+)"/)?.[1] || "";
  if (t === "s") {
    const v = cellXml.match(/<v>([\s\S]*?)<\/v>/)?.[1];
    return shared[Number(v)] ?? "";
  }
  if (t === "inlineStr") {
    let text = "";
    for (const m of cellXml.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)) text += decodeXml(m[1]);
    return text;
  }
  const v = cellXml.match(/<v>([\s\S]*?)<\/v>/)?.[1];
  return v ? decodeXml(v) : "";
}

function rowsFromSheet(dir, sheet, shared) {
  const xml = read(path.join(dir, sheet.target));
  const rows = [];
  for (const row of xml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)) {
    const values = new Map();
    for (const cell of row[2].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
      const ref = cell[1].match(/\br="([^"]+)"/)?.[1];
      if (!ref) continue;
      values.set(colIndex(ref), cellValue(cell[2], cell[1], shared));
    }
    rows.push({ rowNumber: Number(row[1]), values });
  }
  return rows;
}

function normalize(text = "") {
  return String(text)
    .toLowerCase()
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .replace(/đ/g, "d")
    .replace(/\b(ky|ki)\s*luc\s*gia\b/g, " ")
    .replace(/\b(gs|pgs|ts|ths|nsut|nsnd|ong|ba|anh|chi|thay|co|hoa si|nha bao|nghe si|giao su|vien si|nghe nhan|doanh nhan)\b/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

const stopWords = new Set("viet nam nguoi dau tien nhieu nhat lon the gioi chau luc xac lap ky luc gia duoc cua tai trong voi ve va cac cho mot bo suu tap danh hieu hien co nam ngay cap".split(" "));
function tokens(text = "", keepStop = false) {
  const n = normalize(text);
  if (!n) return [];
  return [...new Set(n.split(/\s+/).filter((t) => t.length > 1 && (keepStop || !stopWords.has(t))))];
}

function scoreTokens(a, b) {
  if (!a.length || !b.length) return 0;
  const as = new Set(a);
  const bs = new Set(b);
  let common = 0;
  for (const t of as) if (bs.has(t)) common++;
  return (common / as.size) * 0.55 + (common / bs.size) * 0.45;
}

function trimText(s = "", max = 180) {
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
  return s.replace(/\s+/g, " ");
}

function dateKey(value = "") {
  const s = excelDateText(value).trim();
  const dmy = s.match(/\b(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})\b/);
  if (dmy) {
    const y = dmy[3].length === 2 ? `20${dmy[3]}` : dmy[3];
    return `${y}-${dmy[2].padStart(2, "0")}-${dmy[1].padStart(2, "0")}`;
  }
  const ymd = s.match(/\b(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})\b/);
  if (ymd) return `${ymd[1]}-${ymd[2].padStart(2, "0")}-${ymd[3].padStart(2, "0")}`;
  return normalize(s);
}

function quoteSheet(name) {
  return `'${String(name).replace(/'/g, "''")}'`;
}

function refMeta(sheet) {
  if (sheet === "CSDL KLVN") {
    return { date: "D", dateName: "NGÀY XÁC LẬP", title: "E", titleName: "TÊN KỶ LỤC", owner: "F", ownerName: "KỶ LỤC GIA", province: "G", provinceName: "Tỉnh", field: "H", fieldName: "LĨNH VỰC XÁC LẬP" };
  }
  if (sheet === "CSDL KLVW") {
    return { date: "E", dateName: "NGÀY CẤP", title: "F", titleName: "TÊN KỶ LỤC", owner: "G", ownerName: "CÁ NHÂN/ĐƠN VỊ SỞ HỮU KỶ LỤC", province: "H", provinceName: "ĐỊA CHỈ/TỈNH", field: "I", fieldName: "THÔNG TIN/LĨNH VỰC" };
  }
  return { date: "E", dateName: "THỜI ĐIỂM XÁC LẬP", title: "C", titleName: "TÊN KỶ LỤC", owner: "B", ownerName: "TÊN CÁ NHÂN/ĐƠN VỊ", province: "G", provinceName: "TỈNH/ĐỊA PHƯƠNG", field: "", fieldName: "" };
}

function refAudit(ref) {
  const m = refMeta(ref.sheet);
  const parts = [
    `${quoteSheet(ref.sheet)}!${m.owner}${ref.row}(${m.ownerName})="${trimText(ref.owner, 90)}"`,
    `${quoteSheet(ref.sheet)}!${m.title}${ref.row}(${m.titleName})="${trimText(ref.title, 150)}"`,
    `${quoteSheet(ref.sheet)}!${m.date}${ref.row}(${m.dateName})="${trimText(excelDateText(ref.date), 40)}"`,
  ];
  if (m.province) parts.push(`${quoteSheet(ref.sheet)}!${m.province}${ref.row}(${m.provinceName})="${trimText(ref.province, 80)}"`);
  if (m.field) parts.push(`${quoteSheet(ref.sheet)}!${m.field}${ref.row}(${m.fieldName})="${trimText(ref.field, 90)}"`);
  return parts.join("; ");
}

function loadRefs() {
  const shared = sharedStrings(refDir);
  const sheets = sheetMap(refDir);
  const refs = [];
  for (const sheetName of ["CSDL KLVN", "CSDL KLVW", "CSDL KL CA", "CSDL KLTG"]) {
    const sheet = sheets.get(sheetName);
    if (!sheet) throw new Error(`Không thấy tab chuẩn: ${sheetName}`);
    for (const r of rowsFromSheet(refDir, sheet, shared)) {
      if (r.rowNumber < 3) continue;
      let ref;
      if (sheetName === "CSDL KLVN") {
        ref = { sheet: sheetName, row: r.rowNumber, date: r.values.get(4), title: r.values.get(5), owner: r.values.get(6), province: r.values.get(7), field: r.values.get(8) };
      } else if (sheetName === "CSDL KLVW") {
        ref = { sheet: sheetName, row: r.rowNumber, date: r.values.get(5), title: r.values.get(6), owner: r.values.get(7), province: r.values.get(8), field: r.values.get(9) };
      } else {
        ref = { sheet: sheetName, row: r.rowNumber, date: r.values.get(5) || r.values.get(4), title: r.values.get(3), owner: r.values.get(2), province: r.values.get(7), field: "" };
      }
      ref.owner = String(ref.owner ?? "");
      ref.title = String(ref.title ?? "");
      ref.date = String(ref.date ?? "");
      ref.province = String(ref.province ?? "");
      ref.field = String(ref.field ?? "");
      ref.ownerTokens = tokens(ref.owner, true);
      ref.titleTokens = tokens(ref.title);
      ref.dateKey = dateKey(ref.date);
      if (ref.owner.trim() && ref.title.trim()) refs.push(ref);
    }
  }
  return refs;
}

function findOwnerRefs(name, refs) {
  const nameNorm = normalize(name);
  const nt = tokens(name, true);
  if (!nameNorm || nt.length < 2) return [];
  return refs
    .map((ref) => {
      const ownerNorm = normalize(ref.owner);
      if (!ownerNorm) return { ref, score: 0, isStrictName: false };
      const ownerSet = new Set(ref.ownerTokens);
      const common = nt.filter((t) => ownerSet.has(t)).length;
      const required = nt.length <= 3 ? nt.length : nt.length - 1;
      let score = scoreTokens(nt, ref.ownerTokens);
      if (ownerNorm.includes(nameNorm) || nameNorm.includes(ownerNorm)) score = Math.max(score, 0.98);
      const phraseMatch = ownerNorm.includes(nameNorm) || nameNorm.includes(ownerNorm);
      const compactOwner = ref.ownerTokens.length <= nt.length + 5;
      const isStrictName = (phraseMatch && compactOwner) || (common >= required && score >= 0.7) || score >= 0.94;
      return { ref, score, isStrictName };
    })
    .filter((m) => m.isStrictName)
    .sort((a, b) => b.score - a.score)
    .map((m) => m.ref);
}

function bestRefForRecord(record, candidates, used) {
  const titleTokens = tokens(record.title);
  let best = null;
  for (const ref of candidates) {
    const key = `${ref.sheet}:${ref.row}`;
    const titleScore = scoreTokens(titleTokens, ref.titleTokens);
    const sameDate = record.dateKey && ref.dateKey && record.dateKey === ref.dateKey;
    const usedPenalty = used.has(key) ? 0.08 : 0;
    const score = titleScore + (sameDate ? 0.12 : 0) - usedPenalty;
    if (!best || score > best.score) best = { ref, score, titleScore, sameDate };
  }
  return best;
}

function recordIssues(record, match) {
  const issues = [];
  if (!trimText(record.name)) issues.push(`thiếu KỶ LỤC GIA ở Data ${quoteSheet(record.sheet)}!C${record.row}`);
  if (!trimText(record.title)) issues.push(`thiếu TÊN KỶ LỤC ở Data ${quoteSheet(record.sheet)}!E${record.row}`);
  if (!trimText(record.date)) issues.push(`thiếu NGÀY XÁC LẬP ở Data ${quoteSheet(record.sheet)}!D${record.row}`);
  if (match?.ref && record.dateKey && match.ref.dateKey && record.dateKey !== match.ref.dateKey) {
    issues.push(`ngày Data ${quoteSheet(record.sheet)}!D${record.row}="${excelDateText(record.date)}" khác CSDL chuẩn ${quoteSheet(match.ref.sheet)}!${refMeta(match.ref.sheet).date}${match.ref.row}="${excelDateText(match.ref.date)}"`);
  }
  return issues;
}

function statusFor(record, match) {
  const missingBase = !trimText(record.name) || !trimText(record.title) || !trimText(record.date);
  if (!match?.ref || match.titleScore < 0.42) return missingBase ? "Thiếu thông tin" : "Không khớp CSDL chuẩn";
  const issues = recordIssues(record, match);
  if (missingBase) return "Thiếu thông tin";
  if (match.titleScore >= 0.78 && !issues.length) return "Khớp";
  if (match.titleScore >= 0.78 && issues.some((issue) => issue.includes("ngày Data"))) return "Sai thời gian xác lập";
  return "Không khớp CSDL chuẩn";
}

function groupNote(group, ownerRefs, rowMatches, statuses) {
  const sheet = quoteSheet(group[0].sheet);
  const start = group[0].row;
  const end = group[group.length - 1].row;
  const name = trimText(group[0].name) || `(thiếu tên tại C${start})`;
  const matchedKeys = new Set(rowMatches.filter(Boolean).map((m) => `${m.ref.sheet}:${m.ref.row}`));
  const missingRefs = ownerRefs.filter((ref) => !matchedKeys.has(`${ref.sheet}:${ref.row}`));
  const extraRows = group.filter((r, i) => !rowMatches[i]?.ref || rowMatches[i].titleScore < 0.42);
  const fixRows = group
    .map((r, i) => ({ r, m: rowMatches[i], issues: recordIssues(r, rowMatches[i]) }))
    .filter((x, i) => statuses[i] !== "Khớp");

  const lines = [];
  lines.push(`Kỷ lục gia/đơn vị: ${name}`);
  lines.push(`Phạm vi Data đang rà soát: ${sheet}!C${start}:C${end}; ngày ở D${start}:D${end}; tên kỷ lục ở E${start}:E${end}.`);
  lines.push(`Kết luận nhóm: ${statuses.every((s) => s === "Khớp") && missingRefs.length === 0 ? "GAB đã đủ và khớp CSDL chuẩn, không cần xử lý thêm." : "GAB còn dòng sai thời gian hoặc còn thiếu thành tựu so với CSDL chuẩn."}`);
  lines.push(`Lưu ý cấu trúc: file GAB gốc có cột C fullName, D time, E title, F url, G description. Bản này bổ sung H Check, I Note, J Tỉnh, K LĨNH VỰC XÁC LẬP để chỉ rõ thông tin chuẩn cần sửa/nhập thêm.`);

  if (!ownerRefs.length) {
    lines.push(`Không tìm thấy kỷ lục gia/đơn vị này trong 4 tab CSDL chuẩn theo cột chủ thể. Cần kiểm tra lại tên ở ${sheet}!C${start}:C${end}; nếu tên không thuộc CSDL chuẩn thì các dòng này không nên giữ như dữ liệu chuẩn.`);
  } else {
    lines.push(`Nguồn CSDL chuẩn tìm được theo tên: ${ownerRefs.slice(0, 10).map(refAudit).join(" | ")}${ownerRefs.length > 10 ? ` | còn ${ownerRefs.length - 10} dòng chuẩn khác cùng tên.` : ""}`);
  }

  if (fixRows.length) {
    lines.push("Dòng Data cần xử lý:");
    for (const { r, m, issues } of fixRows.slice(0, 12)) {
      if (!m?.ref || m.titleScore < 0.42) {
        lines.push(`- ${sheet}!C${r.row}/D${r.row}/E${r.row}: dòng GAB hiện có nhưng chưa tìm được thành tựu tương ứng trong 4 tab CSDL chuẩn; Data đang ghi "${trimText(r.title, 140)}". Cần rà soát riêng vì không thuộc 3 kết quả chính khách yêu cầu.`);
      } else {
        lines.push(`- ${sheet}!C${r.row}/D${r.row}/E${r.row}: đối chiếu gần nhất ${refAudit(m.ref)}. ${issues.length ? `Cần sửa: ${issues.join("; ")}.` : `Tên kỷ lục chỉ khớp một phần (${Math.round(m.titleScore * 100)}%), nên cần chỉnh E${r.row} theo cột tên kỷ lục chuẩn.`}`);
      }
    }
    if (fixRows.length > 12) lines.push(`- Còn ${fixRows.length - 12} dòng Data khác trong nhóm cũng cần rà soát tương tự.`);
  }

  if (missingRefs.length) {
    lines.push("Thành tựu còn thiếu trong GAB theo CSDL chuẩn, đã được bổ sung thành dòng mới trống recordId ở cuối sheet:");
    for (const ref of missingRefs.slice(0, 12)) {
      lines.push(`- Bổ sung/đối chiếu từ ${refAudit(ref)}.`);
    }
    if (missingRefs.length > 12) lines.push(`- Còn ${missingRefs.length - 12} thành tựu chuẩn khác cùng kỷ lục gia chưa thấy trong Data.`);
  }

  if (extraRows.length) {
    lines.push(`Dòng có nguy cơ dư/sai tên kỷ lục: ${extraRows.slice(0, 12).map((r) => `${sheet}!E${r.row}="${trimText(r.title, 100)}"`).join("; ")}.`);
  }

  lines.push("Cách khắc phục: dòng Khớp thì bỏ qua; dòng Sai thời gian xác lập thì lấy ngày chuẩn ở CSDL để sửa GAB; dòng GAB thiếu thành tựu thì nhập mới lên GAB theo dòng bổ sung trống recordId.");
  return lines.join("\n");
}

function setCell(rowXml, cellRef, text) {
  const cell = `<c r="${cellRef}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(text)}</t></is></c>`;
  const re = new RegExp(`<c\\b[^>]*\\br="${cellRef}"[^>]*(?:\\/>|>[\\s\\S]*?<\\/c>)`);
  if (re.test(rowXml)) return rowXml.replace(re, cell);
  const cells = [...rowXml.matchAll(/<c\b[^>]*\br="([A-Z]+\d+)"[^>]*(?:\/>|>[\s\S]*?<\/c>)/g)];
  for (const c of cells) {
    if (colIndex(c[1]) > colIndex(cellRef)) return rowXml.slice(0, c.index) + cell + rowXml.slice(c.index);
  }
  return rowXml.replace("</row>", `${cell}</row>`);
}

function rowXmlFromValues(rowNumber, values) {
  let rowXml = `<row r="${rowNumber}"></row>`;
  values.forEach((value, idx) => {
    rowXml = setCell(rowXml, `${colLetters(idx + 1)}${rowNumber}`, value ?? "");
  });
  return rowXml;
}

function styleHeader(rowXml, cellRef, styleFromRef, label) {
  const re = new RegExp(`<c\\b([^>]*)\\br="${cellRef}"([^>]*)(?:\\/>|>[\\s\\S]*?<\\/c>)`);
  const style = styleFromRef ? ` s="${styleFromRef}"` : "";
  const cell = `<c r="${cellRef}"${style} t="inlineStr"><is><t xml:space="preserve">${escapeXml(label)}</t></is></c>`;
  return re.test(rowXml) ? rowXml.replace(re, cell) : setCell(rowXml, cellRef, label);
}

function addMerges(xml, refs) {
  if (!refs.length) return xml;
  xml = xml.replace(/<mergeCells\b[^>]*>[\s\S]*?<\/mergeCells>/g, "");
  const block = `<mergeCells count="${refs.length}">${refs.map((ref) => `<mergeCell ref="${ref}"/>`).join("")}</mergeCells>`;
  return xml.replace("</sheetData>", `</sheetData>${block}`);
}

function clearOldHiMerges(xml) {
  return xml.replace(/<mergeCells\b([^>]*)>([\s\S]*?)<\/mergeCells>/, (all, attrs, body) => {
    const kept = [...body.matchAll(/<mergeCell\b[^>]*ref="([^"]+)"[^/]*\/>/g)]
      .map((m) => m[0])
      .filter((tag) => !/\bref="[HI]\d+:[HI]\d+"/.test(tag));
    return kept.length ? `<mergeCells count="${kept.length}">${kept.join("")}</mergeCells>` : "";
  });
}

function dataRowHasContent(row) {
  return Boolean(`${row.values.get(1) || ""}${row.values.get(2) || ""}${row.values.get(3) || ""}${row.values.get(4) || ""}${row.values.get(5) || ""}${row.values.get(6) || ""}${row.values.get(7) || ""}`.trim());
}

const refs = loadRefs();
const targetSheets = sheetMap(targetDir);
const targetShared = sharedStrings(targetDir);
const dataSheet = [...targetSheets.values()].find((s) => s.name.startsWith("Data_record_01-01-2022_01-01-20"));
if (!dataSheet) throw new Error("Không thấy tab Data_record_01-01-2022_01-01-20 trong file Data.");
const prioritySheet = targetSheets.get("Danh sách Hội Ngộ KLG") || [...targetSheets.values()].find((s) => s.name.startsWith("Trang tính1"));
const priorityNames = new Set();
const priorityGabIds = new Set();
if (prioritySheet) {
  for (const r of rowsFromSheet(targetDir, prioritySheet, targetShared)) {
    if (r.rowNumber <= 1) continue;
    const gabId = String(r.values.get(2) ?? "").trim();
    const name = String(r.values.get(3) ?? "").trim();
    if (gabId) priorityGabIds.add(gabId);
    if (normalize(name)) priorityNames.add(normalize(name));
  }
}

const dataRows = rowsFromSheet(targetDir, dataSheet, targetShared).filter((r) => r.rowNumber > 1 && dataRowHasContent(r));
const records = dataRows.map((r) => ({
  sheet: dataSheet.name,
  row: r.rowNumber,
  recordId: String(r.values.get(1) ?? ""),
  gabId: String(r.values.get(2) ?? ""),
  name: String(r.values.get(3) ?? ""),
  date: String(r.values.get(4) ?? ""),
  title: String(r.values.get(5) ?? ""),
  url: String(r.values.get(6) ?? ""),
  description: String(r.values.get(7) ?? ""),
  dateKey: dateKey(r.values.get(4) ?? ""),
}));

const groups = [];
for (const rec of records) {
  const key = normalize(rec.name) || `__row_${rec.row}`;
  const last = groups.at(-1);
  if (last && last.key === key) last.records.push(rec);
  else groups.push({ key, records: [rec] });
}

function isPriorityName(name) {
  const n = normalize(name);
  if (!n) return false;
  if (priorityNames.has(n)) return true;
  for (const p of priorityNames) {
    if (n.includes(p) || p.includes(n)) return true;
  }
  return false;
}

function isPriorityRecord(record) {
  return priorityGabIds.has(record.gabId) || isPriorityName(record.name);
}

const statusByRow = new Map();
const noteByGroupStart = new Map();
const mergeRefs = [];
const provinceByRow = new Map();
const fieldByRow = new Map();
const priorityByRow = new Map();
const missingAddRows = [];
const missingByGroup = [];
const matchedStandardKeysGlobal = new Set();
let matched = 0, wrongDate = 0, missingInfo = 0, unmatched = 0;

for (const group of groups) {
  const nameForMatch = group.records.find((r) => normalize(r.name))?.name || group.records[0].name;
  const ownerRefs = findOwnerRefs(nameForMatch, refs);
  const used = new Set();
  const matches = [];
  const statuses = [];
  for (const rec of group.records) {
    const match = bestRefForRecord(rec, ownerRefs, used);
    if (match?.ref && match.titleScore >= 0.42) used.add(`${match.ref.sheet}:${match.ref.row}`);
    matches.push(match);
    const status = statusFor(rec, match);
    statuses.push(status);
    statusByRow.set(rec.row, status);
    if (isPriorityRecord(rec)) priorityByRow.set(rec.row, "Ưu tiên hội ngộ");
    if (match?.ref && match.titleScore >= 0.42) {
      provinceByRow.set(rec.row, match.ref.province || "");
      fieldByRow.set(rec.row, match.ref.field || "");
    }
    if (match?.ref && match.titleScore >= 0.78) {
      matchedStandardKeysGlobal.add(`${match.ref.sheet}:${match.ref.row}`);
    }
    if (status === "Khớp") matched++;
    else if (status === "Sai thời gian xác lập") wrongDate++;
    else if (status === "Thiếu thông tin") missingInfo++;
    else unmatched++;
  }
  const matchedKeys = new Set(matches.filter((m) => m?.ref && m.titleScore >= 0.78).map((m) => `${m.ref.sheet}:${m.ref.row}`));
  const groupMissingRefs = ownerRefs.filter((ref) => !matchedKeys.has(`${ref.sheet}:${ref.row}`));
  if (groupMissingRefs.length) missingByGroup.push({ name: nameForMatch, row: group.records[0].row, count: groupMissingRefs.length });
  const note = groupNote(group.records, ownerRefs, matches, statuses);
  noteByGroupStart.set(group.records[0].row, note);
  if (group.records.length > 1) mergeRefs.push(`I${group.records[0].row}:I${group.records.at(-1).row}`);
}

function findGabGroupForRef(ref) {
  const ownerNorm = normalize(ref.owner);
  if (!ownerNorm) return null;
  let best = null;
  for (const group of groups) {
    const rec = group.records.find((r) => normalize(r.name));
    if (!rec) continue;
    const nameNorm = normalize(rec.name);
    const nt = tokens(rec.name, true);
    const rt = tokens(ref.owner, true);
    const score = scoreTokens(nt, rt);
    const compact = rt.length <= nt.length + 5;
    const phrase = ownerNorm.includes(nameNorm) || nameNorm.includes(ownerNorm);
    const finalScore = phrase && compact ? Math.max(score, 0.98) : score;
    if (!best || finalScore > best.score) best = { group, score: finalScore };
  }
  return best && best.score >= 0.7 ? best.group : null;
}

for (const ref of refs) {
  const key = `${ref.sheet}:${ref.row}`;
  if (matchedStandardKeysGlobal.has(key)) continue;
  const group = findGabGroupForRef(ref);
  missingAddRows.push({ group: group || { records: [{ gabId: "", row: "" }] }, ref, sourceStartRow: group?.records?.[0]?.row || "" });
}

missingAddRows.sort((a, b) => {
  const ap = a.group?.records?.some(isPriorityRecord) || isPriorityName(a.ref.owner);
  const bp = b.group?.records?.some(isPriorityRecord) || isPriorityName(b.ref.owner);
  if (ap !== bp) return ap ? -1 : 1;
  return a.ref.row - b.ref.row;
});

const sheetPath = path.join(targetDir, dataSheet.target);
let xml = read(sheetPath);
const headerRow = xml.match(/<row\b[^>]*\br="1"[^>]*>[\s\S]*?<\/row>/)?.[0] || "";
const gStyle = headerRow.match(/<c\b[^>]*\br="G1"[^>]*\bs="([^"]+)"/)?.[1] || "";
xml = xml.replace(/<row\b[^>]*\br="1"[^>]*>[\s\S]*?<\/row>/, (m) => {
  let out = styleHeader(m, "H1", gStyle, "Check");
  out = styleHeader(out, "I1", gStyle, "Note");
  out = styleHeader(out, "J1", gStyle, "Tỉnh");
  out = styleHeader(out, "K1", gStyle, "LĨNH VỰC XÁC LẬP");
  out = styleHeader(out, "L1", gStyle, "Ưu tiên");
  return out;
});
for (const rec of records) {
  const rowRe = new RegExp(`<row\\b[^>]*\\br="${rec.row}"[^>]*>[\\s\\S]*?<\\/row>`);
  xml = xml.replace(rowRe, (m) => {
    let out = setCell(m, `H${rec.row}`, statusByRow.get(rec.row) || "");
    out = setCell(out, `I${rec.row}`, noteByGroupStart.get(rec.row) || "");
    out = setCell(out, `J${rec.row}`, provinceByRow.get(rec.row) || "");
    out = setCell(out, `K${rec.row}`, fieldByRow.get(rec.row) || "");
    out = setCell(out, `L${rec.row}`, priorityByRow.get(rec.row) || "");
    return out;
  });
}
xml = xml.replace(/<row\b[^>]*\br="(9[4-9]\d|1[0-7]\d\d|18[0-5]\d|186[0-6])"[^>]*>[\s\S]*?<\/row>/g, "");
let appendRow = Math.max(...records.map((r) => r.row)) + 1;
function missingNote(rowNo, item) {
  const { ref, sourceStartRow } = item;
  return [
    "Kết luận: GAB thiếu thành tựu.",
    `Dòng bổ sung này lấy từ CSDL chuẩn nên recordId ở A${rowNo} để trống; gabId ở B${rowNo} lấy theo nhóm KLG đang có trong GAB để dev dễ nhập bổ sung.`,
    sourceStartRow ? `Kỷ lục gia/đơn vị trong GAB dùng để đối chiếu: ${quoteSheet(dataSheet.name)}!C${sourceStartRow}.` : "Chưa tìm thấy gabId/profile tương ứng trong file GAB hiện tại; dòng này vẫn được bổ sung vì CSDL chuẩn có thành tựu.",
    `Nguồn CSDL chuẩn: ${refAudit(ref)}.`,
    `Cần nhập thêm lên GAB: C${rowNo}="${trimText(ref.owner, 100)}"; D${rowNo}="${excelDateText(ref.date)}"; E${rowNo}="${trimText(ref.title, 180)}"; J${rowNo}="${trimText(ref.province, 100)}"; K${rowNo}="${trimText(ref.field, 120)}".`
  ].join("\n");
}
const appendedRowXml = [];
for (const item of missingAddRows) {
  const rowNo = appendRow++;
  item.appendRow = rowNo;
  const gabId = item.group.records[0].gabId || "";
  let rowXml = `<row r="${rowNo}"></row>`;
  rowXml = setCell(rowXml, `A${rowNo}`, "");
  rowXml = setCell(rowXml, `B${rowNo}`, gabId);
  rowXml = setCell(rowXml, `C${rowNo}`, item.ref.owner);
  rowXml = setCell(rowXml, `D${rowNo}`, item.ref.date);
  rowXml = setCell(rowXml, `E${rowNo}`, item.ref.title);
  rowXml = setCell(rowXml, `F${rowNo}`, "");
  rowXml = setCell(rowXml, `G${rowNo}`, `Bổ sung từ ${item.ref.sheet} dòng ${item.ref.row}.`);
  rowXml = setCell(rowXml, `H${rowNo}`, "GAB thiếu thành tựu");
  rowXml = setCell(rowXml, `I${rowNo}`, missingNote(rowNo, item));
  rowXml = setCell(rowXml, `J${rowNo}`, item.ref.province);
  rowXml = setCell(rowXml, `K${rowNo}`, item.ref.field);
  rowXml = setCell(rowXml, `L${rowNo}`, (item.group?.records?.some(isPriorityRecord) || isPriorityName(item.ref.owner)) ? "Ưu tiên hội ngộ" : "");
  appendedRowXml.push(rowXml);
}
xml = xml.replace("</sheetData>", `${appendedRowXml.join("")}</sheetData>`);
xml = clearOldHiMerges(xml);
xml = addMerges(xml, mergeRefs);
xml = xml.replace(/<dimension\b[^>]*ref="[^"]*"[^/]*\/>/, (m) => m.replace(/ref="[^"]*"/, `ref="A1:L${appendRow - 1}"`));
write(sheetPath, xml);

if (prioritySheet) {
  const wbPath = path.join(targetDir, "xl", "workbook.xml");
  let wbXml = read(wbPath);
  wbXml = wbXml.replace(/<sheet\b([^>]*?)name="Trang tính1"([^>]*)>/, `<sheet$1name="Danh sách Hội Ngộ KLG"$2>`);
  write(wbPath, wbXml);

  const priorityHeaders = [
    "Dòng trong Data",
    "recordId",
    "gabId",
    "KỶ LỤC GIA",
    "NGÀY XÁC LẬP",
    "TÊN KỶ LỤC",
    "url",
    "description / nguồn",
    "Check",
    "Note",
    "Tỉnh",
    "LĨNH VỰC XÁC LẬP",
    "Ưu tiên",
  ];
  const priorityRows = [];
  for (const rec of records.filter(isPriorityRecord)) {
    priorityRows.push({
      order: statusByRow.get(rec.row) === "GAB thiếu thành tựu" ? 0 : 1,
      values: [
        rec.row,
        rec.recordId,
        rec.gabId,
        rec.name,
        rec.date,
        rec.title,
        rec.url,
        rec.description,
        statusByRow.get(rec.row) || "",
        noteByGroupStart.get(rec.row) || "",
        provinceByRow.get(rec.row) || "",
        fieldByRow.get(rec.row) || "",
        "Ưu tiên hội ngộ",
      ],
    });
  }
  for (const item of missingAddRows.filter((item) => item.group?.records?.some(isPriorityRecord) || isPriorityName(item.ref.owner))) {
    priorityRows.push({
      order: 0,
      values: [
        item.appendRow,
        "",
        item.group.records[0].gabId || "",
        item.ref.owner,
        item.ref.date,
        item.ref.title,
        "",
        `Bổ sung từ ${item.ref.sheet} dòng ${item.ref.row}.`,
        "GAB thiếu thành tựu",
        missingNote(item.appendRow, item),
        item.ref.province,
        item.ref.field,
        "Ưu tiên hội ngộ",
      ],
    });
  }
  priorityRows.sort((a, b) => a.order - b.order || Number(a.values[0]) - Number(b.values[0]));
  const allPriorityRows = [priorityHeaders, ...priorityRows.map((r) => r.values)];
  const sheetDataPriority = allPriorityRows.map((values, idx) => rowXmlFromValues(idx + 1, values)).join("");
  const prXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:M${allPriorityRows.length}"/><cols><col min="1" max="1" width="14" customWidth="1"/><col min="2" max="3" width="34" customWidth="1"/><col min="4" max="4" width="28" customWidth="1"/><col min="5" max="5" width="16" customWidth="1"/><col min="6" max="6" width="62" customWidth="1"/><col min="7" max="8" width="42" customWidth="1"/><col min="9" max="10" width="38" customWidth="1"/><col min="11" max="13" width="24" customWidth="1"/></cols><sheetData>${sheetDataPriority}</sheetData><autoFilter ref="A1:M${allPriorityRows.length}"/></worksheet>`;
  write(path.join(targetDir, prioritySheet.target), prXml);
}

const zipOut = path.join(outDir, "out.zip");
if (fs.existsSync(outputXlsx)) fs.rmSync(outputXlsx, { force: true });
ps(`Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::CreateFromDirectory('${targetDir.replace(/'/g, "''")}', '${zipOut.replace(/'/g, "''")}')`);
fs.copyFileSync(zipOut, outputXlsx);

const summary = {
  sheet: dataSheet.name,
  dataRows: records.length,
  csdlRows: refs.length,
  groups: groups.length,
  priorityKlgFromTrangTinh1: priorityNames.size,
  priorityExistingGabRows: [...priorityByRow.values()].length,
  priorityMissingAddedRows: missingAddRows.filter((item) => item.group?.records?.some(isPriorityRecord) || isPriorityName(item.ref.owner)).length,
  matched,
  wrongDate,
  gabMissingAddedRows: missingAddRows.length,
  missingInfo,
  unmatched,
  topMissingGroups: missingByGroup.sort((a, b) => b.count - a.count).slice(0, 20),
  outputXlsx,
};
fs.writeFileSync(path.join(outDir, "summary.json"), JSON.stringify(summary, null, 2), "utf8");
console.log(JSON.stringify(summary, null, 2));
