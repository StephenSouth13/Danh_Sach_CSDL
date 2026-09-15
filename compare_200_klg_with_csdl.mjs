import fs from "node:fs";
import path from "node:path";
import { execFileSync } from "node:child_process";

const root = process.cwd();
const sourceXlsx = path.join(root, "Dulieutruyxuat", "Bản sao của DỮ LIỆU THÀNH TỰU GAB - 15_24, 13 tháng 9.xlsx");
const refXlsx = path.join(root, "Dulieutruyxuat", "MÔ TẢ YÊU CẦU NHẬP THÔNG TIN GAB.xlsx");
const outRoot = path.join(root, "outputs", "compare_200_klg_with_csdl");
const sourceDir = path.join(outRoot, "source");
const refDir = path.join(outRoot, "ref");
const buildDir = path.join(outRoot, "xlsx");
const outputXlsx = path.join(root, "outputs", "so_sanh_200_KLG_CHUAN_HOA_vs_4_tab_CSDL.xlsx");

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
function esc(s = "") {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
function dec(s = "") {
  return String(s)
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
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
function sheetMap(dir) {
  const wb = read(path.join(dir, "xl", "workbook.xml"));
  const rels = read(path.join(dir, "xl", "_rels", "workbook.xml.rels"));
  const relMap = new Map();
  for (const m of rels.matchAll(/<Relationship\b([^>]*)\/>/g)) {
    const id = m[1].match(/\bId="([^"]+)"/)?.[1];
    const target = m[1].match(/\bTarget="([^"]+)"/)?.[1];
    if (id && target) relMap.set(id, target);
  }
  const map = new Map();
  for (const m of wb.matchAll(/<sheet\b([^>]*)\/>/g)) {
    const name = dec(m[1].match(/\bname="([^"]+)"/)?.[1] || "");
    const id = m[1].match(/\br:id="([^"]+)"/)?.[1] || m[1].match(/\bid="([^"]+)"/)?.[1];
    let target = relMap.get(id);
    if (!name || !target) continue;
    if (target.startsWith("/")) target = target.slice(1);
    if (!target.startsWith("xl/")) target = `xl/${target}`;
    map.set(name, { name, target });
  }
  return map;
}
function sharedStrings(dir) {
  const p = path.join(dir, "xl", "sharedStrings.xml");
  if (!fs.existsSync(p)) return [];
  const xml = read(p);
  return [...xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map((si) =>
    [...si[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((t) => dec(t[1])).join("")
  );
}
function cellValue(cellXml, attrs, shared) {
  const t = attrs.match(/\bt="([^"]+)"/)?.[1] || "";
  if (t === "s") return shared[Number(cellXml.match(/<v>([\s\S]*?)<\/v>/)?.[1])] ?? "";
  if (t === "inlineStr") return [...cellXml.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((m) => dec(m[1])).join("");
  return dec(cellXml.match(/<v>([\s\S]*?)<\/v>/)?.[1] || "");
}
function rowsFromSheet(dir, sheet, shared) {
  const xml = read(path.join(dir, sheet.target));
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
  return String(s)
    .toLowerCase()
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .replace(/đ/g, "d")
    .replace(/\b(ky|ki)\s*luc\s*gia\b/g, " ")
    .replace(/\b(klg|gs|pgs|ts|ths|bs|nsut|nsnd|ong|ba|anh|chi|hoa si|nha bao|nghe si|nghe nhan|giao su|vien si|tien si|thac si|danh du)\b/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}
const stop = new Set("xac lap ky luc viet nam the gioi chau a toan cau nguoi dau tien nhieu nhat lon nhat cao nhat co cua va voi trong tai nam ngay bo duoc bang bang nhan".split(" "));
function tokens(s = "", keepStop = false) {
  const n = normalize(s);
  if (!n) return [];
  return [...new Set(n.split(" ").filter((x) => x.length > 1 && (keepStop || !stop.has(x))))];
}
function score(a, b) {
  if (!a.length || !b.length) return 0;
  const as = new Set(a);
  const bs = new Set(b);
  let common = 0;
  for (const t of as) if (bs.has(t)) common++;
  const ca = common / as.size;
  const cb = common / bs.size;
  return Math.max(ca, cb) * 0.65 + Math.min(ca, cb) * 0.35;
}
function short(s = "", max = 130) {
  const t = String(s ?? "").replace(/\s+/g, " ").trim();
  return t.length > max ? `${t.slice(0, max - 3)}...` : t;
}
function excelDateText(value = "") {
  const s = String(value ?? "").trim();
  if (!s) return "";
  if (/^\d+(\.0+)?$/.test(s)) {
    const serial = Number(s);
    if (serial > 20000 && serial < 70000) {
      const d = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
      return `${String(d.getUTCDate()).padStart(2, "0")}/${String(d.getUTCMonth() + 1).padStart(2, "0")}/${d.getUTCFullYear()}`;
    }
  }
  return s.replace(/\s+/g, " ");
}
function parseDateParts(value = "") {
  const s = excelDateText(value);
  if (/^\d+(\.0+)?$/.test(String(value ?? "").trim())) {
    const serial = Number(String(value).trim());
    if (serial <= 20000 || serial >= 70000) return { y: "", m: "", d: "", level: "", text: "" };
  }
  const dmy = s.match(/\b(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})\b/);
  if (dmy) {
    const y = dmy[3].length === 2 ? `20${dmy[3]}` : dmy[3];
    return { y, m: dmy[2].padStart(2, "0"), d: dmy[1].padStart(2, "0"), level: "day", text: s };
  }
  const my = s.match(/\b(\d{1,2})[\/\-.](\d{4})\b/);
  if (my) return { y: my[2], m: my[1].padStart(2, "0"), d: "", level: "month", text: s };
  const y = s.match(/\b(19|20)\d{2}\b/)?.[0];
  if (y) return { y, m: "", d: "", level: "year", text: s };
  return { y: "", m: "", d: "", level: "", text: s };
}
function validDateText(value = "") {
  const parsed = parseDateParts(value);
  return parsed.y ? parsed.text : "";
}
function parseAllDateParts(value = "") {
  if (/^\d+(\.0+)?$/.test(String(value ?? "").trim())) {
    const parsed = parseDateParts(value);
    return parsed.y ? [parsed] : [];
  }
  const s = excelDateText(value);
  const result = [];
  for (const m of s.matchAll(/\b(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})\b/g)) {
    const y = m[3].length === 2 ? `20${m[3]}` : m[3];
    result.push({ y, m: m[2].padStart(2, "0"), d: m[1].padStart(2, "0"), level: "day", text: m[0] });
  }
  if (result.length) return result;
  for (const m of s.matchAll(/\b(\d{1,2})[\/\-.](\d{4})\b/g)) {
    result.push({ y: m[2], m: m[1].padStart(2, "0"), d: "", level: "month", text: m[0] });
  }
  if (result.length) return result;
  for (const m of s.matchAll(/\b(19|20)\d{2}\b/g)) {
    result.push({ y: m[0], m: "", d: "", level: "year", text: m[0] });
  }
  return result;
}
function dateCompare(sourceValues, refValue) {
  const refs = parseAllDateParts(refValue);
  const sources = sourceValues.flatMap(parseAllDateParts).filter((x) => x.y);
  if (!refs.length) return { ok: false, note: "CSDL chuẩn chưa có ngày/thời điểm rõ" };
  for (const src of sources) {
    for (const ref of refs) {
      if (src.y === ref.y && src.m && ref.m && src.m === ref.m && src.d && ref.d && src.d === ref.d) return { ok: true, note: "khớp ngày" };
      if (src.y === ref.y && src.m && ref.m && src.m === ref.m && (src.level === "month" || ref.level === "month" || !src.d || !ref.d)) return { ok: true, note: "khớp tháng/năm" };
      if (src.y === ref.y && (src.level === "year" || ref.level === "year" || (!src.m && !ref.m))) return { ok: true, note: "khớp năm" };
    }
  }
  if (!sources.length) return { ok: false, note: "tab 200 chưa có thời gian để so" };
  return { ok: false, note: `200: ${sources.map((x) => x.text).filter(Boolean).join(" / ")}; CSDL: ${excelDateText(refValue)}` };
}
function typeFromText(s = "") {
  const n = normalize(s);
  if (n.includes("the gioi")) return "CSDL KLTG";
  if (n.includes("chau a")) return "CSDL KL CA";
  if (n.includes("toan cau")) return "CSDL KLVW";
  if (n.includes("viet nam")) return "CSDL KLVN";
  return "";
}
function refTypeFromSheet(sheet) {
  return sheet === "CSDL KLVN" ? "KỶ LỤC VIỆT NAM"
    : sheet === "CSDL KL CA" ? "KỶ LỤC CHÂU Á"
      : sheet === "CSDL KLTG" ? "KỶ LỤC THẾ GIỚI"
        : "KỶ LỤC NGƯỜI VIỆT TOÀN CẦU";
}
function statusStyle(status) {
  if (status.startsWith("KHỚP")) return 1;
  if (status.includes("THỜI GIAN")) return 2;
  if (status.includes("LỆCH")) return 3;
  if (status.includes("200 CÓ")) return 4;
  if (status.includes("CSDL CÓ")) return 5;
  return 0;
}

function refFromRow(sheetName, r) {
  let owner = "", title = "", date = "", id = "", link = "", province = "";
  if (sheetName === "CSDL KLVN") {
    id = r.values.get(3) || "";
    date = r.values.get(4) || "";
    title = r.values.get(5) || "";
    owner = r.values.get(6) || "";
    province = r.values.get(7) || "";
  } else if (sheetName === "CSDL KLVW") {
    id = r.values.get(4) || "";
    date = r.values.get(5) || "";
    title = r.values.get(6) || "";
    owner = r.values.get(7) || "";
    province = r.values.get(8) || "";
  } else {
    owner = r.values.get(2) || "";
    title = r.values.get(3) || "";
    id = r.values.get(4) || "";
    date = r.values.get(5) || r.values.get(4) || "";
    link = r.values.get(6) || "";
    province = r.values.get(7) || "";
  }
  if (!String(owner + title).trim()) return null;
  if (!owner && title.includes("\n")) {
    const lines = title.split(/\r?\n/).map((x) => x.trim()).filter(Boolean);
    if (lines.length > 1) {
      owner = lines[0];
      title = lines.slice(1).join(" ");
    }
  }
  return {
    sheet: sheetName,
    row: r.rowNumber,
    id: String(id).trim(),
    owner: String(owner).trim(),
    title: String(title).trim(),
    date: String(date).trim(),
    link: String(link).trim(),
    province: String(province).trim(),
    type: refTypeFromSheet(sheetName),
    ownerTokens: tokens(owner, true),
    titleTokens: tokens(title),
  };
}
function sourceFromRow(r) {
  const title = String(r.values.get(5) || "").trim();
  const gabTitle = String(r.values.get(14) || "").trim();
  const name = String(r.values.get(7) || "").trim();
  return {
    row: r.rowNumber,
    stt: String(r.values.get(1) || "").trim(),
    hoSo: String(r.values.get(2) || "").trim(),
    id: String(r.values.get(3) || "").trim(),
    dateOfficial: String(r.values.get(4) || "").trim(),
    title,
    profile: String(r.values.get(6) || "").trim(),
    name,
    birth: String(r.values.get(8) || "").trim(),
    province: String(r.values.get(9) || "").trim(),
    type: String(r.values.get(11) || "").trim(),
    otherTitle: String(r.values.get(12) || "").trim(),
    dateGab: String(r.values.get(13) || "").trim(),
    gabTitle,
    articleLink: String(r.values.get(15) || "").trim(),
    desc: String(r.values.get(16) || "").trim(),
    gabContent: String(r.values.get(20) || "").trim(),
    timeInContent: String(r.values.get(21) || "").trim(),
    review: String(r.values.get(23) || "").trim(),
    nameTokens: tokens(name, true),
    titleTokens: tokens([title, gabTitle].filter(Boolean).join(" ")),
  };
}

resetDir(outRoot);
resetDir(sourceDir);
resetDir(refDir);
resetDir(buildDir);
fs.copyFileSync(sourceXlsx, path.join(outRoot, "source.zip"));
fs.copyFileSync(refXlsx, path.join(outRoot, "ref.zip"));
ps(`Expand-Archive -LiteralPath '${path.join(outRoot, "source.zip").replace(/'/g, "''")}' -DestinationPath '${sourceDir.replace(/'/g, "''")}' -Force`);
ps(`Expand-Archive -LiteralPath '${path.join(outRoot, "ref.zip").replace(/'/g, "''")}' -DestinationPath '${refDir.replace(/'/g, "''")}' -Force`);

const sourceSheets = sheetMap(sourceDir);
const refSheets = sheetMap(refDir);
const sourceShared = sharedStrings(sourceDir);
const refShared = sharedStrings(refDir);
const sourceSheet = sourceSheets.get("200 KLG CHUẨN HÓA");
if (!sourceSheet) throw new Error("Không thấy tab 200 KLG CHUẨN HÓA");

const refs = [];
for (const name of ["CSDL KLVN", "CSDL KLVW", "CSDL KL CA", "CSDL KLTG"]) {
  const sheet = refSheets.get(name);
  if (!sheet) throw new Error(`Không thấy tab chuẩn ${name}`);
  for (const r of rowsFromSheet(refDir, sheet, refShared)) {
    if (r.rowNumber < 3) continue;
    const ref = refFromRow(name, r);
    if (ref) refs.push(ref);
  }
}
const sourceRows = rowsFromSheet(sourceDir, sourceSheet, sourceShared)
  .filter((r) => r.rowNumber >= 3)
  .map(sourceFromRow)
  .filter((r) => `${r.stt}${r.name}${r.title}${r.gabTitle}${r.articleLink}`.trim());

function bestMatch(src) {
  const wantedSheet = typeFromText(src.type || src.gabTitle || src.title);
  let best = null;
  for (const ref of refs) {
    const idHit = src.id && ref.id && normalize(src.id) === normalize(ref.id);
    const typeBonus = wantedSheet && ref.sheet === wantedSheet ? 0.12 : 0;
    const nameScore = score(src.nameTokens, ref.ownerTokens);
    const titleScore = score(src.titleTokens, ref.titleTokens);
    const idScore = idHit ? 1 : 0;
    const total = idScore * 0.35 + nameScore * 0.25 + titleScore * 0.40 + typeBonus;
    if (!best || total > best.total) best = { ref, total, nameScore, titleScore, idHit };
  }
  return best;
}

const matchedRefKeys = new Set();
const compareRows = [];
for (const src of sourceRows) {
  const match = bestMatch(src);
  const ref = match?.ref;
  let status = "200 CÓ - CSDL CHƯA THẤY";
  let issue = "Không tìm thấy bản ghi đủ chắc trong 4 tab chuẩn.";
  let action = "Rà lại tên KLG/tên kỷ lục. Nếu đây là dữ liệu mới chưa cập nhật vào CSDL chuẩn thì bổ sung vào tab chuẩn tương ứng; nếu không có căn cứ thì tách khỏi danh sách chuẩn.";
  let dateNote = "";
  if (ref && (match.idHit || (match.nameScore >= 0.72 && match.titleScore >= 0.45) || (match.nameScore >= 0.55 && match.titleScore >= 0.68))) {
    matchedRefKeys.add(`${ref.sheet}|${ref.row}`);
    const date = dateCompare([src.dateOfficial, src.dateGab, src.timeInContent], ref.date);
    dateNote = date.note;
    const titleOk = match.titleScore >= 0.70 || match.idHit;
    const nameOk = match.nameScore >= 0.72 || match.idHit;
    if (nameOk && titleOk && date.ok) {
      status = "KHỚP ĐỦ";
      issue = "Tên/chủ thể, nội dung kỷ lục và thời gian khớp với CSDL chuẩn.";
      action = "Giữ nguyên, ưu tiên dùng thông tin ở 200 nếu phần mô tả/link GAB đã hoàn chỉnh.";
    } else if (nameOk && titleOk && !date.ok) {
      status = "LỆCH THỜI GIAN";
      issue = `Bản ghi khớp tên và kỷ lục nhưng thời gian chưa khớp. ${date.note}`;
      action = `Chuẩn hóa thời gian theo CSDL: ${excelDateText(ref.date) || "chưa rõ"}; nếu tab 200 chỉ ghi tháng/năm thì xác nhận quy tắc nhập trước khi sửa.`;
    } else if (!nameOk) {
      status = "LỆCH TÊN/CHỦ THỂ";
      issue = `Nội dung gần với CSDL nhưng tên/chủ thể chưa đủ khớp (${Math.round(match.nameScore * 100)}%).`;
      action = `So lại cột G của 200 với chủ thể CSDL "${ref.owner}".`;
    } else {
      status = "LỆCH TIÊU ĐỀ KỶ LỤC";
      issue = `Tên Kỷ lục/GAB title chưa đủ khớp với CSDL (${Math.round(match.titleScore * 100)}%).`;
      action = `So lại cột E/N của 200 với tên chuẩn "${ref.title}".`;
    }
  }
  compareRows.push({
    status,
    style: statusStyle(status),
    row: src.row,
    stt: src.stt,
    name200: src.name,
    type200: src.type,
    id200: src.id,
    date200: [src.dateOfficial, src.dateGab, src.timeInContent].map(validDateText).filter(Boolean).join(" | "),
    title200: src.title || src.gabTitle,
    gabTitle200: src.gabTitle,
    link200: src.articleLink,
    refSheet: ref?.sheet || "",
    refRow: ref?.row || "",
    refName: ref?.owner || "",
    refType: ref?.type || "",
    refId: ref?.id || "",
    refDate: ref ? excelDateText(ref.date) : "",
    refTitle: ref?.title || "",
    nameScore: match ? `${Math.round(match.nameScore * 100)}%` : "",
    titleScore: match ? `${Math.round(match.titleScore * 100)}%` : "",
    issue,
    action,
    dateNote,
  });
}

const missingRows = refs
  .filter((ref) => !matchedRefKeys.has(`${ref.sheet}|${ref.row}`))
  .map((ref) => ({
    status: "CSDL CÓ - 200 THIẾU",
    style: 5,
    refSheet: ref.sheet,
    refRow: ref.row,
    refName: ref.owner,
    refType: ref.type,
    refId: ref.id,
    refDate: excelDateText(ref.date),
    refTitle: ref.title,
    refLink: ref.link,
    refProvince: ref.province,
    action: "Nếu thuộc phạm vi 200 KLG cần cập nhật, bổ sung bản ghi này vào tab 200/GAB; nếu không thuộc danh sách 200 thì đánh dấu ngoài phạm vi.",
  }));
const sourceNameTokens = sourceRows.map((src) => ({ src, tokens: src.nameTokens })).filter((x) => x.tokens.length);
const missingSameOwnerRows = missingRows.filter((row) => {
  const ref = refs.find((x) => x.sheet === row.refSheet && x.row === row.refRow);
  if (!ref) return false;
  return sourceNameTokens.some((x) => score(x.tokens, ref.ownerTokens) >= 0.72);
});

const counts = {
  total200: sourceRows.length,
  totalRefs: refs.length,
  ok: compareRows.filter((r) => r.status === "KHỚP ĐỦ").length,
  time: compareRows.filter((r) => r.status === "LỆCH THỜI GIAN").length,
  title: compareRows.filter((r) => r.status === "LỆCH TIÊU ĐỀ KỶ LỤC").length,
  name: compareRows.filter((r) => r.status === "LỆCH TÊN/CHỦ THỂ").length,
  newIn200: compareRows.filter((r) => r.status === "200 CÓ - CSDL CHƯA THẤY").length,
  missingIn200: missingRows.length,
  missingSameOwnerIn200: missingSameOwnerRows.length,
};

function cellXml(value, ref, style = 0) {
  const s = style ? ` s="${style}"` : "";
  return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${esc(value ?? "")}</t></is></c>`;
}
function sheetXml(rows, colWidths = [], freeze = true) {
  const maxCols = Math.max(...rows.map((r) => r.length), 1);
  const cols = colWidths.length
    ? `<cols>${colWidths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("")}</cols>`
    : "";
  const sheetViews = freeze ? `<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>` : "";
  const sheetData = rows.map((row, i) => {
    const r = i + 1;
    const cells = row.map((cell, j) => {
      const value = typeof cell === "object" && cell !== null ? cell.value : cell;
      const style = typeof cell === "object" && cell !== null ? cell.style || 0 : (i === 0 ? 6 : 0);
      return cellXml(value, `${colLetters(j + 1)}${r}`, style);
    }).join("");
    return `<row r="${r}">${cells}</row>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${sheetViews}<dimension ref="A1:${colLetters(maxCols)}${rows.length}"/>${cols}<sheetData>${sheetData}</sheetData><autoFilter ref="A1:${colLetters(maxCols)}${rows.length}"/></worksheet>`;
}
function writeWorkbook(sheets) {
  fs.mkdirSync(path.join(buildDir, "_rels"), { recursive: true });
  fs.mkdirSync(path.join(buildDir, "xl", "_rels"), { recursive: true });
  fs.mkdirSync(path.join(buildDir, "xl", "worksheets"), { recursive: true });
  fs.mkdirSync(path.join(buildDir, "docProps"), { recursive: true });
  const overrides = sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("");
  fs.writeFileSync(path.join(buildDir, "[Content_Types].xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${overrides}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`);
  fs.writeFileSync(path.join(buildDir, "_rels", ".rels"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
  fs.writeFileSync(path.join(buildDir, "xl", "workbook.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets></workbook>`);
  fs.writeFileSync(path.join(buildDir, "xl", "_rels", "workbook.xml.rels"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("")}<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
  fs.writeFileSync(path.join(buildDir, "xl", "styles.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/><color rgb="FFFFFFFF"/></font></fonts><fills count="8"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9EAD3"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF4CCCC"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFCFE2F3"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9D2E9"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="7"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="6" borderId="0" xfId="0" applyFill="1"/><xf numFmtId="0" fontId="1" fillId="7" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>`);
  for (const [i, s] of sheets.entries()) fs.writeFileSync(path.join(buildDir, "xl", "worksheets", `sheet${i + 1}.xml`), s.xml);
  fs.writeFileSync(path.join(buildDir, "docProps", "core.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>So sanh 200 KLG voi CSDL</dc:title></cp:coreProperties>`);
  fs.writeFileSync(path.join(buildDir, "docProps", "app.xml"), `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Codex</Application></Properties>`);
}

const summaryRows = [
  ["Hạng mục", "Số lượng", "Ghi chú"],
  ["Dòng có dữ liệu trong tab 200", counts.total200, "Tab 200 KLG CHUẨN HÓA, từ dòng 3 trở đi"],
  ["Bản ghi trong 4 tab CSDL chuẩn", counts.totalRefs, "CSDL KLVN, KLVW, KL CA, KLTG"],
  ["KHỚP ĐỦ", counts.ok, "Tô xanh"],
  ["LỆCH THỜI GIAN", counts.time, "Tô vàng"],
  ["LỆCH TIÊU ĐỀ KỶ LỤC", counts.title, "Tô đỏ"],
  ["LỆCH TÊN/CHỦ THỂ", counts.name, "Tô đỏ"],
  ["200 CÓ - CSDL CHƯA THẤY", counts.newIn200, "Tô xanh dương"],
  ["CSDL CÓ - 200 THIẾU", counts.missingIn200, "Tô tím, xem sheet CSDL_thieu_trong_200"],
  ["CSDL CÓ - 200 THIẾU cùng tên KLG trong tab 200", counts.missingSameOwnerIn200, "Bản lọc quan trọng hơn, xem sheet CSDL_thieu_cung_KLG"],
];
const compareHeader = ["Status", "Dòng 200", "STT", "Tên 200", "Danh vị 200", "Số XL 200", "Thời gian 200", "Tên kỷ lục 200", "Tiêu đề GAB 200", "Link bài viết 200", "Tab CSDL", "Dòng CSDL", "Tên CSDL", "Loại CSDL", "Số/Năm XL CSDL", "Thời gian CSDL", "Tên kỷ lục CSDL", "Điểm tên", "Điểm tiêu đề", "Chỗ khác / nhận định", "Cách xử lý"];
const compareData = [
  compareHeader.map((value) => ({ value, style: 6 })),
  ...compareRows.map((r) => [
    { value: r.status, style: r.style },
    r.row, r.stt, r.name200, r.type200, r.id200, r.date200, r.title200, r.gabTitle200, r.link200,
    r.refSheet, r.refRow, r.refName, r.refType, r.refId, r.refDate, r.refTitle, r.nameScore, r.titleScore, r.issue, r.action,
  ]),
];
const missingHeader = ["Status", "Tab CSDL", "Dòng CSDL", "Tên CSDL", "Loại", "Số/Năm XL", "Thời gian", "Tên kỷ lục", "Link", "Tỉnh/Thành", "Cách xử lý"];
const missingData = [
  missingHeader.map((value) => ({ value, style: 6 })),
  ...missingRows.map((r) => [
    { value: r.status, style: r.style },
    r.refSheet, r.refRow, r.refName, r.refType, r.refId, r.refDate, r.refTitle, r.refLink, r.refProvince, r.action,
  ]),
];
const missingSameOwnerData = [
  missingHeader.map((value) => ({ value, style: 6 })),
  ...missingSameOwnerRows.map((r) => [
    { value: r.status, style: r.style },
    r.refSheet, r.refRow, r.refName, r.refType, r.refId, r.refDate, r.refTitle, r.refLink, r.refProvince, "Ưu tiên rà sheet này trước: CSDL có thêm bản ghi cho KLG/đơn vị đang xuất hiện trong tab 200 nhưng chưa match với dòng nào của 200.",
  ]),
];

function sameOwnerRefsForName(name) {
  const nameTokens = tokens(name, true);
  return refs.filter((ref) => score(nameTokens, ref.ownerTokens) >= 0.72);
}
function sameOwnerMissingForName(name) {
  const nameTokens = tokens(name, true);
  return missingSameOwnerRows.filter((row) => {
    const ref = refs.find((x) => x.sheet === row.refSheet && x.row === row.refRow);
    return ref && score(nameTokens, ref.ownerTokens) >= 0.72;
  });
}
function duplicateKey(row) {
  return `${normalize(row.name200)}|${normalize(row.title200 || row.gabTitle200)}|${normalize(row.id200)}|${normalize(row.date200)}`;
}
const duplicateFirst = new Map();
const duplicateIssues = [];
for (const row of compareRows) {
  const key = duplicateKey(row);
  if (!normalize(row.title200 || row.gabTitle200)) continue;
  if (duplicateFirst.has(key)) {
    duplicateIssues.push({ ...row, duplicateOf: duplicateFirst.get(key) });
  } else {
    duplicateFirst.set(key, row.row);
  }
}

const oneHeader = [
  "KLG/đơn vị",
  "Loại vấn đề",
  "Kết luận nhìn nhanh",
  "Số dòng 200",
  "Số bản ghi CSDL cùng KLG",
  "Dòng 200",
  "STT 200",
  "Loại/Danh vị 200",
  "Số XL 200",
  "Thời gian 200",
  "Tên kỷ lục 200",
  "Link 200",
  "Dòng/Tab CSDL",
  "Loại CSDL",
  "Số XL CSDL",
  "Thời gian CSDL",
  "Tên kỷ lục CSDL",
  "Giá trị chuẩn/đề xuất",
  "Khác/thiếu/sai cụ thể",
  "Cách xử lý",
  "Màu trạng thái",
];
function oneRow(style, name, issueType, conclusion, count200, countRef, row200, stt, type200, id200, time200, title200, link200, csdlWhere, typeCsdl, idCsdl, timeCsdl, titleCsdl, standardValue, note, action) {
  return [
    { value: name, style },
    { value: issueType, style },
    { value: conclusion, style },
    count200,
    countRef,
    row200,
    stt,
    type200,
    id200,
    time200,
    title200,
    link200,
    csdlWhere,
    typeCsdl,
    idCsdl,
    timeCsdl,
    titleCsdl,
    standardValue,
    note,
    action,
    style === 2 ? "Vàng: lệch thời gian" : style === 3 ? "Đỏ: sai/lệch/trùng cần xử lý" : style === 4 ? "Xanh dương: 200 có thêm, CSDL thiếu" : style === 5 ? "Tím: CSDL có, 200 thiếu" : "",
  ];
}
function standardValueForCompare(r) {
  if (r.style === 2) return r.refDate ? `Thời gian chuẩn theo CSDL: ${r.refDate}` : "Cần xác minh thời gian chuẩn";
  if (r.style === 3 && r.status.includes("TIÊU")) return r.refTitle ? `Tên kỷ lục chuẩn theo CSDL: ${short(r.refTitle, 260)}` : "Cần xác minh tên kỷ lục chuẩn";
  if (r.style === 3 && r.status.includes("TÊN")) return r.refName ? `Tên/chủ thể chuẩn theo CSDL: ${r.refName}` : "Cần xác minh tên/chủ thể chuẩn";
  if (r.style === 4) return "Nếu research 200 đúng: bổ sung bản ghi này vào CSDL chuẩn";
  return "Cần xác minh";
}
function standardValueForMissing(r) {
  return "Nếu CSDL đúng và thuộc phạm vi 200: bổ sung dòng này vào tab 200/GAB; nếu 200 đúng: ghi chú CSDL dư/ngoài phạm vi";
}

const groupedNames = [...new Set(compareRows.map((r) => r.name200).filter(Boolean))].sort((a, b) => a.localeCompare(b, "vi"));
const oneSheetRows = [oneHeader.map((value) => ({ value, style: 6 }))];
oneSheetRows.push(oneRow(
  6,
  "TỔNG QUAN",
  "Nguồn đối chiếu",
  "Chỉ liệt kê dòng sai/thiếu/lệch/trùng. Dòng đã đủ thì không đưa vào sheet này.",
  counts.total200,
  "",
  "",
  "",
  "",
  "",
  "",
  "",
  `Nguồn trọng tâm: tab 200 KLG CHUẨN HÓA\n${sourceXlsx}`,
  "",
  "",
  "",
  "",
  "",
  `Nguồn chuẩn: 4 tab CSDL KLVN/KLVW/KL CA/KLTG\n${refXlsx}`,
  `Cần xử lý: ${counts.time + counts.title + counts.name + counts.newIn200 + counts.missingSameOwnerIn200 + duplicateIssues.length} mục chi tiết. Khớp đủ: ${counts.ok} dòng.`,
  "Ưu tiên đọc theo từng KLG. 'CSDL THIẾU THEO 200' nghĩa là 200 có thêm research; 'CSDL CÓ THÊM SO VỚI 200' nghĩa là CSDL có bản ghi chưa thấy trong 200.",
));

for (const name of groupedNames) {
  const group = compareRows.filter((r) => r.name200 === name);
  const missingForName = sameOwnerMissingForName(name);
  const matchedRefKeysForName = new Set(group.filter((r) => r.refSheet && r.refRow && r.style !== 4).map((r) => `${r.refSheet}|${r.refRow}`));
  for (const row of missingForName) matchedRefKeysForName.add(`${row.refSheet}|${row.refRow}`);
  const refCount = matchedRefKeysForName.size;
  const extraForName = group.filter((r) => r.style === 4);
  const badForName = group.filter((r) => [2, 3].includes(r.style));
  const dupForName = duplicateIssues.filter((r) => r.name200 === name);
  const countDiff = group.length !== refCount;
  if (!countDiff && !missingForName.length && !extraForName.length && !badForName.length && !dupForName.length) continue;
  const parts = [];
  if (group.length > refCount) parts.push(`tab 200 nhiều hơn CSDL ${group.length - refCount} dòng`);
  if (refCount > group.length) parts.push(`CSDL nhiều hơn tab 200 ${refCount - group.length} bản ghi`);
  if (extraForName.length) parts.push(`${extraForName.length} dòng 200 chưa thấy chắc trong CSDL`);
  if (missingForName.length) parts.push(`${missingForName.length} bản ghi CSDL chưa có trong 200`);
  if (badForName.length) parts.push(`${badForName.length} dòng lệch thông tin`);
  if (dupForName.length) parts.push(`${dupForName.length} dòng có khả năng trùng trong 200`);
  oneSheetRows.push(oneRow(
    6,
    name,
    "TỔNG HỢP KLG",
    parts.join("; "),
    group.length,
    refCount,
    "",
    "",
    "",
    "",
    "",
    "",
    `Tab 200 có ${group.length} dòng/thành tựu cho KLG này.`,
    "",
    "",
    "",
    "",
    "",
    `4 tab CSDL có ${refCount} bản ghi đã match/còn thiếu cho KLG này.`,
    "Dòng tổng hợp để nhìn nhanh chênh lệch số lượng trước khi xem chi tiết bên dưới.",
    "Rà các dòng chi tiết ngay dưới tên KLG này.",
  ));
  for (const r of extraForName) {
    oneSheetRows.push(oneRow(
      4,
      name,
      "CSDL THIẾU THEO 200",
      "Tab 200 có thành tựu này nhưng 4 tab CSDL chưa có bản ghi đủ chắc. Nếu research 200 đúng thì cần bổ sung CSDL.",
      group.length,
      refCount,
      r.row,
      r.stt,
      r.type200,
      r.id200,
      r.date200,
      short(r.title200 || r.gabTitle200, 260),
      short(r.link200, 180),
      r.refSheet ? `${r.refSheet}!${r.refRow}` : "",
      r.refType,
      r.refId,
      r.refDate,
      short(r.refTitle, 260),
      standardValueForCompare(r),
      r.issue,
      "Xác minh nguồn/link ở tab 200. Nếu đúng, bổ sung vào tab CSDL chuẩn tương ứng; nếu không đúng thì sửa/xóa dòng 200.",
    ));
  }
  for (const r of badForName) {
    const issueType = r.style === 2 ? "LỆCH THỜI GIAN" : r.status;
    oneSheetRows.push(oneRow(
      r.style,
      name,
      issueType,
      r.style === 2 ? "Cùng kỷ lục/cùng KLG nhưng thời gian ở 200 và CSDL chưa khớp." : "Cùng nhóm dữ liệu nhưng tên/chủ thể/tiêu đề chưa khớp.",
      group.length,
      refCount,
      r.row,
      r.stt,
      r.type200,
      r.id200,
      r.date200,
      short(r.title200 || r.gabTitle200, 260),
      short(r.link200, 180),
      r.refSheet ? `${r.refSheet}!${r.refRow}` : "",
      r.refType,
      r.refId,
      r.refDate,
      short(r.refTitle, 260),
      standardValueForCompare(r),
      r.issue,
      r.action,
    ));
  }
  for (const r of dupForName) {
    oneSheetRows.push(oneRow(
      3,
      name,
      "TRÙNG TRONG TAB 200",
      `Dòng này có nội dung giống/gần giống dòng 200 số ${r.duplicateOf}.`,
      group.length,
      refCount,
      r.row,
      r.stt,
      "",
      "",
      r.type200,
      r.id200,
      r.date200,
      short(r.title200 || r.gabTitle200, 260),
      short(r.link200, 180),
      "",
      "",
      "",
      "",
      `Dòng 200 trùng với dòng ${r.duplicateOf}`,
      `Có khả năng trùng thành tựu trong chính tab 200 với dòng ${r.duplicateOf}.`,
      "Kiểm tra 2 dòng trong tab 200. Nếu là cùng một thành tựu thì gộp/xóa một dòng; nếu là 2 thành tựu khác nhau thì bổ sung chi tiết phân biệt.",
    ));
  }
  for (const r of missingForName) {
    oneSheetRows.push(oneRow(
      5,
      name,
      "CSDL CÓ THÊM SO VỚI 200",
      "4 tab CSDL có bản ghi cùng KLG nhưng tab 200 chưa thấy dòng tương ứng.",
      group.length,
      refCount,
      "",
      "",
      "",
      "",
      "",
      "",
      "",
      `${r.refSheet}!${r.refRow}`,
      r.refType,
      r.refId,
      r.refDate,
      short(r.refTitle, 260),
      standardValueForMissing(r),
      "CSDL đang có thêm bản ghi này so với tab 200. Có thể tab 200 thiếu, CSDL dư/ngoài phạm vi, hoặc 2 nguồn đang tách/gộp thành tựu khác nhau.",
      "Xác minh phạm vi theo research 200. Nếu CSDL đúng và thuộc KLG trong danh sách 200 thì bổ sung vào tab 200/GAB; nếu 200 đúng thì ghi chú CSDL dư/ngoài phạm vi.",
    ));
  }
}
const oneWidths = [28, 24, 48, 12, 18, 10, 10, 18, 18, 18, 58, 44, 22, 18, 18, 18, 58, 54, 68, 68, 26];
writeWorkbook([
  { name: "Sai_thieu_can_xu_ly", xml: sheetXml(oneSheetRows, oneWidths) },
]);

const zipOut = path.join(outRoot, "compare.zip");
if (fs.existsSync(zipOut)) fs.rmSync(zipOut, { force: true });
if (fs.existsSync(outputXlsx)) fs.rmSync(outputXlsx, { force: true });
ps(`Add-Type -AssemblyName System.IO.Compression.FileSystem; [System.IO.Compression.ZipFile]::CreateFromDirectory('${buildDir.replace(/'/g, "''")}', '${zipOut.replace(/'/g, "''")}')`);
fs.copyFileSync(zipOut, outputXlsx);
console.log(JSON.stringify({ outputXlsx, counts }, null, 2));
