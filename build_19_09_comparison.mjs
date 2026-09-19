import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';

const root = process.cwd();
const dir = path.join(root, 'outputs', 'compare_19_09_2026');
const source = path.join(root, 'Dulieutruyxuat', '19_09_2026.xlsx');
const read = (p) => fs.readFileSync(p, 'utf8');
const write = (p, v) => fs.writeFileSync(p, v, 'utf8');
const esc = (v = '') => String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const dec = (v = '') => String(v).replace(/&#x([0-9a-f]+);/gi, (_, x) => String.fromCodePoint(parseInt(x, 16))).replace(/&#(\d+);/g, (_, x) => String.fromCodePoint(Number(x))).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const colNo = (ref) => { let n = 0; for (const c of String(ref).match(/^[A-Z]+/)?.[0] ?? '') n = n * 26 + c.charCodeAt(0) - 64; return n; };
const colName = (n) => { let s = ''; while (n) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };
const norm = (v = '') => String(v).toLocaleLowerCase('vi').normalize('NFD').replace(/\p{Diacritic}/gu, '').replace(/đ/g, 'd').replace(/\b(xac lap|ky luc|viet nam|chau a|the gioi|dong duong|klvn|klca|kltg|ky luc gia|klg|ong|ba|anh|chi|gs|pgs|ts|ths|bs|nghe nhan|nghe si|hoa si|tien si|thac si)\b/g, ' ').replace(/[^a-z0-9]+/g, ' ').replace(/\s+/g, ' ').trim();
const gabKey = (v = '') => String(v).match(/gab\.world\/vi\/bank\/([^/?#]+)/i)?.[1]?.toLowerCase() ?? '';
const validUrl = (v = '') => /^https?:\/\//i.test(String(v).trim());
const tokenScore = (a, b) => {
  const x = new Set(norm(a).split(' ').filter(Boolean));
  const y = new Set(norm(b).split(' ').filter(Boolean));
  if (!x.size || !y.size) return 0;
  let hits = 0; for (const t of x) if (y.has(t)) hits++;
  return (2 * hits) / (x.size + y.size);
};
const excelDate = (v = '') => {
  const s = String(v).trim();
  if (!s) return '';
  if (/^\d+(\.0+)?$/.test(s)) {
    const n = Number(s);
    if (n > 20000 && n < 70000) {
      const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
      return `${String(d.getUTCDate()).padStart(2, '0')}/${String(d.getUTCMonth() + 1).padStart(2, '0')}/${d.getUTCFullYear()}`;
    }
    if (/^(19|20)\d{2}$/.test(s)) return s;
  }
  return s.replace(/^Năm\s+/i, '');
};
const dateYear = (v = '') => String(v).match(/(?:19|20)\d{2}/)?.[0] ?? '';

const sharedPath = path.join(dir, 'xl/sharedStrings.xml');
const shared = fs.existsSync(sharedPath) ? [...read(sharedPath).matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map((m) => [...m[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((x) => dec(x[1])).join('')) : [];
function rows(sheetFile) {
  const result = [];
  const xml = read(path.join(dir, 'xl/worksheets', sheetFile));
  for (const m of xml.matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>/g)) {
    const row = Number(m[1].match(/\br="(\d+)"/)?.[1]);
    const values = new Map();
    for (const c of m[2].matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const attrs = c[1] || '', ref = attrs.match(/\br="([^"]+)"/)?.[1];
      if (!ref) continue;
      const type = attrs.match(/\bt="([^"]+)"/)?.[1] || '', body = c[2] || '';
      let value = '';
      if (type === 's') value = shared[Number(body.match(/<v>([\s\S]*?)<\/v>/)?.[1])] ?? '';
      else if (type === 'inlineStr') value = [...body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((x) => dec(x[1])).join('');
      else value = dec(body.match(/<v>([\s\S]*?)<\/v>/)?.[1] ?? '');
      values.set(colNo(ref), value);
    }
    result.push({ row, values });
  }
  return result;
}

const gabRaw = rows('sheet1.xml').filter((r) => r.row > 1).map((r) => ({
  sourceRow: r.row,
  recordId: String(r.values.get(1) ?? '').trim(),
  gab: String(r.values.get(2) ?? '').trim(),
  name: String(r.values.get(3) ?? '').trim(),
  time: excelDate(r.values.get(4) ?? ''),
  title: String(r.values.get(5) ?? '').trim(),
  url: String(r.values.get(6) ?? '').trim(),
  description: String(r.values.get(7) ?? '').trim(),
})).filter((r) => r.name && r.title);

const gabDedup = new Map();
for (const r of gabRaw) {
  const key = r.recordId || `${gabKey(r.gab)}|${norm(r.title)}|${r.time}`;
  if (!gabDedup.has(key)) gabDedup.set(key, r);
}
const gab = [...gabDedup.values()];

const standard = rows('sheet2.xml').filter((r) => r.row > 2).map((r) => ({
  sourceRow: r.row,
  stt: String(r.values.get(1) ?? '').replace(/\.0$/, ''),
  profile: String(r.values.get(2) ?? '').trim(),
  cert: String(r.values.get(3) ?? '').trim(),
  officialDate: excelDate(r.values.get(4) ?? ''),
  recordTitle: String(r.values.get(5) ?? '').trim(),
  gab: String(r.values.get(6) ?? '').trim(),
  name: String(r.values.get(7) ?? '').trim(),
  birth: String(r.values.get(8) ?? '').replace(/\.0$/, ''),
  province: String(r.values.get(9) ?? '').trim(),
  recordType: String(r.values.get(11) ?? '').trim(),
  otherTitle: String(r.values.get(12) ?? '').trim(),
  gabTime: excelDate(r.values.get(13) ?? ''),
  gabTitle: String(r.values.get(14) ?? '').trim(),
  article: String(r.values.get(15) ?? '').trim(),
  description: String(r.values.get(16) ?? '').trim(),
  note: String(r.values.get(19) ?? '').trim(),
  verified: String(r.values.get(20) ?? '').trim(),
})).filter((r) => r.name && (r.recordTitle || r.gabTitle || r.verified));

const gabByPerson = new Map();
for (const r of gab) {
  const keys = [gabKey(r.gab), norm(r.name)].filter(Boolean);
  for (const key of keys) { if (!gabByPerson.has(key)) gabByPerson.set(key, []); gabByPerson.get(key).push(r); }
}
const standardPeople = new Map();
for (const r of standard) {
  const key = gabKey(r.gab) || norm(r.name);
  if (!standardPeople.has(key)) standardPeople.set(key, { name: r.name, gab: r.gab, rows: [] });
  standardPeople.get(key).rows.push(r);
}

function achievementScore(s, g) {
  if (validUrl(s.article) && validUrl(g.url) && s.article.replace(/\/$/, '').toLowerCase() === g.url.replace(/\/$/, '').toLowerCase()) return 1;
  const st = s.gabTitle || s.recordTitle;
  const nt = norm(st), ng = norm(g.title);
  if (nt && nt === ng) return 0.98;
  if (nt && ng && (nt.includes(ng) || ng.includes(nt))) return 0.88;
  return tokenScore(st, g.title);
}

const output = [];
for (const person of standardPeople.values()) {
  const key = gabKey(person.gab);
  const candidates = [...new Map([...(key ? gabByPerson.get(key) ?? [] : []), ...(gabByPerson.get(norm(person.name)) ?? [])].map((x) => [x.recordId || `${x.sourceRow}`, x])).values()];
  const used = new Set();
  for (const s of person.rows) {
    let best = null;
    for (const g of candidates) {
      if (used.has(g.sourceRow)) continue;
      const score = achievementScore(s, g);
      if (!best || score > best.score) best = { g, score };
    }
    if (!best || best.score < 0.62) {
      const status = candidates.length ? 'GAB THIẾU THÀNH TỰU' : 'CHƯA CÓ HỒ SƠ TRÊN GAB';
      output.push({ status, issues: candidates.length ? 'Không tìm thấy thành tựu tương ứng trong dữ liệu GAB xuất về.' : 'Không tìm thấy người theo link GAB hoặc tên chuẩn hóa.', s, g: null, score: best?.score ?? 0 });
      continue;
    }
    used.add(best.g.sourceRow);
    const g = best.g, issues = [];
    const sTime = s.gabTime || s.officialDate;
    if (sTime && !g.time) issues.push('GAB thiếu thời gian');
    else if (sTime && g.time && dateYear(sTime) && dateYear(g.time) && dateYear(sTime) !== dateYear(g.time)) issues.push(`Lệch thời gian: chuẩn ${sTime}, GAB ${g.time}`);
    if ((s.gabTitle || s.recordTitle) && best.score < 0.8) issues.push('Tiêu đề chưa đồng nhất');
    if (validUrl(s.article) && !validUrl(g.url)) issues.push('GAB thiếu link bài');
    else if (validUrl(s.article) && validUrl(g.url) && s.article.replace(/\/$/, '').toLowerCase() !== g.url.replace(/\/$/, '').toLowerCase() && best.score < 0.9) issues.push('Link bài khác');
    if (s.description && !g.description) issues.push('GAB thiếu mô tả');
    const status = issues.length ? 'SAI/CHƯA ĐỒNG NHẤT' : 'ĐỦ - KHỚP';
    output.push({ status, issues: issues.join('; ') || 'Các trường chính đã khớp.', s, g, score: best.score });
  }
  for (const g of candidates.filter((x) => !used.has(x.sourceRow))) {
    output.push({ status: 'GAB CÓ THÊM - 200 CHƯA CÓ', issues: 'Thành tựu có trong dữ liệu GAB nhưng chưa tìm thấy dòng tương ứng trong tab 200; cần kiểm tra để cập nhật danh sách chuẩn.', s: { name: person.name, gab: person.gab }, g, score: 0 });
  }
}

const order = new Map([['GAB THIẾU THÀNH TỰU', 1], ['CHƯA CÓ HỒ SƠ TRÊN GAB', 2], ['SAI/CHƯA ĐỒNG NHẤT', 3], ['GAB CÓ THÊM - 200 CHƯA CÓ', 4], ['ĐỦ - KHỚP', 5]]);
output.sort((a, b) => (order.get(a.status) ?? 9) - (order.get(b.status) ?? 9) || norm(a.s.name || a.g?.name).localeCompare(norm(b.s.name || b.g?.name)) || (a.s.sourceRow ?? 0) - (b.s.sourceRow ?? 0));
const counts = Object.fromEntries([...order.keys()].map((k) => [k, output.filter((x) => x.status === k).length]));
const unique200 = standardPeople.size;
const matchedPeople = new Set(output.filter((x) => x.g).map((x) => gabKey(x.s.gab) || norm(x.s.name))).size;
const statusNames = [...order.keys()];
const personSummaries = [...standardPeople.values()].map((person) => {
  const key = gabKey(person.gab) || norm(person.name);
  const items = output.filter((x) => (gabKey(x.s?.gab || x.g?.gab) || norm(x.s?.name || x.g?.name)) === key);
  const compared = items.filter((x) => x.s?.sourceRow);
  const present = compared.filter((x) => x.g);
  const missing = compared.filter((x) => !x.g);
  const extras = items.filter((x) => x.status === statusNames[3]);
  const list = (rows, pick) => rows.map((x, i) => `${i + 1}. ${pick(x)}`).join('\n');
  return {
    name: person.name,
    gab: person.gab,
    standardCount: person.rows.length,
    presentCount: present.length,
    missingCount: missing.length,
    extraCount: extras.length,
    status: missing.length ? `THIẾU ${missing.length} THÀNH TỰU TRÊN GAB` : (extras.length ? `ĐÃ ĐỦ CHUẨN; GAB CÓ THÊM ${extras.length}` : 'ĐỦ - KHỚP'),
    presentTitles: list(present, (x) => x.g?.title || x.s?.gabTitle || x.s?.recordTitle || ''),
    missingTitles: list(missing, (x) => x.s?.gabTitle || x.s?.recordTitle || ''),
    extraTitles: list(extras, (x) => x.g?.title || ''),
  };
}).sort((a, b) => norm(a.name).localeCompare(norm(b.name)));

let styles = read(path.join(dir, 'xl/styles.xml'));
function append(xml, tag, item) {
  const re = new RegExp(`<${tag}\\b([^>]*)count="(\\d+)"([^>]*)>([\\s\\S]*?)<\\/${tag}>`);
  const m = xml.match(re); if (!m) throw new Error(`Không tìm thấy ${tag}`);
  return xml.replace(re, `<${tag}${m[1]}count="${Number(m[2]) + 1}"${m[3]}>${m[4]}${item}</${tag}>`);
}
const fontStart = Number(styles.match(/<fonts\b[^>]*count="(\d+)"/)?.[1] ?? 0);
styles = append(styles, 'fonts', '<font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Arial"/></font>');
const fillStart = Number(styles.match(/<fills\b[^>]*count="(\d+)"/)?.[1] ?? 0);
for (const color of ['FF17365D', 'FFD9EAF7', 'FFF4CCCC', 'FFFFE5B4', 'FFDDEBF7', 'FFE2F0D9', 'FFF2F2F2']) styles = append(styles, 'fills', `<fill><patternFill patternType="solid"><fgColor rgb="${color}"/><bgColor indexed="64"/></patternFill></fill>`);
const xfStart = Number(styles.match(/<cellXfs\b[^>]*count="(\d+)"/)?.[1] ?? 0);
const xfs = [
  `<xf numFmtId="0" fontId="${fontStart}" fillId="${fillStart}" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" horizontal="center" wrapText="1"/></xf>`,
  `<xf numFmtId="0" fontId="0" fillId="${fillStart + 1}" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>`,
  `<xf numFmtId="0" fontId="0" fillId="${fillStart + 2}" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>`,
  `<xf numFmtId="0" fontId="0" fillId="${fillStart + 3}" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>`,
  `<xf numFmtId="0" fontId="0" fillId="${fillStart + 4}" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>`,
  `<xf numFmtId="0" fontId="0" fillId="${fillStart + 5}" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>`,
  `<xf numFmtId="0" fontId="0" fillId="${fillStart + 6}" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>`,
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>',
];
for (const xf of xfs) styles = append(styles, 'cellXfs', xf);
write(path.join(dir, 'xl/styles.xml'), styles);

const cell = (v, ref, style) => `<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
const makeRow = (vals, n, style, height = 48) => `<row r="${n}" ht="${height}" customHeight="1">${vals.map((v, i) => cell(v, `${colName(i + 1)}${n}`, Array.isArray(style) ? style[i] : style)).join('')}</row>`;
const makeCells = (vals, n, startCol, style) => vals.map((v, i) => cell(v, `${colName(startCol + i)}${n}`, Array.isArray(style) ? style[i] : style)).join('');
const table = [];
table.push(makeRow(['ĐỐI CHIẾU DỮ LIỆU GAB VỚI DANH SÁCH 200 NGƯỜI ĐÃ CHUẨN'], 1, xfStart, 32));
table.push(makeRow([`Phạm vi: ${unique200} KLG trong danh sách chuẩn | ${standard.length} dòng thành tựu chuẩn | ${gab.length} thành tựu GAB sau khi loại ${gabRaw.length - gab.length} dòng trùng | ${matchedPeople} KLG đã ghép được với GAB.`], 2, xfStart + 1, 28));
table.push(makeRow([`GAB thiếu: ${counts['GAB THIẾU THÀNH TỰU']} | Chưa có hồ sơ GAB: ${counts['CHƯA CÓ HỒ SƠ TRÊN GAB']} | Sai/chưa đồng nhất: ${counts['SAI/CHƯA ĐỒNG NHẤT']} | GAB có thêm: ${counts['GAB CÓ THÊM - 200 CHƯA CÓ']} | Đủ-khớp: ${counts['ĐỦ - KHỚP']}`], 3, xfStart + 1, 28));
table.push(makeRow(['Ưu tiên xử lý các dòng màu đỏ/cam. Dòng xanh dương là thành tựu mới phía GAB có thể cần bổ sung ngược vào danh sách 200; dòng xanh lá đã khớp các trường chính.'], 4, xfStart + 1, 28));
table.push('<row r="5" ht="8" customHeight="1"/>');
const headers = ['STT', 'TRẠNG THÁI', 'THIẾU / SAI / CŨ', 'HỌ TÊN CHUẨN', 'LINK HỒ SƠ GAB', 'LOẠI KỶ LỤC', 'SỐ XÁC LẬP', 'NGÀY CHUẨN', 'NGÀY GAB', 'TIÊU ĐỀ CHUẨN', 'TIÊU ĐỀ GAB', 'LINK BÀI CHUẨN', 'LINK BÀI GAB', 'MÔ TẢ CHUẨN', 'MÔ TẢ GAB', 'THÀNH TỰU XÁC THỰC / GHI CHÚ', 'ĐỘ KHỚP', 'DÒNG TAB 200', 'DÒNG TAB GAB'];
table.push(makeRow(headers, 6, xfStart, 58));
const statusStyle = {
  'GAB THIẾU THÀNH TỰU': xfStart + 2,
  'CHƯA CÓ HỒ SƠ TRÊN GAB': xfStart + 2,
  'SAI/CHƯA ĐỒNG NHẤT': xfStart + 3,
  'GAB CÓ THÊM - 200 CHƯA CÓ': xfStart + 4,
  'ĐỦ - KHỚP': xfStart + 5,
};
output.forEach((r, i) => {
  const s = r.s || {}, g = r.g || {};
  const vals = [i + 1, r.status, r.issues, s.name || g.name || '', s.gab || g.gab || '', s.recordType || '', s.cert || '', s.gabTime || s.officialDate || '', g.time || '', s.gabTitle || s.recordTitle || '', g.title || '', s.article || '', g.url || '', s.description || '', g.description || '', s.verified || s.note || '', r.score ? `${Math.round(r.score * 100)}%` : '', s.sourceRow || '', g.sourceRow || ''];
  const base = statusStyle[r.status] ?? xfStart + 7;
  table.push(makeRow(vals, i + 7, vals.map((_, c) => c === 1 || c === 2 ? base : (c >= 8 && c <= 14 ? xfStart + 6 : xfStart + 7)), 64));
});
const summaryHeaders = ['KỶ LỤC GIA', 'LINK HỒ SƠ GAB', 'TỔNG CHUẨN', 'ĐÃ CÓ / TRÙNG', 'CÒN THIẾU', 'GAB CÓ THÊM', 'TÌNH TRẠNG KLG', 'CỤ THỂ ĐÃ CÓ / TRÙNG', 'CỤ THỂ CÒN THIẾU', 'CỤ THỂ GAB CÓ THÊM'];
const summaryRows = new Map();
summaryRows.set(1, makeCells(['TỔNG HỢP THEO KỶ LỤC GIA'], 1, 21, xfStart));
summaryRows.set(2, makeCells(['Mỗi KLG một dòng: xem ngay thành tựu nào đã có/trùng, còn thiếu và dữ liệu GAB có thêm.'], 2, 21, xfStart + 1));
summaryRows.set(6, makeCells(summaryHeaders, 6, 21, xfStart));
personSummaries.forEach((p, i) => {
  const n = i + 7;
  const style = p.missingCount ? xfStart + 2 : (p.extraCount ? xfStart + 4 : xfStart + 5);
  summaryRows.set(n, makeCells([p.name, p.gab, p.standardCount, p.presentCount, p.missingCount, p.extraCount, p.status, p.presentTitles, p.missingTitles, p.extraTitles], n, 21, [style, xfStart + 7, style, style, style, style, style, xfStart + 7, xfStart + 7, xfStart + 7]));
});
for (const [n, cells] of summaryRows) {
  const rowIndex = table.findIndex((xml) => xml.startsWith(`<row r="${n}"`));
  if (rowIndex >= 0) table[rowIndex] = table[rowIndex].replace('</row>', `${cells}</row>`);
  else table.push(`<row r="${n}" ht="64" customHeight="1">${cells}</row>`);
}
table.sort((a, b) => Number(a.match(/<row r="(\d+)"/)?.[1]) - Number(b.match(/<row r="(\d+)"/)?.[1]));
const lastDetail = output.length + 6;
const last = Math.max(lastDetail, personSummaries.length + 6);
const widths = [7, 27, 48, 28, 47, 24, 18, 16, 16, 58, 58, 50, 50, 70, 70, 66, 11, 13, 13, 3, 28, 47, 12, 14, 12, 14, 34, 72, 72, 72];
const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:AD${last}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="6" topLeftCell="A7" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>${widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols><sheetData>${table.join('')}</sheetData><autoFilter ref="A6:S${lastDetail}"/><mergeCells count="6"><mergeCell ref="A1:S1"/><mergeCell ref="A2:S2"/><mergeCell ref="A3:S3"/><mergeCell ref="A4:S4"/><mergeCell ref="U1:AD1"/><mergeCell ref="U2:AD2"/></mergeCells></worksheet>`;
write(path.join(dir, 'xl/worksheets/sheet3.xml'), sheet);

let wb = read(path.join(dir, 'xl/workbook.xml'));
wb = wb.replace(/<sheet[^>]*name="ĐỐI CHIẾU GAB - 200"[^>]*\/>/g, '').replace('</sheets>', '<sheet sheetId="3" name="ĐỐI CHIẾU GAB - 200" state="visible" r:id="rId3"/></sheets>');
write(path.join(dir, 'xl/workbook.xml'), wb);
let rels = read(path.join(dir, 'xl/_rels/workbook.xml.rels'));
rels = rels.replace(/<Relationship[^>]*Id="rId3"[^>]*\/>/g, '').replace('</Relationships>', '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/></Relationships>');
write(path.join(dir, 'xl/_rels/workbook.xml.rels'), rels);
let types = read(path.join(dir, '[Content_Types].xml'));
if (!types.includes('/xl/worksheets/sheet3.xml')) types = types.replace('</Types>', '<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
write(path.join(dir, '[Content_Types].xml'), types);

const zip = path.join(root, 'outputs', 'compare_19_09_2026_result.zip');
if (fs.existsSync(zip)) fs.rmSync(zip, { force: true });
execFileSync('powershell', ['-NoProfile', '-Command', `Add-Type -AssemblyName System.IO.Compression.FileSystem; [IO.Compression.ZipFile]::CreateFromDirectory('${dir.replace(/'/g, "''")}','${zip.replace(/'/g, "''")}')`]);
fs.copyFileSync(zip, source);
console.log(JSON.stringify({ source, gabRaw: gabRaw.length, gabDeduplicated: gab.length, duplicateGabRows: gabRaw.length - gab.length, standardRows: standard.length, unique200, matchedPeople, outputRows: output.length, counts }, null, 2));
