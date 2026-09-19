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
const personKey = (name = '', gab = '') => norm(name) || gabKey(gab);
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
  avatar: String(r.values.get(10) ?? '').trim(),
  recordType: String(r.values.get(11) ?? '').trim(),
  otherTitle: String(r.values.get(12) ?? '').trim(),
  gabTime: excelDate(r.values.get(13) ?? ''),
  gabTitle: String(r.values.get(14) ?? '').trim(),
  article: String(r.values.get(15) ?? '').trim(),
  description: String(r.values.get(16) ?? '').trim(),
  image: String(r.values.get(17) ?? '').trim(),
  youtube: String(r.values.get(18) ?? '').trim(),
  social: String(r.values.get(19) ?? '').trim(),
  verified: String(r.values.get(20) ?? '').trim(),
})).filter((r) => r.name && (r.recordTitle || r.gabTitle || r.verified));

const gabByPerson = new Map();
for (const r of gab) {
  const keys = [gabKey(r.gab), norm(r.name)].filter(Boolean);
  for (const key of keys) { if (!gabByPerson.has(key)) gabByPerson.set(key, []); gabByPerson.get(key).push(r); }
}
const standardPeople = new Map();
for (const r of standard) {
  const key = norm(r.name) || gabKey(r.gab);
  if (!standardPeople.has(key)) standardPeople.set(key, { name: r.name, gab: r.gab, gabLinks: new Set(), rows: [] });
  const person = standardPeople.get(key);
  if (r.gab) person.gabLinks.add(r.gab);
  if (!person.gab && r.gab) person.gab = r.gab;
  person.rows.push(r);
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
  const profileCandidates = [...person.gabLinks].flatMap((link) => gabByPerson.get(gabKey(link)) ?? []);
  const candidates = [...new Map([...profileCandidates, ...(gabByPerson.get(norm(person.name)) ?? [])].map((x) => [x.recordId || `${x.sourceRow}`, x])).values()];
  const used = new Set();
  for (const s of person.rows) {
    let best = null;
    for (const g of candidates) {
      if (used.has(g.sourceRow)) continue;
      const score = achievementScore(s, g);
      if (!best || score > best.score) best = { g, score };
    }
    if (!best) {
      const status = candidates.length ? 'GAB THIẾU THÀNH TỰU' : 'CHƯA CÓ HỒ SƠ TRÊN GAB';
      output.push({ status, issues: candidates.length ? 'Không tìm thấy thành tựu tương ứng trong dữ liệu GAB xuất về.' : 'Không tìm thấy người theo link GAB hoặc tên chuẩn hóa.', s, g: null, profileGab: candidates[0]?.gab || '', score: best?.score ?? 0 });
      continue;
    }
    used.add(best.g.sourceRow);
    const g = best.g, issues = [];
    if (!validUrl(s.gab) && validUrl(g.gab)) issues.push('Tab chuẩn thiếu link hồ sơ GAB nhưng đã tìm thấy theo tên');
    else if (validUrl(s.gab) && validUrl(g.gab) && gabKey(s.gab) !== gabKey(g.gab)) issues.push(`Link hồ sơ chuẩn khác gabId: chuẩn ${s.gab}, GAB ${g.gab}`);
    if (norm(s.name) && norm(g.name) && norm(s.name) !== norm(g.name)) issues.push(`Tên khác nhau: chuẩn "${s.name}", GAB "${g.name}"`);
    const sTime = s.gabTime || s.officialDate;
    if (sTime && !g.time) issues.push('GAB thiếu thời gian');
    else if (sTime && g.time && dateYear(sTime) && dateYear(g.time) && dateYear(sTime) !== dateYear(g.time)) issues.push(`Lệch thời gian: chuẩn ${sTime}, GAB ${g.time}`);
    if ((s.gabTitle || s.recordTitle) && best.score < 0.8) issues.push('Tiêu đề chưa đồng nhất');
    if (best.score < 0.62) issues.push(`Cần xác minh ghép: độ khớp tiêu đề thấp (${Math.round(best.score * 100)}%) nhưng record thuộc đúng hồ sơ KLG`);
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
const matchedPeople = new Set(output.filter((x) => x.g).map((x) => personKey(x.s.name || x.g?.name, x.s.gab || x.g?.gab))).size;
const statusNames = [...order.keys()];
const personSummaries = [...standardPeople.values()].map((person) => {
  const key = personKey(person.name, person.gab);
  const items = output.filter((x) => personKey(x.s?.name || x.g?.name, x.s?.gab || x.g?.gab) === key);
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
    status: missing.length ? `THIẾU ${missing.length} THÀNH TỰU TRÊN GAB` : (extras.length ? `KHÔNG THIẾU THEO CHUẨN; ${extras.length} DÒNG GAB CẦN XÁC MINH` : 'ĐỦ - KHỚP'),
    presentTitles: list(present, (x) => x.g?.title || x.s?.gabTitle || x.s?.recordTitle || ''),
    missingTitles: list(missing, (x) => x.s?.gabTitle || x.s?.recordTitle || ''),
    extraTitles: list(extras, (x) => x.g?.title || ''),
  };
}).sort((a, b) => norm(a.name).localeCompare(norm(b.name)));
const summaryByPerson = new Map(personSummaries.map((p) => [personKey(p.name, p.gab), p]));
const standardCompareKey = (s) => `${gabKey(s.gab) || norm(s.name)}|${norm(s.gabTitle || s.recordTitle)}|${s.gabTime || s.officialDate}`;
const standardKeyCount = new Map();
for (const s of standard) standardKeyCount.set(standardCompareKey(s), (standardKeyCount.get(standardCompareKey(s)) ?? 0) + 1);
const duplicateStandardSet = new Set([...standardKeyCount].filter(([, count]) => count > 1).map(([key]) => key));
const recordIdPeople = new Map();
for (const item of output) {
  if (!item.g?.recordId) continue;
  if (!recordIdPeople.has(item.g.recordId)) recordIdPeople.set(item.g.recordId, new Set());
  recordIdPeople.get(item.g.recordId).add(norm(item.s?.name || item.g.name));
}
const sharedRecordIds = new Set([...recordIdPeople].filter(([, names]) => names.size > 1).map(([recordId]) => recordId));

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
table.push(makeRow([`PHẠM VI CỐ ĐỊNH THEO TAB CHUẨN: ${unique200} tên KLG | ${standard.length} dòng chuẩn, mỗi dòng xuất đúng một dòng đối chiếu | Dữ liệu GAB chỉ dùng để ghép và kết luận, không sinh thêm dòng.`], 2, xfStart + 1, 32));
table.push(makeRow([`GAB thiếu: ${counts['GAB THIẾU THÀNH TỰU']} | Chưa có hồ sơ GAB: ${counts['CHƯA CÓ HỒ SƠ TRÊN GAB']} | Sai/chưa đồng nhất: ${counts['SAI/CHƯA ĐỒNG NHẤT']} | Chỉ có trên GAB - cần xác minh: ${counts['GAB CÓ THÊM - 200 CHƯA CÓ']} | Đủ-khớp: ${counts['ĐỦ - KHỚP']}`], 3, xfStart + 1, 28));
table.push(makeRow(['ĐÃ KIỂM TOÁN: đủ 366/366 dòng chuẩn, không mất/lặp dòng, không có recordId giả và không sai lệch trường GAB khi kéo dữ liệu. Các dòng khác link hồ sơ, khác tên/pháp danh, trùng chuẩn hoặc độ khớp thấp đều được chuyển sang trạng thái cần xác minh.'], 4, xfStart + 1, 42));
table.push('<row r="5" ht="8" customHeight="1"/>');
const headers = ['KẾT LUẬN ĐỐI CHIẾU', 'THIẾU / SAI / CŨ', 'HƯỚNG XỬ LÝ', 'ĐỘ KHỚP', 'TỔNG TT CHUẨN CỦA KLG', 'ĐÃ GHÉP', 'THIẾU TRÊN GAB', 'GAB DƯ CẦN XÁC MINH', 'DANH SÁCH TIÊU ĐỀ GAB DƯ', 'DÒNG TAB CHUẨN', 'STD - STT', 'STD - THÔNG TIN HỒ SƠ', 'STD - SỐ XÁC LẬP KL', 'STD - NGÀY XÁC LẬP KL', 'STD - TÊN KỶ LỤC', 'STD - LINK HỒ SƠ GAB', 'STD - HỌ VÀ TÊN', 'STD - NĂM SINH', 'STD - TỈNH THÀNH', 'STD - HÌNH ĐẠI DIỆN', 'STD - LOẠI KỶ LỤC GIA', 'STD - DANH VỊ KHÁC', 'STD - THỜI GIAN XÁC LẬP GAB', 'STD - TIÊU ĐỀ THÀNH TỰU GAB', 'STD - LINK BÀI VIẾT', 'STD - MÔ TẢ / GIÁ TRỊ', 'STD - HÌNH ẢNH BÀI VIẾT', 'STD - LINK YOUTUBE', 'STD - LINK TIKTOK/FACEBOOK', 'STD - THÀNH TỰU XÁC THỰC', 'GAB - recordId', 'GAB - gabId', 'GAB - fullName', 'GAB - time', 'GAB - title', 'GAB - url', 'GAB - description', 'DÒNG TAB GAB'];
table.push(makeRow(headers, 6, xfStart, 58));
const statusStyle = {
  'GAB THIẾU THÀNH TỰU': xfStart + 2,
  'CHƯA CÓ HỒ SƠ TRÊN GAB': xfStart + 2,
  'SAI/CHƯA ĐỒNG NHẤT': xfStart + 3,
  'GAB CÓ THÊM - 200 CHƯA CÓ': xfStart + 4,
  'ĐỦ - KHỚP': xfStart + 5,
};
const detailOutput = output.filter((r) => r.s?.sourceRow).sort((a, b) => a.s.sourceRow - b.s.sourceRow);
detailOutput.forEach((r, i) => {
  const s = r.s || {}, g = r.g || {};
  const summary = summaryByPerson.get(personKey(s.name || g.name, s.gab || g.gab));
  const gabOnly = r.status === statusNames[3];
  const duplicatedStandard = s.sourceRow && duplicateStandardSet.has(standardCompareKey(s));
  const sharedRecord = g.recordId && sharedRecordIds.has(g.recordId);
  const displayStatus = duplicatedStandard ? 'TRÙNG TRONG FILE CHUẨN - CẦN XÁC MINH' : (sharedRecord ? 'recordId DÙNG CHUNG NHIỀU TÊN - CẦN XÁC MINH' : (gabOnly ? 'CHỈ CÓ TRÊN GAB - CẦN XÁC MINH' : r.status));
  const displayIssues = duplicatedStandard ? `Nội dung chuẩn đang lặp ở nhiều dòng; kiểm tra trước khi kết luận GAB thiếu. ${r.issues}` : (sharedRecord ? `Cùng recordId xuất hiện cho nhiều tên KLG: ${[...recordIdPeople.get(g.recordId)].join(', ')}. Có thể là biệt danh hoặc thành tựu chung.` : r.issues);
  const action = duplicatedStandard
    ? 'Kiểm tra hai dòng nguồn trong tab chuẩn; gộp hoặc giữ riêng nếu có bằng chứng là hai lần xác lập khác nhau.'
    : (sharedRecord
      ? 'Xác minh đây là một người dùng nhiều tên hay thành tựu chung của nhiều người; không tự động xóa hoặc nhân bản.'
      : (gabOnly
      ? 'Kiểm tra đúng KLG, đúng loại thành tựu và nguồn xác thực; chỉ bổ sung file chuẩn sau khi xác minh.'
      : (r.g ? 'Kiểm tra các trường đang lệch; dòng đủ-khớp không cần xử lý.' : 'Kiểm tra lại link, thời gian và nội dung rồi bổ sung thành tựu chuẩn vào GAB.')));
  const noRecord = !g.recordId;
  const gabProfile = g.gab || r.profileGab || (validUrl(s.gab) ? s.gab : 'KHÔNG TÌM THẤY HỒ SƠ GAB');
  const vals = [displayStatus, displayIssues, action, r.score ? `${Math.round(r.score * 100)}%` : 'KHÔNG GHÉP', summary?.standardCount ?? '', summary?.presentCount ?? '', summary?.missingCount ?? '', summary?.extraCount ?? '', summary?.extraTitles || 'KHÔNG CÓ', s.sourceRow || '', s.stt || '', s.profile || '', s.cert || '', s.officialDate || '', s.recordTitle || '', s.gab || 'CHƯA CÓ LINK HỒ SƠ TRONG TAB CHUẨN', s.name || '', s.birth || 'CHƯA CÓ', s.province || 'CHƯA CÓ', s.avatar || 'CHƯA CÓ', s.recordType || 'CHƯA PHÂN LOẠI', s.otherTitle || 'KHÔNG CÓ', s.gabTime || 'CHƯA CÓ', s.gabTitle || 'CHƯA CÓ', s.article || 'CHƯA CÓ', s.description || 'CHƯA CÓ', s.image || 'CHƯA CÓ', s.youtube || 'CHƯA CÓ', s.social || 'CHƯA CÓ', s.verified || 'CHƯA CÓ', g.recordId || 'KHÔNG CÓ RECORD TƯƠNG ỨNG', gabProfile, g.name || (noRecord ? 'KHÔNG CÓ THÀNH TỰU TƯƠNG ỨNG' : 'GAB THIẾU TÊN'), g.time || (noRecord ? 'KHÔNG CÓ THÀNH TỰU TƯƠNG ỨNG' : 'GAB THIẾU THỜI GIAN'), g.title || 'KHÔNG CÓ THÀNH TỰU TƯƠNG ỨNG', g.url || (noRecord ? 'KHÔNG CÓ THÀNH TỰU TƯƠNG ỨNG' : 'GAB THIẾU LINK BÀI'), g.description || (noRecord ? 'KHÔNG CÓ THÀNH TỰU TƯƠNG ỨNG' : 'GAB THIẾU MÔ TẢ'), g.sourceRow || 'KHÔNG CÓ'];
  const base = statusStyle[r.status] ?? xfStart + 7;
  table.push(makeRow(vals, i + 7, vals.map((_, c) => c <= 2 ? base : (c >= 10 && c <= 29 ? xfStart + 6 : (c >= 30 ? xfStart + 4 : xfStart + 7))), 64));
});
const lastDetail = detailOutput.length + 6;
const last = lastDetail;
const widths = [34, 48, 50, 11, 13, 11, 14, 17, 65, 14, 10, 20, 18, 16, 58, 47, 28, 12, 20, 35, 23, 24, 18, 58, 48, 68, 45, 45, 45, 56, 23, 47, 28, 15, 58, 48, 68, 13];
const endCol = colName(headers.length);
const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:${endCol}${last}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane xSplit="4" ySplit="6" topLeftCell="E7" activePane="bottomRight" state="frozen"/><selection pane="bottomRight" activeCell="E7" sqref="E7"/></sheetView></sheetViews><cols>${widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols><sheetData>${table.join('')}</sheetData><autoFilter ref="A6:${endCol}${lastDetail}"/><mergeCells count="4"><mergeCell ref="A1:${endCol}1"/><mergeCell ref="A2:${endCol}2"/><mergeCell ref="A3:${endCol}3"/><mergeCell ref="A4:${endCol}4"/></mergeCells></worksheet>`;
write(path.join(dir, 'xl/worksheets/sheet3.xml'), sheet);

let wb = read(path.join(dir, 'xl/workbook.xml'));
wb = wb.replace(/<sheet[^>]*name="ĐỐI CHIẾU GAB - 200"[^>]*\/>/g, '').replace('</sheets>', '<sheet sheetId="3" name="ĐỐI CHIẾU GAB - 200" state="visible" r:id="rId3"/></sheets>');
write(path.join(dir, 'xl/workbook.xml'), wb);
let rels = read(path.join(dir, 'xl/_rels/workbook.xml.rels'));
rels = rels.replace(/<Relationship[^>]*Id="rId3"[^>]*\/>/g, '').replace('</Relationships>', '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/></Relationships>');
if (!rels.includes('/sharedStrings')) rels = rels.replace('</Relationships>', '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>');
write(path.join(dir, 'xl/_rels/workbook.xml.rels'), rels);
let types = read(path.join(dir, '[Content_Types].xml'));
if (!types.includes('/xl/worksheets/sheet3.xml')) types = types.replace('</Types>', '<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
write(path.join(dir, '[Content_Types].xml'), types);

const zip = path.join(root, 'outputs', 'compare_19_09_2026_result.zip');
if (fs.existsSync(zip)) fs.rmSync(zip, { force: true });
execFileSync('powershell', ['-NoProfile', '-Command', `Add-Type -AssemblyName System.IO.Compression.FileSystem; [IO.Compression.ZipFile]::CreateFromDirectory('${dir.replace(/'/g, "''")}','${zip.replace(/'/g, "''")}')`]);
fs.copyFileSync(zip, source);
const exactProfilePeople = [...standardPeople.values()].filter((p) => gabKey(p.gab) && (gabByPerson.get(gabKey(p.gab))?.length ?? 0) > 0).length;
const nameOnlyPeople = [...standardPeople.values()].filter((p) => !(gabKey(p.gab) && (gabByPerson.get(gabKey(p.gab))?.length ?? 0) > 0) && (gabByPerson.get(norm(p.name))?.length ?? 0) > 0).length;
const lowConfidenceMatches = output.filter((x) => x.g && x.s?.sourceRow && x.score >= 0.62 && x.score < 0.8).length;
const duplicateStandardKeys = standard.length - new Set(standard.map(standardCompareKey)).size;
const standardKeyRows = new Map();
for (const s of standard) {
  const key = standardCompareKey(s);
  if (!standardKeyRows.has(key)) standardKeyRows.set(key, []);
  standardKeyRows.get(key).push(s);
}
const duplicateStandardGroups = [...standardKeyRows.values()].filter((rows) => rows.length > 1).map((rows) => ({ name: rows[0].name, title: rows[0].gabTitle || rows[0].recordTitle, sourceRows: rows.map((r) => r.sourceRow) }));
console.log(JSON.stringify({ source, gabRaw: gabRaw.length, gabDeduplicated: gab.length, duplicateGabRows: gabRaw.length - gab.length, standardRows: standard.length, unique200, matchedPeople, exactProfilePeople, nameOnlyPeople, lowConfidenceMatches, duplicateStandardKeys, duplicateStandardGroups, comparisonRows: detailOutput.length, gabOnlyRowsSummarizedNotExported: counts['GAB CÓ THÊM - 200 CHƯA CÓ'], counts }, null, 2));
