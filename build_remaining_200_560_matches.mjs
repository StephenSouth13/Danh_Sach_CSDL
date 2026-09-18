import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';

const root = process.cwd();
const expanded = path.join(root, 'outputs', 'complete_standard_work', 'xlsx');
const source = path.join(root, 'Dulieutruyxuat', 'So sánh.xlsx');
const read = (p) => fs.readFileSync(p, 'utf8');
const write = (p, v) => fs.writeFileSync(p, v, 'utf8');
const esc = (v = '') => String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const dec = (v = '') => String(v).replace(/&#x([0-9a-f]+);/gi, (_, x) => String.fromCodePoint(parseInt(x, 16))).replace(/&#(\d+);/g, (_, x) => String.fromCodePoint(Number(x))).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const colNo = (ref) => { let n = 0; for (const c of String(ref).match(/^[A-Z]+/)?.[0] ?? '') n = n * 26 + c.charCodeAt(0) - 64; return n; };
const colName = (n) => { let s = ''; while (n) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; };
const norm = (v = '') => String(v).toLocaleLowerCase('vi').normalize('NFD').replace(/\p{Diacritic}/gu, '').replace(/đ/g, 'd').replace(/\b(ky luc gia|klg|ong|ba|anh|chi|gs|pgs|ts|ths|bs|nghe nhan|nghe si|hoa si|tien si|thac si|viet nam|chau a|the gioi)\b/g, ' ').replace(/[^a-z0-9]+/g, ' ').replace(/\s+/g, ' ').trim();
const gabKey = (v = '') => String(v).match(/gab\.world\/vi\/bank\/([^/?#]+)/i)?.[1]?.toLowerCase() ?? '';

const sharedPath = path.join(expanded, 'xl', 'sharedStrings.xml');
const shared = fs.existsSync(sharedPath) ? [...read(sharedPath).matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map((m) => [...m[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((x) => dec(x[1])).join('')) : [];
function rows(sheetFile) {
  const result = [];
  const xml = read(path.join(expanded, 'xl', 'worksheets', sheetFile));
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

const list560 = rows('sheet1.xml').filter((r) => r.row > 1).map((r) => ({ row: r.row, name: String(r.values.get(4) ?? '').trim(), link: String(r.values.get(3) ?? '').trim() })).filter((r) => r.name);
const old43Rows = rows('sheet5.xml').filter((r) => r.row >= 6);
const oldNames = new Set();
const oldGabs = new Set();
for (const r of old43Rows) {
  for (const col of [2, 3, 4]) { const n = norm(r.values.get(col) ?? ''); if (n) oldNames.add(n); }
  for (const col of [6]) { const g = gabKey(r.values.get(col) ?? ''); if (g) oldGabs.add(g); }
}

const standardRows = rows('sheet2.xml').filter((r) => r.row >= 3 && String(r.values.get(7) ?? '').trim());
const people = new Map();
for (const r of standardRows) {
  const name = String(r.values.get(7) ?? '').trim();
  const link = String(r.values.get(6) ?? '').trim();
  const key = gabKey(link) || norm(name);
  if (!people.has(key)) people.set(key, { name, link, rows: [] });
  people.get(key).rows.push(r);
}

const matches = [];
for (const person of people.values()) {
  const personGab = gabKey(person.link), personName = norm(person.name);
  if ((personGab && oldGabs.has(personGab)) || oldNames.has(personName)) continue;
  let hit = personGab ? list560.find((x) => gabKey(x.link) === personGab) : null;
  let basis = hit ? 'Trùng chính xác link GAB' : '';
  if (!hit) {
    hit = list560.find((x) => norm(x.name) === personName);
    if (hit) basis = 'Trùng chính xác tên đã chuẩn hóa';
  }
  if (!hit) continue;
  for (const row of person.rows) matches.push({ person, sourceRow: row, hit, basis });
}
matches.sort((a, b) => norm(a.person.name).localeCompare(norm(b.person.name)) || a.sourceRow.row - b.sourceRow.row);
const uniquePeople = new Set(matches.map((x) => gabKey(x.person.link) || norm(x.person.name))).size;

let styles = read(path.join(expanded, 'xl', 'styles.xml'));
function append(xml, tag, item) {
  const re = new RegExp(`<${tag}\\b([^>]*)count="(\\d+)"([^>]*)>([\\s\\S]*?)<\\/${tag}>`);
  const m = xml.match(re); if (!m) throw new Error(`Không tìm thấy ${tag}`);
  return xml.replace(re, `<${tag}${m[1]}count="${Number(m[2]) + 1}"${m[3]}>${m[4]}${item}</${tag}>`);
}
const fontId = Number(styles.match(/<fonts\b[^>]*count="(\d+)"/)?.[1] ?? 0);
styles = append(styles, 'fonts', '<font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Arial"/></font>');
const fillId = Number(styles.match(/<fills\b[^>]*count="(\d+)"/)?.[1] ?? 0);
for (const color of ['FF17365D', 'FFD9EAF7', 'FFE2F0D9']) styles = append(styles, 'fills', `<fill><patternFill patternType="solid"><fgColor rgb="${color}"/><bgColor indexed="64"/></patternFill></fill>`);
const xfId = Number(styles.match(/<cellXfs\b[^>]*count="(\d+)"/)?.[1] ?? 0);
for (const xf of [
  `<xf numFmtId="0" fontId="${fontId}" fillId="${fillId}" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>`,
  `<xf numFmtId="0" fontId="0" fillId="${fillId + 1}" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>`,
  `<xf numFmtId="0" fontId="0" fillId="${fillId + 2}" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>`,
  '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>',
]) styles = append(styles, 'cellXfs', xf);
write(path.join(expanded, 'xl', 'styles.xml'), styles);

const cell = (v, ref, style) => `<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
const makeRow = (vals, n, stylesForRow, height = 46) => `<row r="${n}" ht="${height}" customHeight="1">${vals.map((v, i) => cell(v, `${colName(i + 1)}${n}`, Array.isArray(stylesForRow) ? stylesForRow[i] : stylesForRow)).join('')}</row>`;
const table = [];
table.push(makeRow(['KLG TRÙNG MỚI GIỮA 200 NGƯỜI ĐÃ CHUẨN VÀ DANH SÁCH 560'], 1, xfId, 30));
table.push(makeRow([`Đã loại nhóm cũ trong “THÀNH TỰU TRÙNG 43 KLG”. Còn ${uniquePeople} KLG trùng mới, tương ứng ${matches.length} dòng dữ liệu/thành tựu.`], 2, xfId + 1, 26));
table.push(makeRow(['Mỗi dòng giữ nguyên đầy đủ dữ liệu A:T từ tab 200 người đã chuẩn; ba cột cuối cho biết kết quả đối chiếu với danh sách 560.'], 3, xfId + 1, 26));
table.push('<row r="4" ht="8" customHeight="1"/>');
const originalHeaders = Array.from({ length: 20 }, (_, i) => String(rows('sheet2.xml').find((r) => r.row === 1)?.values.get(i + 1) || rows('sheet2.xml').find((r) => r.row === 2)?.values.get(i + 1) || colName(i + 1)));
const headers = [...originalHeaders, 'TÊN TRONG DANH SÁCH 560', 'LINK GAB TRONG DANH SÁCH 560', 'CĂN CỨ ĐỐI CHIẾU'];
table.push(makeRow(headers, 5, xfId, 54));
matches.forEach((m, index) => {
  const vals = Array.from({ length: 20 }, (_, i) => String(m.sourceRow.values.get(i + 1) ?? ''));
  vals.push(m.hit.name, m.hit.link, m.basis);
  table.push(makeRow(vals, index + 6, vals.map((_, i) => i >= 20 ? xfId + 2 : xfId + 3), 58));
});
const last = matches.length + 5;
const widths = [7, 22, 17, 16, 52, 45, 28, 12, 18, 18, 27, 20, 17, 64, 50, 68, 20, 30, 46, 68, 30, 48, 27];
const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><dimension ref="A1:W${last}"/><cols>${widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')}</cols><sheetData>${table.join('')}</sheetData><autoFilter ref="A5:W${last}"/><mergeCells count="3"><mergeCell ref="A1:W1"/><mergeCell ref="A2:W2"/><mergeCell ref="A3:W3"/></mergeCells></worksheet>`;
write(path.join(expanded, 'xl', 'worksheets', 'sheet7.xml'), sheet);

let wb = read(path.join(expanded, 'xl', 'workbook.xml'));
wb = wb.replace(/<sheet[^>]*name="TRÙNG MỚI 200 - 560"[^>]*\/>/g, '').replace('</sheets>', '<sheet sheetId="7" name="TRÙNG MỚI 200 - 560" state="visible" r:id="rId10"/></sheets>');
write(path.join(expanded, 'xl', 'workbook.xml'), wb);
let rels = read(path.join(expanded, 'xl', '_rels', 'workbook.xml.rels'));
rels = rels.replace(/<Relationship[^>]*Id="rId10"[^>]*\/>/g, '').replace('</Relationships>', '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet7.xml"/></Relationships>');
write(path.join(expanded, 'xl', '_rels', 'workbook.xml.rels'), rels);
let types = read(path.join(expanded, '[Content_Types].xml'));
if (!types.includes('/xl/worksheets/sheet7.xml')) types = types.replace('</Types>', '<Override PartName="/xl/worksheets/sheet7.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
write(path.join(expanded, '[Content_Types].xml'), types);

const zip = path.join(root, 'outputs', 'complete_standard_work', 'remaining_matches.zip');
if (fs.existsSync(zip)) fs.rmSync(zip, { force: true });
execFileSync('powershell', ['-NoProfile', '-Command', `Add-Type -AssemblyName System.IO.Compression.FileSystem; [IO.Compression.ZipFile]::CreateFromDirectory('${expanded.replace(/'/g, "''")}','${zip.replace(/'/g, "''")}')`]);
fs.copyFileSync(zip, source);
console.log(JSON.stringify({ uniquePeople, dataRows: matches.length, names: [...new Set(matches.map((x) => x.person.name))] }, null, 2));
