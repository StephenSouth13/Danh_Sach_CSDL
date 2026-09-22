import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import XLSX from 'xlsx';

const root = process.cwd();
const file = path.join(root, 'Dulieutruyxuat', '19_09_2026.xlsx');
const temp = path.join(root, 'outputs', 'customer_sheet_update');
const inputZip = path.join(root, 'outputs', 'customer_sheet_input.zip');
const outputZip = path.join(root, 'outputs', 'customer_sheet_update.zip');
fs.rmSync(temp, { recursive: true, force: true });
fs.mkdirSync(temp, { recursive: true });
fs.copyFileSync(file, inputZip);
execFileSync('powershell', ['-NoProfile', '-Command', `Expand-Archive -LiteralPath '${inputZip.replace(/'/g, "''")}' -DestinationPath '${temp.replace(/'/g, "''")}' -Force`]);

const workbook = XLSX.readFile(file, { raw: false });
const compare = XLSX.utils.sheet_to_json(workbook.Sheets['ĐỐI CHIẾU GAB - 200'], { header: 1, defval: '' });
const sourceRows = compare.slice(6).filter(row => String(row[0] || '').trim());
const groups = new Map();
for (const row of sourceRows) {
  const name = String(row[0]).trim();
  if (!groups.has(name)) groups.set(name, []);
  groups.get(name).push(row);
}

const uniq = values => [...new Set(values.map(value => String(value ?? '').trim()).filter(Boolean))];
const normalize = value => String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/gi, 'd').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
const tokenSet = value => new Set(normalize(value).split(/\s+/).filter(token => token.length > 2));
const similarity = (left, right) => {
  const a = tokenSet(left), b = tokenSet(right);
  if (!a.size || !b.size) return 0;
  let common = 0;
  for (const token of a) if (b.has(token)) common++;
  return common / Math.max(a.size, b.size);
};

const output = [
  ['DANH SÁCH THÀNH TỰU CÒN THIẾU TRÊN GAB'],
  ['Mỗi thành tựu thiếu là một hàng riêng. Các trường nghiệp vụ lấy từ đúng dòng tương ứng trong tab Danh sách 200 người đã chuẩn.'],
  ['recordId ghi KHÔNG CÓ vì thành tựu chưa tồn tại trong dữ liệu GAB xuất về; gabId giữ link hồ sơ KLG nếu đã có.'],
  ['HỌ TÊN KLG', 'THÔNG TIN CÁ NHÂN', 'TỈNH THÀNH MỚI', 'LOẠI KỶ LỤC', 'THỜI GIAN XÁC LẬP', 'TIÊU ĐỀ THÀNH TỰU TRÊN GAB', 'LINK BÀI VIẾT TRÊN GAB', 'THÔNG SỐ KỸ THUẬT + THỜI GIAN + Ý NGHĨA/GIÁ TRỊ/ẢNH HƯỞNG TÍCH CỰC', 'recordId', 'gabId', 'KẾT LUẬN NGẮN', 'TỔNG THÀNH TỰU CHUẨN', 'ĐÃ CÓ TRÊN GAB', 'SỐ THÀNH TỰU THIẾU']
];

for (const [name, rows] of groups) {
  const summary = rows[0];
  const total = Number(summary[5] || 0);
  const present = Number(summary[6] || 0);
  const missing = Number(summary[7] || 0);
  if (missing <= 0) continue;
  const missingItems = String(summary[10] || '').trim().split(/\n(?=\s*\d+\.\s*)/).map(item => item.replace(/^\s*\d+\.\s*/, '').trim()).filter(item => item && item !== 'KHÔNG CÓ');
  const gabIds = uniq(rows.map(row => row[36])).filter(value => !/^KHÔNG CÓ/i.test(value));
  const conclusion = `Chuẩn ${total} | GAB đã có ${present} | Thiếu ${missing}`;
  for (const achievement of missingItems) {
    const best = rows.map(row => ({ row, score: Math.max(similarity(achievement, row[28]), similarity(achievement, row[19])) })).sort((a, b) => b.score - a.score)[0]?.row || summary;
    const personal = String(best[20] || 'CHƯA CÓ');
    const gabId = gabIds[0] || (/^https?:\/\/gab\.world/i.test(personal) ? personal : 'KHÔNG CÓ');
    output.push([String(best[21] || name), personal, String(best[23] || 'CHƯA CÓ'), String(best[25] || 'CHƯA CÓ'), String(best[27] || 'CHƯA CÓ'), achievement, String(best[29] || 'CHƯA CÓ'), String(best[30] || 'CHƯA CÓ'), 'KHÔNG CÓ', gabId, conclusion, total, present, missing]);
  }
}

const escapeXml = value => String(value ?? '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const columnName = number => { let result = ''; while (number) { const remainder = (number - 1) % 26; result = String.fromCharCode(65 + remainder) + result; number = Math.floor((number - 1) / 26); } return result; };
const sheetRows = output.map((row, index) => {
  const rowNumber = index + 1;
  const cells = row.map((value, columnIndex) => {
    if (value === '' || value == null) return '';
    const style = rowNumber === 4 ? 1 : rowNumber < 4 ? 2 : columnIndex === 8 ? 6 : columnIndex === 9 ? 5 : 3;
    return `<c r="${columnName(columnIndex + 1)}${rowNumber}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
  }).join('');
  return `<row r="${rowNumber}" ht="${rowNumber === 4 ? 60 : rowNumber < 4 ? 28 : 92}" customHeight="1">${cells}</row>`;
}).join('');
const widths = [28, 48, 22, 25, 20, 72, 62, 88, 24, 48, 42, 16, 16, 16].map((width, index) => `<col min="${index + 1}" max="${index + 1}" width="${width}" customWidth="1"/>`).join('');
const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:N${output.length}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane xSplit="2" ySplit="4" topLeftCell="C5" activePane="bottomRight" state="frozen"/></sheetView></sheetViews><cols>${widths}</cols><sheetData>${sheetRows}</sheetData><autoFilter ref="A4:N${output.length}"/><mergeCells count="3"><mergeCell ref="A1:N1"/><mergeCell ref="A2:N2"/><mergeCell ref="A3:N3"/></mergeCells></worksheet>`;
fs.writeFileSync(path.join(temp, 'xl', 'worksheets', 'sheet4.xml'), sheetXml, 'utf8');
fs.rmSync(outputZip, { force: true });
execFileSync('powershell', ['-NoProfile', '-Command', `Add-Type -AssemblyName System.IO.Compression.FileSystem; [IO.Compression.ZipFile]::CreateFromDirectory('${temp.replace(/'/g, "''")}','${outputZip.replace(/'/g, "''")}')`]);
fs.copyFileSync(outputZip, file);
console.log(JSON.stringify({ file, customerRows: output.length - 4, sourceRows: sourceRows.length, sourcePeople: groups.size }, null, 2));
