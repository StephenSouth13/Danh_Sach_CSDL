import fs from 'node:fs';
import path from 'node:path';

const file = path.join('outputs', 'complete_standard_work', 'xlsx', 'xl', 'worksheets', 'sheet2.xml');
const row = 346;
const values = {
  D: '10/05/2025',
  E: 'Thành viên sáng lập kiêm Tổng đạo diễn sự kiện xếp hình lá cờ Tổ quốc có số lượng người tham gia cùng lúc đông nhất Việt Nam',
  K: 'KỶ LỤC VIỆT NAM',
  M: '10/05/2025',
  N: 'Tham gia xác lập Kỷ lục Việt Nam “Sự kiện xếp hình lá cờ Tổ quốc có số lượng người tham gia cùng lúc đông nhất Việt Nam” với vai trò thành viên sáng lập kiêm Tổng đạo diễn',
  O: 'https://kyluc.vn/tin-tuc/ky-luc/hai-phong-gan-7-000-nguoi-mac-ao-co-do-sao-vang-tao-nen-la-co-to-quoc-lon-nhat-viet-nam',
  S: 'Đã đối chiếu đúng người theo tên và đơn vị ACCT Hoàng Ngân. Vai trò thành viên sáng lập kiêm Tổng đạo diễn được xác nhận từ bài giới thiệu cá nhân; nội dung và ngày xác lập Kỷ lục được đối chiếu với bài chính thức Kyluc.vn.',
};
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const setCell = (rowXml, col, value) => {
  const ref = `${col}${row}`;
  const re = new RegExp(`<c\\b[^>]*\\br="${ref}"[^>]*>[\\s\\S]*?<\\/c>|<c\\b[^>]*\\br="${ref}"[^>]*/>`);
  const old = rowXml.match(re)?.[0];
  const style = old?.match(/\bs="(\d+)"/)?.[1] ?? '1';
  const cell = `<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`;
  return old ? rowXml.replace(re, cell) : rowXml.replace('</row>', `${cell}</row>`);
};

let xml = fs.readFileSync(file, 'utf8');
const re = new RegExp(`<row\\b[^>]*\\br="${row}"[^>]*>[\\s\\S]*?<\\/row>`);
const original = xml.match(re)?.[0];
if (!original || !original.includes('NGÔ THỊ HOÀNG NGÂN')) throw new Error('Không tìm thấy đúng hàng NGÔ THỊ HOÀNG NGÂN');
let changed = original;
for (const [col, value] of Object.entries(values)) changed = setCell(changed, col, value);
xml = xml.replace(re, changed);
fs.writeFileSync(file, xml, 'utf8');
console.log('Updated row 346 - NGÔ THỊ HOÀNG NGÂN');
