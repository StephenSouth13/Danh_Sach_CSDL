import fs from 'node:fs';
import path from 'node:path';

const root = path.resolve('outputs');
const targets = [
  ['ngo_thi_hoang_ngan', 'Ngô Thị Hoàng Ngân'],
  ['nguyen_ngoc_phuong_lam', 'Nguyễn Ngọc Phương Lam'],
  ['nguyen_thi_kim_oanh', 'Nguyễn Thị Kim Oanh'],
  ['nguyen_quang_thang', 'Nguyễn Quang Thắng'],
  ['dang_hong_mi', 'Đặng Hồng Mi'],
  ['ngo_thu_an', 'Ngô Thu An'],
  ['tran_quoc_huy', 'Trần Quốc Huy'],
  ['nguyen_hoang_bach', 'Nguyễn Hoàng Bách'],
  ['tran_hoai_thuan', 'Trần Hoài Thuận'],
  ['luong_thanh_nhat', 'Lương Thành Nhật'],
  ['huynh_hoang_son', 'Huỳnh Hoàng Sơn'],
];

const decode = (s) => s
  .replace(/<[^>]*>/g, ' ')
  .replace(/&nbsp;/g, ' ')
  .replace(/&amp;/g, '&')
  .replace(/&#(d+);/g, (_, n) => String.fromCodePoint(Number(n)))
  .replace(/\s+/g, ' ')
  .trim();
const norm = (s) => decode(s).normalize('NFC').toLocaleLowerCase('vi');

const output = {};
const broadOutput = {};
for (const [slug, name] of targets) {
  const candidates = [];
  const broad = [];
  for (let page = 1; page <= 100; page++) {
    const file = path.join(root, `search_${slug}_${page}.html`);
    if (!fs.existsSync(file)) continue;
    const html = fs.readFileSync(file, 'utf8');
    const blocks = html.match(/<div class="search-result-item">[\s\S]*?<\/div>\s*<\/div>/g) ?? [];
    for (const block of blocks) {
      const titleMatch = block.match(/class="search-result-title" href="([^"]+)">([\s\S]*?)<\/a>/);
      const descMatch = block.match(/class="search-result-desc">([\s\S]*?)<\/div>/);
      if (!titleMatch) continue;
      const title = decode(titleMatch[2]);
      const description = decode(descMatch?.[1] ?? '');
      const haystack = norm(`${title} ${description}`);
      const tokens = norm(name).split(' ');
      const tokenHits = tokens.filter((token) => haystack.includes(token)).length;
      const recordSignal = /kỷ lục|vietkings|xác lập|record/i.test(`${title} ${description}`);
      if (recordSignal) broad.push({ page, tokenHits, title, description, url: `https://kyluc.vn${titleMatch[1]}` });
      if (haystack.includes(norm(name)) || (tokenHits >= Math.max(2, tokens.length - 1) && recordSignal)) {
        candidates.push({ page, tokenHits, recordSignal, title, description, url: `https://kyluc.vn${titleMatch[1]}` });
      }
    }
  }
  output[name] = candidates.sort((a, b) => b.tokenHits - a.tokenHits || Number(b.recordSignal) - Number(a.recordSignal));
  broadOutput[name] = broad.sort((a, b) => b.tokenHits - a.tokenHits);
}

fs.writeFileSync(path.join(root, 'kyluc_candidates.json'), JSON.stringify(output, null, 2));
fs.writeFileSync(path.join(root, 'kyluc_broad_candidates.json'), JSON.stringify(broadOutput, null, 2));
for (const [name, rows] of Object.entries(output)) {
  console.log(`\n## ${name}: ${rows.length}`);
  for (const row of rows) console.log(`- p${row.page} ${row.title}\n  ${row.url}\n  ${row.description}`);
}
