import fs from 'node:fs';

const dir = 'outputs/complete_standard_work/xlsx';
const sheetFile = process.argv[2] || 'sheet2.xml';
const xml = fs.readFileSync(`${dir}/xl/worksheets/${sheetFile}`, 'utf8');
const sharedPath = `${dir}/xl/sharedStrings.xml`;
const decode = (s = '') => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const shared = fs.existsSync(sharedPath)
  ? [...fs.readFileSync(sharedPath, 'utf8').matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map((m) => [...m[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((t) => decode(t[1])).join(''))
  : [];
const value = (attrs, body = '') => {
  const type = attrs.match(/\bt="([^"]+)"/)?.[1];
  if (type === 's') return shared[Number(body.match(/<v>([\s\S]*?)<\/v>/)?.[1])] ?? '';
  if (type === 'inlineStr') return [...body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((m) => decode(m[1])).join('');
  return decode(body.match(/<v>([\s\S]*?)<\/v>/)?.[1] ?? '');
};
for (const match of xml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)) {
  const row = Number(match[1]);
  if (sheetFile === 'sheet2.xml' && row > 2 && (row < 340 || row > 365)) continue;
  const out = {};
  for (const cell of match[2].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>|<c\b([^>]*)\/>/g)) {
    const attrs = cell[1] || cell[3] || '';
    const ref = attrs.match(/\br="([A-Z]+)\d+"/)?.[1];
    if (ref && ref <= 'T') out[ref] = value(attrs, cell[2] ?? '');
  }
  const line = JSON.stringify({ row, ...out });
  if (sheetFile === 'sheet2.xml' || /NGÔ THỊ HOÀNG NGÂN|NGUYỄN NGỌC PHƯƠNG LAM|NGUYỄN THỊ KIM OANH|NGUYỄN QUANG THẮNG|ĐẶNG HỒNG MI|NGÔ THU AN|TRẦN QUỐC HUY|NGUYỄN HOÀNG BÁCH|TRẦN HOÀI THUẬN|LƯƠNG THÀNH NHẬT|HUỲNH HOÀNG SƠN/i.test(line)) console.log(line);
}
