import fs from 'node:fs';

const dir = 'outputs/complete_standard_work/xlsx';
const xml = fs.readFileSync(`${dir}/xl/worksheets/sheet2.xml`, 'utf8');
const sharedXml = fs.readFileSync(`${dir}/xl/sharedStrings.xml`, 'utf8');
const decode = (s = '') => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const shared = [...sharedXml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map((m) => [...m[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((t) => decode(t[1])).join(''));
const value = (attrs, body = '') => {
  const type = attrs.match(/\bt="([^"]+)"/)?.[1];
  if (type === 's') return shared[Number(body.match(/<v>([\s\S]*?)<\/v>/)?.[1])] ?? '';
  if (type === 'inlineStr') return [...body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((m) => decode(m[1])).join('');
  return decode(body.match(/<v>([\s\S]*?)<\/v>/)?.[1] ?? '');
};

const results = [];
for (const rowMatch of xml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)) {
  const row = Number(rowMatch[1]);
  const cells = {};
  for (const cell of rowMatch[2].matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>|<c\b([^>]*)\/>/g)) {
    const attrs = cell[1] || cell[3] || '';
    const col = attrs.match(/\br="([A-Z]+)\d+"/)?.[1];
    if (col) cells[col] = value(attrs, cell[2] ?? '');
  }
  const status = cells.K ?? '';
  const missingAchievement = cells.G && !cells.E;
  const unresolvedStatus = /CHƯA XÁC MINH|CÓ GAB - CHƯA/i.test(status);
  if (missingAchievement || unresolvedStatus) results.push({ row, name: cells.G, gab: cells.F, status, note: cells.S });
}
console.log(JSON.stringify(results, null, 2));
console.log(`TOTAL=${results.length}`);
