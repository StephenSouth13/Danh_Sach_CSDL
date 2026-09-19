import fs from 'node:fs';
import path from 'node:path';

const dir = path.resolve('outputs/compare_19_09_2026');
const dec = (v = '') => String(v).replace(/&#x([0-9a-f]+);/gi, (_, x) => String.fromCodePoint(parseInt(x, 16))).replace(/&#(\d+);/g, (_, x) => String.fromCodePoint(Number(x))).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
const colNo = (ref) => { let n = 0; for (const c of String(ref).match(/^[A-Z]+/)?.[0] ?? '') n = n * 26 + c.charCodeAt(0) - 64; return n; };
const wb = fs.readFileSync(path.join(dir, 'xl/workbook.xml'), 'utf8');
const rels = fs.readFileSync(path.join(dir, 'xl/_rels/workbook.xml.rels'), 'utf8');
const relMap = new Map([...rels.matchAll(/<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g)].map((m) => [m[1], m[2]]));
const sheets = [...wb.matchAll(/<sheet\b[^>]*name="([^"]+)"[^>]*(?:r:id|id)="([^"]+)"/g)].map((m) => ({ name: dec(m[1]), target: relMap.get(m[2]) }));
const sharedPath = path.join(dir, 'xl/sharedStrings.xml');
const shared = fs.existsSync(sharedPath) ? [...fs.readFileSync(sharedPath, 'utf8').matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map((m) => [...m[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((x) => dec(x[1])).join('')) : [];
function parseSheet(target) {
  const xml = fs.readFileSync(path.join(dir, 'xl', target.replace(/^xl\//, '')), 'utf8');
  const rows = [];
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
    rows.push({ row, values });
  }
  return rows;
}

for (const sheet of sheets) {
  const rows = parseSheet(sheet.target);
  const usedCols = Math.max(0, ...rows.flatMap((r) => [...r.values.keys()]));
  console.log(`\n## ${sheet.name} | rows=${rows.length} | maxCol=${usedCols} | target=${sheet.target}`);
  for (const r of rows.slice(0, 8)) console.log(JSON.stringify({ row: r.row, values: Object.fromEntries(r.values) }));
  const nonEmptyByCol = {};
  for (let c = 1; c <= usedCols; c++) nonEmptyByCol[c] = rows.filter((r) => String(r.values.get(c) ?? '').trim()).length;
  console.log('nonEmptyByCol', JSON.stringify(nonEmptyByCol));
}
