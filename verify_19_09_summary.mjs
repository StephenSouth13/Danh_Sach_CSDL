import fs from 'node:fs';

const xml = fs.readFileSync('outputs/compare_19_09_2026/xl/worksheets/sheet3.xml', 'utf8');
console.log('dimension', xml.match(/<dimension ref="([^"]+)/)?.[1]);
for (const rowNumber of [1, 2, 6, 7, 8]) {
  const body = xml.match(new RegExp(`<row r="${rowNumber}"[^>]*>([\\s\\S]*?)<\\/row>`))?.[1] ?? '';
  const cells = [...body.matchAll(/<c r="([^"]+)"[^>]*>[\s\S]*?<t[^>]*>([\s\S]*?)<\/t>[\s\S]*?<\/c>/g)]
    .filter((m) => /^(U|V|W|X|Y|Z|AA|AB|AC|AD)/.test(m[1]))
    .map((m) => `${m[1]}: ${m[2].slice(0, 120)}`);
  console.log(`row ${rowNumber}`, cells);
}
console.log('KLG summary rows', [...xml.matchAll(/<c r="U(\d+)"/g)].filter((m) => Number(m[1]) >= 7).length);
