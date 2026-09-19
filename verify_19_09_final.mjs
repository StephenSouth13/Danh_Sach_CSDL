import fs from 'node:fs';

const xml = fs.readFileSync('outputs/compare_19_09_2026/xl/worksheets/sheet3.xml', 'utf8');
const decode = (s = '') => s.replace(/&quot;/g, '"').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>');
const row = (n) => xml.match(new RegExp(`<row r="${n}"[^>]*>([\\s\\S]*?)<\\/row>`))?.[1] ?? '';
const cells = (body) => Object.fromEntries([...body.matchAll(/<c r="([A-Z]+)\d+"[^>]*>[\s\S]*?<t[^>]*>([\s\S]*?)<\/t>[\s\S]*?<\/c>/g)].map((m) => [m[1], decode(m[2])]));
const headers = cells(row(6));
const duplicateWarnings = (xml.match(/TRÙNG TRONG FILE CHUẨN - CẦN XÁC MINH/g) ?? []).length;
const gabOnlyWarnings = (xml.match(/CHỈ CÓ TRÊN GAB - CẦN XÁC MINH/g) ?? []).length;
const recordIds = [...xml.matchAll(/<c r="Q\d+"[^>]*>[\s\S]*?<t[^>]*>([\s\S]*?)<\/t>[\s\S]*?<\/c>/g)].map((m) => decode(m[1])).filter(Boolean);
const gabIds = [...xml.matchAll(/<c r="R\d+"[^>]*>[\s\S]*?<t[^>]*>([\s\S]*?)<\/t>[\s\S]*?<\/c>/g)].map((m) => decode(m[1])).filter(Boolean);
const recordRows = new Map();
for (let n = 7; n <= 658; n++) {
  const c = cells(row(n));
  if (!c.Q) continue;
  if (!recordRows.has(c.Q)) recordRows.set(c.Q, []);
  recordRows.get(c.Q).push({ row: n, name: c.B, status: c.C, title: c.L, gabId: c.R });
}
const repeatedRecordIds = [...recordRows].filter(([, rows]) => rows.length > 1).map(([recordId, rows]) => ({ recordId, rows }));
console.log(JSON.stringify({ dimension: xml.match(/<dimension ref="([^"]+)/)?.[1], headers, duplicateWarnings, gabOnlyWarnings, recordIdValues: recordIds.length, uniqueRecordIds: new Set(recordIds).size, repeatedRecordIds, gabIdValues: gabIds.length, uniqueGabIds: new Set(gabIds).size, frozenPane: xml.match(/<pane[^>]+>/)?.[0] }, null, 2));
