import fs from 'node:fs';
const dir = 'outputs/compare_19_09_2026/xl';
const dec = (s='') => s.replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"');
const shared = [...fs.readFileSync(`${dir}/sharedStrings.xml`,'utf8').matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map(m => [...m[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(x=>dec(x[1])).join(''));
const xml = fs.readFileSync(`${dir}/worksheets/sheet2.xml`,'utf8');
for (const n of [1,2]) {
  const body = xml.match(new RegExp(`<row r="${n}"[^>]*>([\\s\\S]*?)<\\/row>`))?.[1] ?? '';
  const vals=[];
  for(const m of body.matchAll(/<c\b([^>]*?)>([\s\S]*?)<\/c>/g)){
    const ref=m[1].match(/r="([A-Z]+)\d+"/)?.[1]; const type=m[1].match(/t="([^"]+)"/)?.[1];
    const raw=m[2].match(/<v>([\s\S]*?)<\/v>/)?.[1]??''; vals.push(`${ref}: ${type==='s'?(shared[+raw]??''):dec(raw)}`);
  }
  console.log(`ROW ${n}`, vals);
}
