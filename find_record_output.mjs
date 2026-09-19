import fs from 'node:fs';
const xml=fs.readFileSync('outputs/compare_19_09_2026/xl/worksheets/sheet3.xml','utf8');
const id='8hfS1XzkmtjNooLmWELn';
const match=[...xml.matchAll(/<row r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)].find(m=>m[2].includes(id));
if(!match){console.log('NOT FOUND');process.exit();}
const decode=s=>s.replace(/&quot;/g,'"').replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>');
const cells=[...match[2].matchAll(/<c r="([A-Z]+\d+)"[^>]*>[\s\S]*?<t[^>]*>([\s\S]*?)<\/t>[\s\S]*?<\/c>/g)].map(m=>[m[1],decode(m[2])]);
console.log(JSON.stringify({row:Number(match[1]),cells:Object.fromEntries(cells)},null,2));
