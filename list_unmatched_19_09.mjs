import fs from 'node:fs';
const x=fs.readFileSync('outputs/compare_19_09_2026/xl/worksheets/sheet3.xml','utf8');
const dec=s=>s.replace(/&quot;/g,'"').replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>');
const result=[];
for(const m of x.matchAll(/<row r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)){
 const v={};for(const c of m[2].matchAll(/<c r="([A-Z]+)\d+"[^>]*>[\s\S]*?<t[^>]*>([\s\S]*?)<\/t>[\s\S]*?<\/c>/g))v[c[1]]=dec(c[2]);
 if(v.A==='CHƯA CÓ HỒ SƠ TRÊN GAB')result.push({row:+m[1],name:v.Q,stdGab:v.P,title:v.X});
}
const people=[...new Map(result.map(r=>[r.name,{name:r.name,stdGab:r.stdGab,rows:[]}])).values()];for(const p of people)p.rows=result.filter(r=>r.name===p.name).map(r=>r.row);
console.log(JSON.stringify({rows:result.length,people:people.length,list:people},null,2));
