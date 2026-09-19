import fs from 'node:fs';

const dir='outputs/compare_19_09_2026/xl';
const dec=(s='')=>s.replace(/&#x([0-9a-f]+);/gi,(_,x)=>String.fromCodePoint(parseInt(x,16))).replace(/&#(\d+);/g,(_,x)=>String.fromCodePoint(+x)).replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&amp;/g,'&');
const shared=[...fs.readFileSync(`${dir}/sharedStrings.xml`,'utf8').matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map(m=>[...m[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(x=>dec(x[1])).join(''));
const colNo=(s)=>{let n=0;for(const c of s)n=n*26+c.charCodeAt(0)-64;return n};
const norm=(v='')=>String(v).toLocaleLowerCase('vi').normalize('NFD').replace(/\p{Diacritic}/gu,'').replace(/đ/g,'d').replace(/[^a-z0-9]+/g,' ').replace(/\s+/g,' ').trim();
const rows=(file)=>{
  const xml=fs.readFileSync(`${dir}/worksheets/${file}`,'utf8'), out=[];
  for(const m of xml.matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>/g)){
    const n=+m[1].match(/\br="(\d+)"/)?.[1], v={};
    for(const c of m[2].matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)){
      const ref=c[1].match(/\br="([A-Z]+)\d+"/)?.[1]; if(!ref)continue;
      const type=c[1].match(/\bt="([^"]+)"/)?.[1], body=c[2]||'';
      let value='';
      if(type==='s') value=shared[+(body.match(/<v>([\s\S]*?)<\/v>/)?.[1]??-1)]??'';
      else if(type==='inlineStr') value=[...body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(x=>dec(x[1])).join('');
      else value=dec(body.match(/<v>([\s\S]*?)<\/v>/)?.[1]??'');
      v[colNo(ref)]=value;
    }
    out.push({row:n,v});
  }
  return out;
};
const gab=rows('sheet1.xml').filter(r=>r.row>1&&r.v[1]&&r.v[3]&&r.v[5]);
const std=rows('sheet2.xml').filter(r=>r.row>2&&r.v[7]&&(r.v[5]||r.v[14]||r.v[20]));
const out=rows('sheet3.xml').filter(r=>r.row>=7);
const gabById=new Map(); for(const r of gab)if(!gabById.has(r.v[1]))gabById.set(r.v[1],r);
const stdRows=new Set(std.map(r=>String(r.row))), outStdRows=out.map(r=>String(r.v[10]));
const duplicates=(arr)=>[...new Set(arr.filter((x,i)=>x&&arr.indexOf(x)!==i))];
const missingStd=[...stdRows].filter(x=>!outStdRows.includes(x));
const extraStd=[...new Set(outStdRows)].filter(x=>!stdRows.has(x));
const duplicateStdRows=duplicates(outStdRows);
const missingGabIds=[], fieldMismatches=[], profileMismatches=[], nameMismatches=[];
for(const r of out){
  const id=r.v[31]; if(!id)continue;
  const g=gabById.get(id); if(!g){missingGabIds.push({outRow:r.row,id});continue;}
  const checks=[[32,2,'gabId'],[33,3,'fullName'],[35,5,'title'],[36,6,'url'],[37,7,'description']];
  for(const [oc,gc,label] of checks)if(String(r.v[oc]??'').trim()!==String(g.v[gc]??'').trim())fieldMismatches.push({outRow:r.row,id,label,output:r.v[oc]??'',source:g.v[gc]??''});
  if(r.v[16]&&r.v[32]&&r.v[16]!==r.v[32])profileMismatches.push({outRow:r.row,stdName:r.v[17],stdGab:r.v[16],gabName:r.v[33],gabId:r.v[32],id});
  if(r.v[17]&&r.v[33]&&norm(r.v[17])!==norm(r.v[33]))nameMismatches.push({outRow:r.row,stdName:r.v[17],gabName:r.v[33],id});
}
const samePersonRecord=new Map();
for(const r of out){if(!r.v[31])continue;const k=`${norm(r.v[17])}|${r.v[31]}`;if(!samePersonRecord.has(k))samePersonRecord.set(k,[]);samePersonRecord.get(k).push(r.row)}
const duplicateAssignments=[...samePersonRecord].filter(([,rs])=>rs.length>1).map(([key,rows])=>({key,rows}));
const blankNames=out.filter(r=>!r.v[17]).map(r=>r.row);
console.log(JSON.stringify({standardSourceRows:std.length,outputRows:out.length,uniqueStandardNames:new Set(std.map(r=>norm(r.v[7]))).size,missingStd,extraStd,duplicateStdRows,blankNames,gabSourceRows:gab.length,uniqueGabRecordIds:gabById.size,matchedOutputRows:out.filter(r=>r.v[31]).length,missingGabIds,fieldMismatchCount:fieldMismatches.length,fieldMismatchSamples:fieldMismatches.slice(0,10),profileMismatchCount:profileMismatches.length,profileMismatchSamples:profileMismatches.slice(0,10),nameMismatchCount:nameMismatches.length,nameMismatchSamples:nameMismatches.slice(0,20),duplicateAssignmentCount:duplicateAssignments.length,duplicateAssignments},null,2));
