import fs from 'node:fs';
import path from 'node:path';
import XLSX from 'xlsx';
const dir=path.join(process.cwd(),'Dulieutruyxuat'), main=path.join(dir,'19_09_2026.xlsx');
const norm=v=>String(v??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/đ/gi,'d').toLowerCase().replace(/[^a-z0-9]+/g,' ').trim();
const toks=v=>new Set(norm(v).split(/\s+/).filter(x=>x.length>2));
const sim=(a,b)=>{const A=toks(a),B=toks(b);if(!A.size||!B.size)return 0;let n=0;for(const x of A)if(B.has(x))n++;return n/Math.max(A.size,B.size)};
const mainWb=XLSX.readFile(main,{raw:false}), target=XLSX.utils.sheet_to_json(mainWb.Sheets.sheet6,{header:1,defval:''}).slice(4).filter(r=>r[0]);
const nameKeys=[...new Set(target.map(r=>norm(r[0])))], byName=new Map(nameKeys.map(k=>[k,[]]));
const wanted={
 'Bản sao của DỮ LIỆU THÀNH TỰU GAB - 15_24, 13 tháng 9.xlsx':['Data_record_01-01-2022_01-01-20','200 KLG CHUẨN HÓA','DANH SÁCH 560 NGƯỜI DEV XUẤT VỀ'],
 'MÔ TẢ YÊU CẦU NHẬP THÔNG TIN GAB.xlsx':['CSDL GAB','CSDL KLVN','CSDL KLVW','CSDL KL CA','CSDL KLTG'],
 'So sánh.xlsx':['DANH SÁCH 560 NGƯỜI DEV XUẤT VỀ','200 người đã chuẩn','Full dữ liệu thô 200 người']
};
for(const [filename,sheets] of Object.entries(wanted)){const wb=XLSX.readFile(path.join(dir,filename),{raw:false});for(const sn of sheets){if(!wb.Sheets[sn])continue;const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:''});rows.forEach((values,i)=>{const text=norm(values.join(' | '));for(const key of nameKeys)if(text.includes(key))byName.get(key).push({file:filename,sheet:sn,row:i+1,values:values.filter(v=>String(v).trim())});});}}
const results=[];
for(const r of target){const name=String(r[0]),title=String(r[5]), candidates=(byName.get(norm(name))||[]).map(x=>({...x,score:Math.max(0,...x.values.map(v=>sim(title,v)))})).filter(x=>x.score>=.42).sort((a,b)=>b.score-a.score),best=candidates[0];if(!best)continue;const urls=best.values.flatMap(v=>String(v).match(/https?:\/\/[^\s]+/g)||[]),gab=urls.find(u=>/gab\.world\/vi\/bank\//i.test(u)),link=urls.find(u=>!/gab\.world|youtube|youtu\.be|facebook|tiktok/i.test(u)),desc=best.values.filter(v=>String(v).length>180&&!/^https?:/i.test(String(v))&&sim(title,v)<.9).sort((a,b)=>String(b).length-String(a).length)[0];results.push({name,title,score:+best.score.toFixed(3),gab:gab||'',link:link||'',desc:desc?String(desc):'',source:`${best.file} | ${best.sheet} | ${best.row}`});}
const output={matches:results.length,recoverGab:results.filter(x=>x.gab).length,recoverLink:results.filter(x=>x.link).length,recoverDesc:results.filter(x=>x.desc).length,results};fs.writeFileSync(path.join(process.cwd(),'outputs','sheet6_source_matches.json'),JSON.stringify(output,null,2),'utf8');console.log(JSON.stringify({matches:output.matches,recoverGab:output.recoverGab,recoverLink:output.recoverLink,recoverDesc:output.recoverDesc}));
