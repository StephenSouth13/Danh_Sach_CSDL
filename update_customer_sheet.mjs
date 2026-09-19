import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import XLSX from 'xlsx';

const root=process.cwd();
const file=path.join(root,'Dulieutruyxuat','19_09_2026.xlsx');
const temp=path.join(root,'outputs','customer_sheet_update');
const zip=path.join(root,'outputs','customer_sheet_update.zip');
fs.rmSync(temp,{recursive:true,force:true}); fs.mkdirSync(temp,{recursive:true});
const inputZip=path.join(root,'outputs','customer_sheet_input.zip'); fs.copyFileSync(file,inputZip);
execFileSync('powershell',['-NoProfile','-Command',`Expand-Archive -LiteralPath '${inputZip.replace(/'/g,"''")}' -DestinationPath '${temp.replace(/'/g,"''")}' -Force`]);

const wb=XLSX.readFile(file,{raw:false});
const ws=wb.Sheets['ĐỐI CHIẾU GAB - 200'];
const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:''});
const data=rows.slice(6).filter(r=>String(r[0]||'').trim());
const groups=new Map();
for(const r of data){const name=String(r[0]).trim();if(!groups.has(name))groups.set(name,[]);groups.get(name).push(r)}
const uniq=xs=>[...new Set(xs.map(x=>String(x??'').trim()).filter(Boolean))];
const out=[
 ['DANH SÁCH KLG CÓ DỮ LIỆU THIẾU / SAI SO VỚI GAB'],
 ['Mỗi KLG chỉ có 1 dòng. Chỉ liệt kê phần cần bổ sung hoặc xác minh; thành tựu đã khớp không hiển thị.'],
 ['Màu đỏ: còn thiếu thành tựu hoặc chưa có hồ sơ GAB | Màu cam: sai, lệch, trùng hoặc cần xác minh.'],
 ['HỌ TÊN KLG','KẾT LUẬN NGẮN','TỔNG THÀNH TỰU CHUẨN','ĐÃ CÓ TRÊN GAB','SỐ THÀNH TỰU THIẾU','CỤ THỂ THÀNH TỰU THIẾU','SAI / LỆCH CẦN XÁC MINH']
];
for(const [name,rs] of groups){
 const f=rs[0],total=Number(f[5]||0),present=Number(f[6]||0),missing=Number(f[7]||0),extra=Number(f[8]||0),dup=Number(f[12]||0);
 const statuses=uniq(rs.map(r=>r[1]));
 const issues=uniq(rs.map(r=>r[2])).filter(x=>x!=='Các trường chính đã khớp.');
 const problematic=missing>0||extra>0||dup>0||statuses.some(x=>x!=='ĐỦ - KHỚP')||issues.length>0;
 if(!problematic)continue;
 const conclusion=`Chuẩn ${total} | GAB đã có ${present} | Thiếu ${missing}`;
 const rawMissing=String(f[10]||'').trim();
 const missingItems=rawMissing ? rawMissing.split(/\n(?=\s*\d+\.\s*)/).map(x=>x.replace(/^\s*\d+\.\s*/, '').trim()).filter(Boolean) : [];
 const issueText=issues.join('\n')||statuses.filter(x=>x!=='ĐỦ - KHỚP').join('\n');
 if(missingItems.length){
   for(const achievement of missingItems)out.push([name,conclusion,total,present,missing,achievement,issueText]);
 }else if(issueText||extra>0||dup>0){
   const extraNote=extra>0?`GAB có thêm ${extra} thành tựu cần xác minh.`:'';
   const dupNote=dup>0?`Tab chuẩn có ${dup} dòng trùng cần xác minh.`:'';
   out.push([name,conclusion,total,present,missing,'Không có thành tựu thiếu',[issueText,extraNote,dupNote].filter(Boolean).join('\n')]);
 }
}
const esc=s=>String(s??'').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
const col=n=>{let s='';while(n){const r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26)}return s};
const body=out.map((r,ri)=>{const rn=ri+1;const cells=r.map((v,ci)=>{if(v===''||v==null)return'';let st=rn===4?1:rn<4?2:3;if(rn>4){if(Number(r[4])>0)st=6;else st=7;if(ci===0||ci===1)st=Number(r[4])>0?6:7}return `<c r="${col(ci+1)}${rn}" s="${st}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`}).join('');return `<row r="${rn}" ht="${rn===4?54:rn<4?28:76}" customHeight="1">${cells}</row>`}).join('');
const widths=[28,48,16,16,16,88,80].map((w,i)=>`<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`).join('');
const xml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:G${out.length}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane xSplit="2" ySplit="4" topLeftCell="C5" activePane="bottomRight" state="frozen"/></sheetView></sheetViews><cols>${widths}</cols><sheetData>${body}</sheetData><autoFilter ref="A4:G${out.length}"/><mergeCells count="3"><mergeCell ref="A1:G1"/><mergeCell ref="A2:G2"/><mergeCell ref="A3:G3"/></mergeCells></worksheet>`;
fs.writeFileSync(path.join(temp,'xl','worksheets','sheet4.xml'),xml,'utf8');
fs.rmSync(zip,{force:true});
execFileSync('powershell',['-NoProfile','-Command',`Add-Type -AssemblyName System.IO.Compression.FileSystem; [IO.Compression.ZipFile]::CreateFromDirectory('${temp.replace(/'/g,"''")}','${zip.replace(/'/g,"''")}')`]);
fs.copyFileSync(zip,file);
console.log(JSON.stringify({file,customerRows:out.length-4,sourceRows:data.length,sourcePeople:groups.size},null,2));
