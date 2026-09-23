import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import XLSX from 'xlsx';

const root=process.cwd();
const file=path.join(root,'Dulieutruyxuat','19_09_2026.xlsx');
const work=path.join(root,'outputs','names_tab_update');
const inputZip=path.join(root,'outputs','names_tab_input.zip');
const outputZip=path.join(root,'outputs','names_tab_update.zip');
fs.rmSync(work,{recursive:true,force:true});fs.mkdirSync(work,{recursive:true});fs.copyFileSync(file,inputZip);
execFileSync('powershell',['-NoProfile','-Command',`Expand-Archive -LiteralPath '${inputZip.replace(/'/g,"''")}' -DestinationPath '${work.replace(/'/g,"''")}' -Force`]);

const wb=XLSX.readFile(file,{raw:false});
const source=XLSX.utils.sheet_to_json(wb.Sheets['Danh sách 200 người đã chuẩn'],{header:1,defval:''});
const names=[...new Set(source.slice(2).map(row=>String(row[6]||'').trim()).filter(Boolean))];
const esc=value=>String(value??'').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g,'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
const col=n=>{let s='';while(n){const r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26)}return s};
const rows=[
 ['DANH SÁCH TÊN KỶ LỤC GIA LẤY TỪ TAB CHUẨN'],
 [`Tổng số Kỷ lục gia: ${names.length}`],
 ['STT','HỌ VÀ TÊN KỶ LỤC GIA'],
 ...names.map((name,index)=>[index+1,name])
];
const data=rows.map((row,ri)=>`<row r="${ri+1}" ht="${ri<3?28:22}" customHeight="1">${row.map((value,ci)=>`<c r="${col(ci+1)}${ri+1}" s="${ri===2?1:ri<2?2:3}" t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`).join('')}</row>`).join('');
const xml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:B${rows.length}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomRight" state="frozen"/></sheetView></sheetViews><cols><col min="1" max="1" width="8" customWidth="1"/><col min="2" max="2" width="36" customWidth="1"/></cols><sheetData>${data}</sheetData><autoFilter ref="A3:B${rows.length}"/></worksheet>`;
fs.writeFileSync(path.join(work,'xl','worksheets','sheet6.xml'),xml,'utf8');
const workbookPath=path.join(work,'xl','workbook.xml');let workbookXml=fs.readFileSync(workbookPath,'utf8');workbookXml=workbookXml.replace('</sheets>','<sheet sheetId="6" name="DANH SÁCH TÊN KLG" state="visible" r:id="rId9"/></sheets>');fs.writeFileSync(workbookPath,workbookXml,'utf8');
const relsPath=path.join(work,'xl','_rels','workbook.xml.rels');let rels=fs.readFileSync(relsPath,'utf8');rels=rels.replace('</Relationships>','<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet6.xml"/></Relationships>');fs.writeFileSync(relsPath,rels,'utf8');
const typesPath=path.join(work,'[Content_Types].xml');let types=fs.readFileSync(typesPath,'utf8');types=types.replace('</Types>','<Override PartName="/xl/worksheets/sheet6.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');fs.writeFileSync(typesPath,types,'utf8');
fs.rmSync(outputZip,{force:true});execFileSync('powershell',['-NoProfile','-Command',`Add-Type -AssemblyName System.IO.Compression.FileSystem; [IO.Compression.ZipFile]::CreateFromDirectory('${work.replace(/'/g,"''")}','${outputZip.replace(/'/g,"''")}')`]);fs.copyFileSync(outputZip,file);console.log(JSON.stringify({file,uniqueNames:names.length,rows:names.length},null,2));
