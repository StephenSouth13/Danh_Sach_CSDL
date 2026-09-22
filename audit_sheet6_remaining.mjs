import XLSX from 'xlsx';
const file='D:/KLG/Danh_Sach_CSDL/Dulieutruyxuat/19_09_2026.xlsx';
const wb=XLSX.readFile(file,{raw:false}), data=XLSX.utils.sheet_to_json(wb.Sheets.sheet6,{header:1,defval:''}).slice(4).filter(r=>r[0]);
const bad=v=>/^(CHƯA|KHÔNG|Không)/i.test(String(v).trim());
const rows=data.map((r,i)=>({row:i+5,name:r[0],title:r[5],recordId:r[6],gabId:r[7],province:r[8],type:r[9],time:r[10],link:r[12],desc:r[13]}));
console.log(JSON.stringify({
 badTitles:rows.filter(x=>/không có thành tựu|chưa có tên|xác lập kỷ lục việt nam\s*["“”]*$/i.test(x.title)),
 missingGab:rows.filter(x=>bad(x.gabId)), missingProvince:rows.filter(x=>bad(x.province)), missingType:rows.filter(x=>bad(x.type)), missingTime:rows.filter(x=>bad(x.time)), missingLink:rows.filter(x=>bad(x.link)),
 fallbackDescriptions:rows.filter(x=>String(x.desc).includes('chưa có đủ thông số kỹ thuật chi tiết')).map(x=>({row:x.row,name:x.name,title:x.title}))
},null,2));
