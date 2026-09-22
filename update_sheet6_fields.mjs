import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import XLSX from 'xlsx';

const root = process.cwd();
const file = path.join(root, 'Dulieutruyxuat', '19_09_2026.xlsx');
const work = path.join(root, 'outputs', 'sheet6_update');
const inputZip = path.join(root, 'outputs', 'sheet6_input.zip');
const outputZip = path.join(root, 'outputs', 'sheet6_update.zip');
fs.rmSync(work, { recursive: true, force: true });
fs.mkdirSync(work, { recursive: true });
fs.copyFileSync(file, inputZip);
execFileSync('powershell', ['-NoProfile', '-Command', `Expand-Archive -LiteralPath '${inputZip.replace(/'/g, "''")}' -DestinationPath '${work.replace(/'/g, "''")}' -Force`]);

const wb = XLSX.readFile(file, { raw: false });
const target = XLSX.utils.sheet_to_json(wb.Sheets.sheet6, { header: 1, defval: '' });
const standard = XLSX.utils.sheet_to_json(wb.Sheets['Danh sách 200 người đã chuẩn'], { header: 1, defval: '' }).slice(2);
const gab = XLSX.utils.sheet_to_json(wb.Sheets['DANH SÁCH ĐẦY ĐỦ TRÊN GAB'], { header: 1, defval: '' }).slice(1);

const norm = value => String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/gi, 'd').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
const tokens = value => new Set(norm(value).split(/\s+/).filter(token => token.length > 2));
const score = (left, right) => {
  const a = tokens(left), b = tokens(right);
  if (!a.size || !b.size) return 0;
  let common = 0;
  for (const token of a) if (b.has(token)) common++;
  return common / Math.max(a.size, b.size);
};
const valid = (value, fallback = 'CHƯA CÓ') => {
  const text = String(value ?? '').trim();
  return text && !/^(CHƯA CÓ|KHÔNG CÓ)$/i.test(text) ? text : fallback;
};
const standardByName = new Map();
for (const row of standard) {
  const key = norm(row[6]);
  if (!key) continue;
  if (!standardByName.has(key)) standardByName.set(key, []);
  standardByName.get(key).push(row);
}
const gabByProfile = new Map(), gabByName = new Map();
for (const row of gab) {
  const profile = String(row[1] || '').trim();
  const name = norm(row[2]);
  if (profile) { if (!gabByProfile.has(profile)) gabByProfile.set(profile, []); gabByProfile.get(profile).push(row); }
  if (name) { if (!gabByName.has(name)) gabByName.set(name, []); gabByName.get(name).push(row); }
}

const updates = new Map();
let matchedRecordIds = 0, unresolvedGabIds = 0, weakStandardMatches = 0;
for (let index = 4; index < target.length; index++) {
  const row = target[index];
  const name = String(row[0] || '').trim();
  const achievement = String(row[5] || '').trim();
  if (!name || !achievement) continue;
  const candidates = standardByName.get(norm(name)) || [];
  const ranked = candidates.map(candidate => ({ candidate, value: Math.max(score(achievement, candidate[13]), score(achievement, candidate[4])) })).sort((a, b) => b.value - a.value);
  const std = ranked[0]?.candidate || [];
  if (!ranked.length || ranked[0].value < 0.45) weakStandardMatches++;

  const currentGabId = /^https?:\/\/gab\.world/i.test(String(row[7] || '').trim()) ? String(row[7]).trim() : '';
  const standardGabId = /^https?:\/\/gab\.world/i.test(String(std[5] || '').trim()) ? String(std[5]).trim() : '';
  let gabId = currentGabId || standardGabId;
  let gabCandidates = gabId ? (gabByProfile.get(gabId) || []) : (gabByName.get(norm(name)) || []);
  if (!gabId && gabCandidates.length) gabId = String(gabCandidates[0][1] || '').trim();
  const gabRanked = gabCandidates.map(candidate => ({ candidate, value: score(achievement, candidate[4]) })).sort((a, b) => b.value - a.value);
  const gabMatch = gabRanked[0] && gabRanked[0].value >= 0.78 ? gabRanked[0].candidate : null;
  const recordId = gabMatch ? valid(gabMatch[0], 'KHÔNG CÓ') : 'KHÔNG CÓ';
  if (gabMatch) matchedRecordIds++;
  if (!gabId) { gabId = 'KHÔNG TÌM THẤY HỒ SƠ GAB'; unresolvedGabIds++; }

  const update = {
    G: recordId,
    H: gabId,
    I: valid(std[8]),
    J: valid(std[10]),
    K: valid(std[12]),
    L: valid(std[13], achievement),
    M: valid(std[14]),
    N: valid(std[15])
  };
  if (gabMatch) {
    update.B = `Chuẩn ${Number(row[2] || 1)} | GAB đã có ${Math.max(Number(row[3] || 0), 1)} | Thiếu ${Math.max(Number(row[4] || 1) - 1, 0)} | Cần chuẩn hóa dữ liệu`;
    update.D = Math.max(Number(row[3] || 0), 1);
    update.E = Math.max(Number(row[4] || 1) - 1, 0);
  }
  updates.set(index + 1, update);
}

const xmlPath = path.join(work, 'xl', 'worksheets', 'sheet5.xml');
let xml = fs.readFileSync(xmlPath, 'utf8');
const escapeXml = value => String(value ?? '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const cellRegex = ref => new RegExp(`<c\\b([^>]*\\br="${ref}"[^>]*)(?:\\/>|>([\\s\\S]*?)<\\/c>)`);
const getStyle = (rowXml, preferredRefs) => {
  for (const ref of preferredRefs) {
    const match = rowXml.match(cellRegex(ref));
    const style = match?.[1]?.match(/\bs="(\d+)"/)?.[1];
    if (style) return style;
  }
  return '3';
};
const makeCell = (ref, value, style) => `<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`;
const replaceCell = (rowXml, ref, value, style) => {
  const cell = makeCell(ref, value, style);
  const regex = cellRegex(ref);
  if (regex.test(rowXml)) return rowXml.replace(regex, cell);
  return rowXml.replace('</row>', `${cell}</row>`);
};
const headers = { I: 'TỈNH THÀNH MỚI', J: 'LOẠI KỶ LỤC', K: 'THỜI GIAN XÁC LẬP', L: 'TIÊU ĐỀ THÀNH TỰU TRÊN GAB', M: 'LINK BÀI VIẾT TRÊN GAB', N: 'THÔNG SỐ KỸ THUẬT + THỜI GIAN + Ý NGHĨA/GIÁ TRỊ/ẢNH HƯỞNG TÍCH CỰC' };
xml = xml.replace(/<row\b([^>]*)\br="4"([^>]*)>[\s\S]*?<\/row>/, rowBlock => {
  const style = getStyle(rowBlock, ['H4', 'G4', 'F4']);
  for (const [column, value] of Object.entries(headers)) rowBlock = replaceCell(rowBlock, `${column}4`, value, style);
  return rowBlock;
});
for (const [rowNumber, values] of updates) {
  const rowPattern = new RegExp(`<row\\b([^>]*)\\br="${rowNumber}"([^>]*)>[\\s\\S]*?<\\/row>`);
  xml = xml.replace(rowPattern, rowBlock => {
    const idStyle = getStyle(rowBlock, [`G${rowNumber}`, `H${rowNumber}`, `F${rowNumber}`]);
    const dataStyle = getStyle(rowBlock, [`F${rowNumber}`, `E${rowNumber}`]);
    for (const [column, value] of Object.entries(values)) rowBlock = replaceCell(rowBlock, `${column}${rowNumber}`, value, column === 'G' || column === 'H' ? idStyle : dataStyle);
    return rowBlock;
  });
}
fs.writeFileSync(xmlPath, xml, 'utf8');
fs.rmSync(outputZip, { force: true });
execFileSync('powershell', ['-NoProfile', '-Command', `Add-Type -AssemblyName System.IO.Compression.FileSystem; [IO.Compression.ZipFile]::CreateFromDirectory('${work.replace(/'/g, "''")}','${outputZip.replace(/'/g, "''")}')`]);
fs.copyFileSync(outputZip, file);
console.log(JSON.stringify({ file, updatedRows: updates.size, matchedRecordIds, unresolvedGabIds, weakStandardMatches }, null, 2));
