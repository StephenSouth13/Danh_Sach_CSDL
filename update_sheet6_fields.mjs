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
const comparison = XLSX.utils.sheet_to_json(wb.Sheets['ĐỐI CHIẾU GAB - 200'], { header: 1, defval: '' }).slice(6);
const duplicateByName = new Map();
for (const row of comparison) {
  const name = String(row[0] || '').trim();
  if (name && !duplicateByName.has(name)) duplicateByName.set(name, Number(row[12] || 0));
}
const sourceMatchPath = path.join(root, 'outputs', 'sheet6_source_matches.json');
const sourceMatches = fs.existsSync(sourceMatchPath) ? JSON.parse(fs.readFileSync(sourceMatchPath, 'utf8')).results : [];
const sourceMatchMap = new Map(sourceMatches.map(item => [`${item.name}||${item.title}`, item]));
const manualLinks = new Map([
  ['CHU MẠNH NGUYÊN', 'https://kyluc.vn/tin-tuc/ky-luc/nha-giao-chu-manh-nguyen-la-nguoi-cao-tuoi-nhat-nhan-bang-tien-si-khoa-hoc-giao-duc-tai-viet-nam'],
  ['ĐẶNG ĐỨC THÀNH', 'https://kyluc.vn/tin-tuc/ky-luc/bai-tap-vuot-qua-noi-so-ncov-cua-ceo-dang-duc-thanh-xac-lap-ky-luc-viet-nam'],
  ['ĐỒNG XUÂN TRƯỜNG', 'https://kyluc.vn/tin-tuc/ky-luc-viet-nam/doc-dao-vuon-linh-sam-bonsai-da-dang-the-xac-lap-ky-luc-viet-nam'],
  ['HOÀNG THỊ NGỌC MAI', 'https://kyluc.vn/tin-tuc/ky-luc/tap-tho-me-oi-con-nho-me-cua-nha-giao-hoang-thi-ngoc-mai-chinh-thuc-xac-lap-ky-luc-viet-nam-trong-ngay-ra-mat'],
  ['HUỲNH THỊ HOA', 'https://kyluc.vn/tin-tuc/thong-tin/vietkings-discovery-i-ghe-tham-bao-tang-da-ngoc-chau-tai-thanh-pho-bao-loc'],
  ['LƯƠNG VĂN QUANG', 'https://kyluc.vn/tin-tuc/su-kien-ky-luc/hoi-ngo-ky-luc-gia-viet-nam-lan-thu-51-tai-thai-nguyen-dau-an-phat-trien-cua-cong-dong-ky-luc-gia-viet-nam-sau-18-nam'],
  ['ĐINH VĂN TRỌNG', 'https://topplus.vn/tin-tuc/top-su-kien/ve-kien-giang-chiem-nguong-co-xu-ky-my-cay-mai-vang-kieng-co-xu-chay-toan-than-dang-truc-mot-cot-vua-duoc-xac-lap-gia-tri-ky-luc-doc-ban-viet-nam'],
  ['LÊ TRỌNG HUY', 'https://baolaocai.vn/cau-be-lop-3-lap-ky-luc-kinh-ngac-voi-quyen-con-nhi-khuc-post484185.html']
]);
const manualDescriptions = new Map([
  ['CHU MẠNH NGUYÊN', 'Nhà giáo Chu Mạnh Nguyên sinh năm 1944, được cấp bằng Tiến sĩ Khoa học Giáo dục năm 2015 ở tuổi 71. Ngày 01/02/2020, VietKings trao Kỷ lục Việt Nam cho ông với nội dung người cao tuổi nhất được nhận bằng Tiến sĩ Khoa học Giáo dục tại Việt Nam. Thành tựu lan tỏa tinh thần học tập suốt đời và khẳng định giá trị của sự kiên trì trong nghiên cứu, giáo dục.'],
  ['ĐẶNG ĐỨC THÀNH', 'Bài tập “Vượt qua Nỗi sợ nCoV” gồm 7 động tác đơn giản, được phổ biến qua cuộc thi video trên toàn quốc. Tại thời điểm xác lập có hơn 4.000 clip gửi về, sau đó tăng lên hơn 5.300 bài từ 63 tỉnh, thành. Kỷ lục được xác lập ngày 22/03/2022, góp phần khuyến khích rèn luyện sức khỏe và nâng cao sức đề kháng trong giai đoạn dịch bệnh.'],
  ['ĐỒNG XUÂN TRƯỜNG', 'Ông Đồng Xuân Trường cùng nghệ nhân Trương Văn Hoàng nghiên cứu và phát triển khoảng 2.000 cây Linh sam bonsai với nhiều dáng thế sau gần 10 năm thử nghiệm. Bộ sưu tập được VietKings ghi nhận có số lượng Bonsai Linh sam đa dáng thế nhiều nhất, góp phần phát triển nghệ thuật bonsai và tạo hướng đi mới cho cây kiểng Việt Nam.'],
  ['HOÀNG THỊ NGỌC MAI', 'Tập thơ “Mẹ ơi! Con nhớ Mẹ” gồm 30 bài thơ; 60 bức tranh thư pháp được chuyển thể từ các câu thơ, bài thơ với hình ảnh hoa mai làm chủ đạo. Kỷ lục được trao ngày 12/10/2025, tôn vinh tình mẫu tử, nghệ thuật thư pháp và giá trị gắn kết gia đình Việt.'],
  ['HUỲNH THỊ HOA', 'Bộ tượng Thập Bát La Hán được điêu khắc từ gỗ dâu tằm cổ thụ, trưng bày tại Bảo tàng đá Ngọc Châu thuộc Khu du lịch sinh thái Hoa Tài Ngọc Châu, Bảo Lộc. Công trình được công nhận Kỷ lục Việt Nam năm 2018, thể hiện kỹ thuật điêu khắc gỗ quy mô lớn và góp phần bảo tồn, giới thiệu nghệ thuật tạo tác truyền thống.']
]);
const manualDescriptionsByKey = new Map([
  ['ĐINH VĂN TRỌNG', 'Cây mai vàng “Cổ Xù Kỳ Mỹ” thuộc giống mai xù, có tuổi thọ khoảng 100 năm, nổi bật với thân cổ xù chảy toàn thân và dáng trực một cốt. Ngày 29/09/2024 tại Rạch Giá, Kiên Giang, VietKings xác lập giá trị Kỷ lục Độc bản Việt Nam cho tác phẩm. Việc ghi nhận góp phần tôn vinh nghệ thuật mai kiểng cổ Nam Bộ và khuyến khích bảo tồn các giống cây cảnh có giá trị lâu năm.'],
  ['LÊ TRỌNG HUY', 'Lê Trọng Huy được VietKings ghi nhận là cậu bé nhỏ tuổi nhất biểu diễn thành công 6 thử thách với quyền côn nhị khúc tại chương trình Siêu tài năng nhí năm 2020. Thành tích thể hiện quá trình luyện tập kỹ thuật, khả năng kiểm soát côn và bản lĩnh biểu diễn ở độ tuổi nhỏ, góp phần lan tỏa tinh thần rèn luyện võ thuật trong thanh thiếu nhi.'],
  ['TRÚC PHƯƠNG', 'Trường ca sử thi “Mẹ, Đất nước và Lưu dân” tái hiện các cuộc chiến tranh cùng lịch sử mở cõi về phương Nam của người Việt bằng hình thức trường ca có quy mô lớn. Tác phẩm kết hợp giá trị văn học với tư liệu lịch sử, góp phần lưu giữ ký ức cộng đồng, tôn vinh hành trình dựng nước, giữ nước và quá trình hình thành vùng đất phương Nam.']
]);
const manualLinkRules = [
  ['NGUYỄN ĐÌNH TRANH', '', 'https://kyluc.vn/tin-tuc/ky-luc/vietkings-values-nguoi-sang-tac-tho-ve-cac-loai-hoa-nhieu-nhat-viet-nam'],
  ['NGUYỄN ĐỨC HIỀN', 'PHỦ TIÊN HƯƠNG', 'https://kyluc.vn/tin-tuc/ky-luc/phu-tien-huong-phu-tho-mau-tai-huyen-quoc-oai-tp-ha-noi-tro-thanh-diem-den-tam-linh-dat-ky-luc-viet-nam'],
  ['NGUYỄN ĐỨC HIỀN', 'English Olympics', 'https://kyluc.vn/tin-tuc/thong-tin/cuoc-thi-english-olympics-of-vietnam-2019-vong-so-tuyen-offline-chinh-thuc-bat-dau'],
  ['NGUYỄN ĐỨC LONG', '', 'https://kyluc.vn/tin-tuc/danh-muc-de-xuat/dao-dien-thuc-hien-mo-hinh-cot-moc-chu-quyen-hoang-sa-va-truong-sa-cua-viet-nam'],
  ['NGUYỄN THẾ VINH', '', 'https://kyluc.vn/tin-tuc/ky-luc-viet-nam/chiem-nguong-bo-suu-tap-dien-thoai-phien-ban-dac-biet-va-gioi-han-dat-ky-luc-vn-cua-anh-nguyen-the-vinh'],
  ['NGUYỄN THỊ THANH TÂM', 'Châu Á', 'https://kyluc.vn/tin-tuc/ky-luc-chau-a/hoi-ngo-ky-luc-gia-viet-nam-lan-thu-52-du-nang-hoa-se-no-du-tam-nhin-thay-co-hoi-03-ky-luc-gia-don-nhan-ky-luc-chau-a-moi'],
  ['NGUYỄN THỊ THANH TÂM', '', 'https://kyluc.vn/tin-tuc/van-hoa-nghe-thuat/nha-suu-tap-nguyen-thi-thanh-tam-xac-lap-ky-luc-viet-nam-voi-bo-suu-tap-sen-trong-doi-song-van-hoa-viet'],
  ['PHẠM PHƯƠNG PHI', '', 'https://kyluc.vn/tin-tuc/thong-tin/dau-dua-dr-phi-giai-phap-tu-nhien-vi-suc-khoe-cong-dong'],
  ['PHẠM LÊ QUỐC CƯỜNG', '', 'https://quangducxua.com/dong-thoi-gian/'],
  ['PHÙNG TUẤN GIANG', 'Củ sâm', 'https://kyluc.vn/tin-tuc/ky-luc/worldkings-ky-luc-the-gioi-cua-viet-nam-cu-sam-ngoc-linh-viet-nam-lon-nhat-the-gioi'],
  ['PHÙNG TUẤN GIANG', 'Nam y', 'https://kyluc.vn/tin-tuc/ky-luc-the-gioi/tien-si-luong-y-ky-luc-gia-phung-tuan-giang-don-nhan-dia-vang-cong-hien-tu-vien-ky-luc-the-gioi']
];
const findManualLink = (name, title) => manualLinkRules.find(([ruleName, contains]) => ruleName === name && (!contains || title.includes(contains)))?.[2] || '';

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
const validArticle = value => {
  const text = String(value ?? '').trim();
  return /^https?:\/\//i.test(text) ? text : '';
};
const recordTypeLabel = value => {
  const text = String(value || '').toUpperCase();
  if (text.includes('THẾ GIỚI')) return 'Kỷ lục Thế giới';
  if (text.includes('CHÂU Á')) return 'Kỷ lục Châu Á';
  if (text.includes('ĐÔNG DƯƠNG')) return 'Kỷ lục Đông Dương';
  if (text.includes('TOP')) return 'Top Việt Nam';
  return 'Kỷ lục Việt Nam';
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
  if (/^Không có thành tựu thiếu$/i.test(achievement)) {
    updates.set(index + 1, Object.fromEntries('ABCDEFGHIJKLMN'.split('').map(column => [column, ''])));
    continue;
  }
  const candidates = standardByName.get(norm(name)) || [];
  const duplicateCount = duplicateByName.get(name) || 0;
  const uniqueStandardCount = Math.max(Number(row[2] || 0) - duplicateCount, 0);
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
  if (!gabId) { gabId = 'CHƯA CÓ HỒ SƠ GAB TRONG CÁC NGUỒN ĐÃ KIỂM TRA'; unresolvedGabIds++; }

  const rawStandardTitle = valid(std[13], '');
  const titleIsBroken = !rawStandardTitle || /^(chưa có tên|xác lập kỷ lục việt nam\s*["“”]*)$/i.test(rawStandardTitle);
  const recordName = valid(std[4], achievement);
  const desiredTitle = titleIsBroken ? `Xác lập ${recordTypeLabel(std[10])} “${recordName}”` : rawStandardTitle;
  const sourceMatch = sourceMatchMap.get(`${name}||${achievement}`);
  const standardLink = validArticle(std[14]);
  const standardDescription = valid(std[15], '');
  const articleLink = standardLink || (gabMatch ? validArticle(gabMatch[5]) : '') || findManualLink(name, desiredTitle) || manualLinks.get(name) || (sourceMatch?.score >= .65 ? sourceMatch.link : '') || 'CHƯA TÌM THẤY BÀI VIẾT PHÙ HỢP';
  const researchedDescription = standardDescription || (gabMatch ? valid(gabMatch[6], '') : '') || manualDescriptions.get(name) || (sourceMatch?.score >= .65 ? sourceMatch.desc : '');
  let ruleDescription = '';
  if (name === 'NGUYỄN ĐỨC LONG') ruleDescription = 'Hai mô hình cột mốc chủ quyền Hoàng Sa và Trường Sa có cùng kích thước: cao 3 m, cạnh đáy rộng 0,8 m; được tạo hình bằng thép inox 304, đồng và các phụ kiện khác. Công trình thể hiện tình yêu quê hương, khẳng định ý thức về chủ quyền biển đảo và có giá trị tuyên truyền, giáo dục cộng đồng.';
  if (name === 'NGUYỄN THỊ THANH TÂM') ruleDescription = 'Bộ sưu tập “Sen trong đời sống văn hóa Việt” được hình thành trong hơn 20 năm và ở thời điểm xác lập Kỷ lục Châu Á gồm 433 hiện vật: 258 ảnh và tranh sen trên nhiều chất liệu, 106 lọ hoa, đồ dùng và đồ trang trí, cùng 69 tác phẩm thư pháp và câu đối về sen. Thành tựu góp phần quảng bá biểu tượng sen và các giá trị văn hóa Việt Nam.';
  if (name === 'NGUYỄN THẾ VINH') ruleDescription = 'Bộ sưu tập gồm các mẫu điện thoại phiên bản đặc biệt và giới hạn của nhiều thương hiệu, được VietKings trao Kỷ lục Việt Nam ngày 21/09/2025. Bộ sưu tập lưu giữ những sản phẩm gắn với lịch sử thiết kế và phát triển công nghệ, thể hiện sự kiên trì sưu tầm và trân trọng giá trị của các thiết bị qua từng thời kỳ.';
  if (name === 'PHẠM LÊ QUỐC CƯỜNG') ruleDescription = 'Bộ sưu tập gốm Quảng Đức xưa được lưu giữ tại không gian nhà cổ của gia đình ông Phạm Lê Quốc Cường ở Phú Yên, gồm nhiều sản phẩm dân dụng và mỹ thuật đặc trưng của làng gốm như bình, lọ, chum, chóe, nậm rượu, bình vôi và chậu. Việc sưu tầm, trưng bày góp phần bảo tồn dấu tích làng nghề, giới thiệu kỹ thuật men và hoa văn đặc sắc, đồng thời lan tỏa giá trị văn hóa của dòng gốm Quảng Đức đến cộng đồng.';
  if (name === 'PHÙNG TUẤN GIANG' && desiredTitle.includes('Củ sâm')) ruleDescription = 'Củ sâm Ngọc Linh thuộc sở hữu của lương y Phùng Tuấn Giang có chiều dài khoảng 80 cm, ngang 35 cm, nặng 2,25 kg và được giới thiệu có tuổi đời khoảng 156 năm. Thành tựu được ghi nhận ở cấp thế giới, góp phần quảng bá giá trị dược liệu đặc hữu của Việt Nam và nâng cao nhận thức về bảo tồn nguồn sâm Ngọc Linh.';
  const description = ruleDescription || manualDescriptionsByKey.get(name) || researchedDescription || `Theo hồ sơ chuẩn, thành tựu được ghi nhận với nội dung: ${desiredTitle}. Thời gian xác lập: ${valid(std[12])}. Hồ sơ hiện chưa có đủ thông số kỹ thuật chi tiết; cần bổ sung từ quyết định xác lập hoặc bài viết chính thức để tránh suy diễn dữ liệu.`;
  const update = {
    B: `Chuẩn ${Number(row[2] || 0)} dòng | ${uniqueStandardCount} thành tựu không trùng | GAB đã có ${Number(row[3] || 0)} | Thiếu thật ${Number(row[4] || 0)} | Trùng chuẩn ${duplicateCount}`,
    F: desiredTitle,
    G: recordId,
    H: gabId,
    I: valid(std[8]),
    J: valid(std[10]),
    K: valid(std[12]),
    L: desiredTitle,
    M: articleLink,
    N: description
  };
  if (gabMatch) {
    update.B = `Chuẩn ${Number(row[2] || 1)} dòng | ${uniqueStandardCount} thành tựu không trùng | GAB đã có ${Math.max(Number(row[3] || 0), 1)} | Thiếu thật ${Math.max(Number(row[4] || 1) - 1, 0)} | Trùng chuẩn ${duplicateCount} | Cần chuẩn hóa dữ liệu`;
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
const headers = { C: 'TỔNG DÒNG CHUẨN', I: 'TỈNH THÀNH MỚI', J: 'LOẠI KỶ LỤC', K: 'THỜI GIAN XÁC LẬP', L: 'TIÊU ĐỀ THÀNH TỰU TRÊN GAB', M: 'LINK BÀI VIẾT TRÊN GAB', N: 'THÔNG SỐ KỸ THUẬT + THỜI GIAN + Ý NGHĨA/GIÁ TRỊ/ẢNH HƯỞNG TÍCH CỰC' };
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
