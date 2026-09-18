import fs from 'node:fs';
import path from 'node:path';

const file = path.join('outputs', 'complete_standard_work', 'xlsx', 'xl', 'worksheets', 'sheet2.xml');
const updates = [
  [346, 'NGÔ THỊ HOÀNG NGÂN', 'ĐÃ ĐỐI CHIẾU: Thành viên sáng lập kiêm Tổng đạo diễn sự kiện xếp hình lá cờ Tổ quốc có số lượng người tham gia cùng lúc đông nhất Việt Nam. Sự kiện được VietKings xác lập Kỷ lục Việt Nam ngày 10/05/2025. Nguồn đối chiếu: hồ sơ GAB, bài Kyluc.vn và bài giới thiệu cá nhân.'],
  [348, 'NGUYỄN NGỌC PHƯƠNG LAM', 'CHƯA TÌM THẤY: Chưa xác định được nội dung thành tựu trên GAB hoặc bài đối chiếu đủ căn cứ.'],
  [351, 'NGUYỄN THỊ KIM OANH', 'ĐÃ ĐỐI CHIẾU MỘT PHẦN: Bà Nguyễn Thị Kim Oanh là Phó Giám đốc/đại diện Công ty TNHH Dũng Tân tại Hội ngộ Kỷ lục gia Việt Nam lần thứ 51 ngày 15/12/2022. Kỷ lục “Vũ điệu Thiên nga - Cây bonsai Tùng Kim Cương đạt giá trị Kỷ lục Độc bản Việt Nam” được trao cho Công ty TNHH Dũng Tân, chưa có căn cứ xác định đây là kỷ lục cá nhân của bà. Nguồn chính thức: Kyluc.vn.'],
  [352, 'NGUYỄN QUANG THẮNG', 'CHƯA TÌM THẤY: Đã có link hồ sơ GAB nhưng chưa xác định được nội dung thành tựu và bài đối chiếu đủ căn cứ.'],
  [357, 'ĐẶNG HỒNG MI', 'CHƯA TÌM THẤY: Đã có link hồ sơ GAB nhưng chưa xác định được nội dung thành tựu và bài đối chiếu đủ căn cứ.'],
  [358, 'NGÔ THU AN', 'NGUỒN NGOÀI KYLUC.VN: Nghệ nhân Bàn tay vàng Thư pháp Ngô Thu An sáng tác tác phẩm “Nhân diện thư chữ Phượng thành chân dung nữ tướng Lê Chân”. Tác phẩm đạt giải Nhất Cuộc thi Thư pháp Hải Phòng năm 2025. Nguồn: Hiệp hội Làng nghề Hải Phòng; chưa xác nhận đây là nội dung Kỷ lục Việt Nam trên GAB.'],
  [359, 'TRẦN QUỐC HUY', 'NGUỒN NGOÀI KYLUC.VN: Đại tá, Thạc sĩ, Nghệ nhân thư pháp Trần Quốc Huy tham gia đóng góp chuyên môn tại Hội thảo “Nhân diện thư, Vật điểu thư trong dòng chảy thư pháp Việt Nam đương đại” ngày 23/03/2025. Nguồn: Đời sống và Phát triển; chưa xác nhận một nội dung Kỷ lục Việt Nam cụ thể trên GAB.'],
  [360, 'NGUYỄN HOÀNG BÁCH', 'NỘI DUNG TRÊN GAB: Xác lập Kỷ lục Việt Nam “Học sinh lớp 1 có khả năng thực hiện dãy 100 phép tính cộng, trừ các số có 3 chữ số ngẫu nhiên trong thời gian nhanh nhất”. Nguyễn Hoàng Bách (sinh năm 2018) có thể hoàn thành 100 phép tính trong 36 giây; mỗi phép tính chỉ hiển thị 1 giây. Chưa tìm được bài Kyluc.vn và ngày xác lập cụ thể.'],
  [361, 'TRẦN HOÀI THUẬN', 'NỘI DUNG TRÊN GAB: Xác lập Kỷ lục Việt Nam “Người thực hiện bộ tượng mini bằng đất sét nặn thủ công (không nung) lấy cảm hứng từ các hình tượng trong văn hóa Phật giáo để lưu niệm có số lượng nhiều nhất (322 tượng)”. Ông Trần Hoài Thuận dành 20 năm thực hiện bộ tượng bằng đất sét tự nhiên, kích thước 4,5-12,5 cm, tạo hình thủ công, phơi nắng và tô màu, không nung. Chưa tìm được bài Kyluc.vn và ngày xác lập cụ thể.'],
  [362, 'LƯƠNG THÀNH NHẬT', 'NỘI DUNG TRÊN GAB: Lương Thành Nhật và Huỳnh Hoàng Sơn xác lập Kỷ lục Việt Nam “Đôi nam Trọng tài Yoga quốc gia thực hiện màn trình diễn Acro Yoga duy trì tư thế ‘Con nhện trên tay’ kết hợp di chuyển liên tục trên quãng đường dài nhất - đạt 57,77m”, trong khuôn khổ Giải Cúp Câu lạc bộ Yoga tỉnh Lâm Đồng lần thứ I năm 2026. Chưa tìm được bài Kyluc.vn và ngày xác lập cụ thể.'],
  [363, 'HUỲNH HOÀNG SƠN', 'NỘI DUNG TRÊN GAB: Lương Thành Nhật và Huỳnh Hoàng Sơn xác lập Kỷ lục Việt Nam “Đôi nam Trọng tài Yoga quốc gia thực hiện màn trình diễn Acro Yoga duy trì tư thế ‘Con nhện trên tay’ kết hợp di chuyển liên tục trên quãng đường dài nhất - đạt 57,77m”, trong khuôn khổ Giải Cúp Câu lạc bộ Yoga tỉnh Lâm Đồng lần thứ I năm 2026. Chưa tìm được bài Kyluc.vn và ngày xác lập cụ thể.'],
];

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const setCell = (rowXml, row, col, value) => {
  const ref = `${col}${row}`;
  const re = new RegExp(`<c\\b[^>]*\\br="${ref}"[^>]*>[\\s\\S]*?<\\/c>|<c\\b[^>]*\\br="${ref}"[^>]*/>`);
  const old = rowXml.match(re)?.[0];
  const style = old?.match(/\bs="(\d+)"/)?.[1] ?? '1';
  const cell = `<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`;
  return old ? rowXml.replace(re, cell) : rowXml.replace('</row>', `${cell}</row>`);
};

let xml = fs.readFileSync(file, 'utf8');
for (const [row, name, content] of updates) {
  const re = new RegExp(`<row\\b[^>]*\\br="${row}"[^>]*>[\\s\\S]*?<\\/row>`);
  const original = xml.match(re)?.[0];
  if (!original || !original.includes(name)) throw new Error(`Không tìm thấy đúng hàng ${name}`);
  xml = xml.replace(re, setCell(original, row, 'T', content));
}
fs.writeFileSync(file, xml, 'utf8');
console.log(`Updated column T for ${updates.length} records`);
