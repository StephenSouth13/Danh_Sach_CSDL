import fs from 'node:fs';
import path from 'node:path';

const file = path.join('outputs', 'complete_standard_work', 'xlsx', 'xl', 'worksheets', 'sheet2.xml');
const updates = [
  {
    row: 351,
    name: 'NGUYỄN THỊ KIM OANH',
    values: {
      D: '15/12/2022',
      E: 'Đại diện Công ty TNHH Dũng Tân tại Hội ngộ Kỷ lục gia Việt Nam lần thứ 51; đơn vị được trao Kỷ lục “Vũ điệu Thiên nga - Cây bonsai Tùng Kim Cương đạt giá trị Kỷ lục Độc bản Việt Nam”',
      K: 'ĐẠI DIỆN ĐƠN VỊ SỞ HỮU KỶ LỤC VIỆT NAM',
      M: '15/12/2022',
      N: 'Bà Nguyễn Thị Kim Oanh - Phó Giám đốc Công ty TNHH Dũng Tân, đại diện đơn vị tại Hội ngộ Kỷ lục gia Việt Nam lần thứ 51. Kỷ lục được trao cho Công ty TNHH Dũng Tân, không phải kỷ lục cá nhân.',
      O: 'https://kyluc.vn/tin-tuc/su-kien-ky-luc/hoi-ngo-ky-luc-gia-viet-nam-lan-thu-51-tai-thai-nguyen-dau-an-phat-trien-cua-cong-dong-ky-luc-gia-viet-nam-sau-18-nam',
      S: 'NGUỒN CHÍNH THỨC KYLUC.VN. Bài viết xác nhận vai trò Phó Giám đốc/đại diện Công ty TNHH Dũng Tân; thành tựu thuộc Công ty, chưa có căn cứ ghi là kỷ lục cá nhân của bà Nguyễn Thị Kim Oanh.',
    },
  },
  {
    row: 358,
    name: 'NGÔ THU AN',
    values: {
      D: '10/09/2025',
      E: 'Tác phẩm “Nhân diện thư chữ Phượng thành chân dung nữ tướng Lê Chân” đạt giải Nhất Cuộc thi Thư pháp Hải Phòng năm 2025',
      K: 'THÀNH TỰU NGHỆ THUẬT - NGUỒN NGOÀI',
      M: '10/09/2025',
      N: 'Nghệ nhân Bàn tay vàng Thư pháp Ngô Thu An sáng tác tác phẩm “Nhân diện thư chữ Phượng thành chân dung nữ tướng Lê Chân”; tác phẩm đạt giải Nhất Cuộc thi Thư pháp Hải Phòng năm 2025.',
      O: 'https://hiephoilangnghehaiphong.com/nghe-nhan-ban-tay-vang-thu-phap-ngo-thu-an-va-tac-pham-chu-phuong-thanh-chan-dung-nu-tuong-le-chan/',
      S: 'NGUỒN NGOÀI KYLUC.VN: Hiệp hội Làng nghề Hải Phòng. Nguồn xác nhận danh xưng Nghệ nhân Bàn tay vàng, tác phẩm và giải Nhất; chưa phải bài xác lập Kỷ lục Việt Nam.',
    },
  },
  {
    row: 359,
    name: 'TRẦN QUỐC HUY',
    values: {
      D: '23/03/2025',
      E: 'Đóng góp chuyên môn tại Hội thảo “Nhân diện thư, Vật điểu thư trong dòng chảy thư pháp Việt Nam đương đại”',
      K: 'HOẠT ĐỘNG CHUYÊN MÔN - NGUỒN NGOÀI',
      M: '23/03/2025',
      N: 'Đại tá, Thạc sĩ, Nghệ nhân thư pháp Trần Quốc Huy tham gia đóng góp chuyên môn tại Hội thảo về Nhân diện thư và Vật điểu thư trong dòng chảy thư pháp Việt Nam đương đại.',
      O: 'https://doisongvaphattrien.vn/hoi-thao-nhan-dien-thu-vat-dieu-thu-trong-dong-chay-thu-phap-viet-nam-duong-dai-a47215.html',
      S: 'NGUỒN NGOÀI KYLUC.VN: Đời sống và Phát triển. Nguồn dùng để nhận diện đúng Trần Quốc Huy và hoạt động chuyên môn; chưa xác nhận một nội dung Kỷ lục Việt Nam cụ thể.',
    },
  },
  {
    row: 360,
    name: 'NGUYỄN HOÀNG BÁCH',
    values: {
      D: '',
      E: 'Học sinh lớp 1 có khả năng thực hiện dãy 100 phép tính cộng, trừ các số có 3 chữ số ngẫu nhiên trong thời gian nhanh nhất',
      K: 'KỶ LỤC VIỆT NAM',
      M: '',
      N: 'Xác lập Kỷ lục Việt Nam “Học sinh lớp 1 có khả năng thực hiện dãy 100 phép tính cộng, trừ các số có 3 chữ số ngẫu nhiên trong thời gian nhanh nhất”; hoàn thành 100 phép tính trong 36 giây.',
      O: 'https://gab.world/vi/bank/9Wdfa7Tt3iPVj97W8HLlTJMCInf1',
      S: 'NGUỒN GAB do người dùng cung cấp. Học sinh Nguyễn Hoàng Bách (sinh năm 2018) hoàn thành 100 phép tính trong 36 giây; chưa tìm được bài Kyluc.vn và ngày xác lập cụ thể.',
    },
  },
  {
    row: 361,
    name: 'TRẦN HOÀI THUẬN',
    values: {
      E: 'Người thực hiện bộ tượng mini bằng đất sét nặn thủ công (không nung) lấy cảm hứng từ các hình tượng trong văn hóa Phật giáo để lưu niệm có số lượng nhiều nhất (322 tượng)',
      K: 'KỶ LỤC VIỆT NAM',
      N: 'Xác lập Kỷ lục Việt Nam “Người thực hiện bộ tượng mini bằng đất sét nặn thủ công (không nung) lấy cảm hứng từ các hình tượng trong văn hóa Phật giáo để lưu niệm có số lượng nhiều nhất (322 tượng)”.',
      O: 'https://gab.world/vi/bank/tUgJufFNamdSohTUxhYaddwmT972',
      S: 'NGUỒN GAB do người dùng cung cấp. Ông Trần Hoài Thuận ở xã Tân Phú, tỉnh Vĩnh Long dành 20 năm thực hiện bộ tượng; chưa tìm được bài Kyluc.vn và ngày xác lập cụ thể.',
    },
  },
  ...[361 + 1, 361 + 2].map((row, index) => ({
    row,
    name: index === 0 ? 'LƯƠNG THÀNH NHẬT' : 'HUỲNH HOÀNG SƠN',
    values: {
      D: 'Năm 2026 (chưa rõ ngày)',
      E: 'Đôi nam Trọng tài Yoga quốc gia thực hiện màn trình diễn Acro Yoga duy trì tư thế “Con nhện trên tay” kết hợp di chuyển liên tục trên quãng đường dài nhất - đạt 57,77m',
      K: 'KỶ LỤC VIỆT NAM',
      M: '2026',
      N: 'Xác lập Kỷ lục Việt Nam “Đôi nam Trọng tài Yoga quốc gia thực hiện màn trình diễn Acro Yoga duy trì tư thế ‘Con nhện trên tay’ kết hợp di chuyển liên tục trên quãng đường dài nhất - đạt 57,77m”, trong khuôn khổ Giải Cúp Câu lạc bộ Yoga tỉnh Lâm Đồng lần thứ I năm 2026.',
      O: index === 0 ? 'https://gab.world/vi/bank/R0eSAPU7GRaQGYSbQTs5mFwmBsu1' : 'https://gab.world/vi/bank/AAs07f59GsRj7xFcBYwDteuOISp2',
      S: 'NGUỒN GAB do người dùng cung cấp. Thành tựu đồng sở hữu bởi Lương Thành Nhật và Huỳnh Hoàng Sơn; chưa tìm được bài Kyluc.vn và ngày xác lập cụ thể.',
    },
  })),
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
for (const update of updates) {
  const re = new RegExp(`<row\\b[^>]*\\br="${update.row}"[^>]*>[\\s\\S]*?<\\/row>`);
  const original = xml.match(re)?.[0];
  if (!original || !original.includes(update.name)) throw new Error(`Không tìm thấy đúng hàng ${update.name}`);
  let changed = original;
  for (const [col, value] of Object.entries(update.values)) changed = setCell(changed, update.row, col, value);
  xml = xml.replace(re, changed);
}
fs.writeFileSync(file, xml, 'utf8');
console.log(JSON.stringify(updates.map(({ row, name }) => ({ row, name })), null, 2));
