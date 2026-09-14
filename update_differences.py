import sys
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
import difflib

file_path = r'D:\KLG\Danh_Sach_CSDL\DỮ LIỆU THÀNH TỰU GAB đối chiếu CSDL ĐÃ CHUẨN HÓA (1).xlsx'

print("Đang đọc dữ liệu từ tab 'Kết quả Kiểm tra Kyluc.vn'...")
try:
    xl = pd.ExcelFile(file_path)
    df = pd.read_excel(xl, sheet_name='Kết quả Kiểm tra Kyluc.vn')
except Exception as e:
    print(f"Lỗi đọc file (có thể file đang mở): {e}")
    sys.exit(1)

def get_diff_details(text1, text2):
    if pd.isna(text1): text1 = ""
    if pd.isna(text2): text2 = ""
    
    text1 = str(text1).strip()
    text2 = str(text2).strip()
    
    # Chia thành các từ để so sánh
    words1 = text1.split()
    words2 = text2.split()
    
    diff = difflib.ndiff(words1, words2)
    missing_in_web = []  # Có trong file excel nhưng web không có (thừa)
    added_in_web = []    # Không có trong file excel nhưng web có (thiếu)
    
    for term in diff:
        if term.startswith('- '):
            missing_in_web.append(term[2:])
        elif term.startswith('+ '):
            added_in_web.append(term[2:])
            
    details = []
    if missing_in_web:
        details.append(f"Dư các từ: '{' '.join(missing_in_web)}'")
    if added_in_web:
        details.append(f"Thiếu các từ: '{' '.join(added_in_web)}'")
        
    if not details:
        return "Gần như trùng khớp (chỉ khác khoảng trắng/in hoa)"
        
    return "\n".join(details)

print("Đang phân tích lỗi sai chi tiết...")
details_list = []
for index, row in df.iterrows():
    if row.get('Đánh giá Tiêu đề') == "Sai khác so với bài viết gốc":
        t_excel = row.get('Tiêu đề đang có')
        t_web = row.get('Tiêu đề thực tế trên Web')
        detail = get_diff_details(t_excel, t_web)
        details_list.append(detail)
    else:
        details_list.append("Không có lỗi")

# Thêm cột mới vào sau cột Đánh giá Tiêu đề
# Tìm vị trí cột
try:
    idx = df.columns.get_loc('Đánh giá Tiêu đề') + 1
    df.insert(idx, 'Chi tiết lỗi sai', details_list)
except ValueError:
    df['Chi tiết lỗi sai'] = details_list

print("Đang ghi lại dữ liệu vào file Excel...")
try:
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='Kết quả Kiểm tra Kyluc.vn', index=False)
    print("Xong! Đã cập nhật thành công cột 'Chi tiết lỗi sai'.")
except Exception as e:
    print(f"Lỗi ghi file (Hãy tắt file Excel trước khi chạy): {e}")
