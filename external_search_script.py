import sys
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
from openpyxl import load_workbook
from googlesearch import search
import time
import re

file_path = r'D:\KLG\Danh_Sach_CSDL\DỮ LIỆU THÀNH TỰU GAB đối chiếu CSDL ĐÃ CHUẨN HÓA (1).xlsx'

print("Đọc dữ liệu từ file Excel...")
try:
    xl = pd.ExcelFile(file_path)
    df1 = pd.read_excel(xl, sheet_name=xl.sheet_names[1])
except Exception as e:
    print(f"Lỗi đọc file: {e}")
    sys.exit(1)

# Lấy các tên từ STT >= 201
names_to_search = df1[df1['STT'] >= 201]['Tên Kỷ lục gia'].unique().tolist()
print(f"Tổng số người cần tìm kiếm: {len(names_to_search)}")

results_data = []

stt = 200

# Hàm trích xuất năm/tháng từ văn bản
def extract_date(text):
    # Match dd/mm/yyyy, mm/yyyy, or yyyy
    match = re.search(r'\b(?:0?[1-9]|1[0-2])/(?:19|20)\d{2}\b|\b(?:19|20)\d{2}\b', text)
    if match:
        return match.group(0)
    return 'N/A'

for i, name in enumerate(names_to_search):
    stt += 1
    query = f'"{name}" "kỷ lục" -site:kyluc.vn -site:gab.world'
    print(f"[{i+1}/{len(names_to_search)}] Đang tìm kiếm: {name}...")
    
    found_any = False
    try:
        # Lấy 2 kết quả đầu tiên
        search_res = search(query, num_results=2, sleep_interval=2, advanced=True)
        for res in search_res:
            title = res.title
            desc = res.description
            url = res.url
            
            # extract date from description or title
            date_str = extract_date(desc)
            if date_str == 'N/A':
                date_str = extract_date(title)
                
            results_data.append([
                stt, 
                name, 
                title[:200] + ('...' if len(title) > 200 else ''), 
                date_str, 
                'Nguồn ngoài', 
                'N/A', 
                url
            ])
            found_any = True
            
        time.sleep(4)
        
    except Exception as e:
        print(f"Lỗi khi tìm kiếm {name}: {e}")
        time.sleep(15) # Ngừng lâu hơn nếu bị block

    if not found_any:
        results_data.append([stt, name, "Không tìm thấy thông tin trên nguồn ngoài", "N/A", "Nguồn ngoài", "N/A", ""])
        
    # Thỉnh thoảng save log lại
    if (i + 1) % 10 == 0:
        print(f"Đã xử lý {i+1} người...")

# Tạo DataFrame và ghi vào Trang tính 2
print("Đang tạo Trang tính 2...")
df_out = pd.DataFrame(results_data, columns=["STT", "Tên Kỷ lục gia", "Tiêu đề thành tựu", "Thời gian (CSDL)", "Trạng thái trên GAB", "Thời gian trên GAB", "Nguồn research"])

try:
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_out.to_excel(writer, sheet_name='Trang tính 2', index=False)
    print("Đã hoàn thành lưu Trang tính 2!")
except Exception as e:
    print(f"Lỗi khi ghi file, thử lưu file mới: {e}")
    df_out.to_excel('D:\KLG\Danh_Sach_CSDL\Trang_tinh_2_Tam.xlsx', index=False)
    print("Đã lưu vào Trang_tinh_2_Tam.xlsx")
