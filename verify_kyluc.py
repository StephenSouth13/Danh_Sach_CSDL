import sys
import os
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
from duckduckgo_search import DDGS
import time
import requests
from bs4 import BeautifulSoup
import re

file_path = r'D:\KLG\Danh_Sach_CSDL\DỮ LIỆU THÀNH TỰU GAB đối chiếu CSDL ĐÃ CHUẨN HÓA (1).xlsx'

print("Đang đọc dữ liệu từ file Excel...")
try:
    xl = pd.ExcelFile(file_path)
    df = pd.read_excel(xl, sheet_name='Các KLG hiện tại')
except Exception as e:
    print(f"Lỗi đọc file: {e}")
    sys.exit(1)

# Lấy dữ liệu bắt đầu từ dòng thực tế (dòng 1, bỏ qua header phụ ở dòng 0)
df_data = df.iloc[1:].copy()

results = []
print(f"Tổng số KLG cần kiểm tra: {len(df_data)}")

with DDGS() as ddgs:
    for index, row in df_data.iterrows():
        name = str(row['Unnamed: 6']).strip()
        if pd.isna(row['Unnamed: 6']) or name == 'nan' or name == '':
            continue
            
        current_title = str(row['Tiêu đề thành tựu trên GAB']).strip()
        current_link = str(row['Link bài viết trên GAB']).strip()
        
        print(f"\n--- Đang kiểm tra: {name} ---")
        
        link_status = "Không có link"
        web_title = ""
        suggested_links = []
        
        # 1. Kiểm tra link hiện tại nếu có
        if current_link.startswith('http'):
            try:
                res = requests.get(current_link, timeout=10, headers={'User-Agent': 'Mozilla/5.0'})
                if res.status_code == 200:
                    link_status = "Link hoạt động"
                    soup = BeautifulSoup(res.text, 'html.parser')
                    title_tag = soup.title
                    if title_tag:
                        web_title = title_tag.text.strip().replace(' - Hội kỷ lục gia Việt Nam', '')
                else:
                    link_status = f"Lỗi link ({res.status_code})"
            except Exception as e:
                link_status = f"Lỗi truy cập"
                
        # 2. Tìm kiếm thêm bằng DuckDuckGo
        query = f'"{name}" site:kyluc.vn'
        try:
            ddg_results = ddgs.text(query, max_results=3)
            for res in ddg_results:
                href = res.get('href', '')
                title = res.get('title', '')
                # Loại trừ link hiện tại
                if href != current_link and 'kyluc.vn' in href:
                    suggested_links.append(f"{title} ({href})")
            time.sleep(2)
        except Exception as e:
            print(f"Lỗi tìm kiếm DDG cho {name}: {e}")
            time.sleep(5)
            
        # 3. Đánh giá tiêu đề
        title_eval = "Đúng"
        if web_title and web_title not in current_title and current_title not in web_title:
             title_eval = "Sai khác so với bài viết gốc"
        elif not web_title:
             title_eval = "Không thể đối chiếu tiêu đề gốc"
             
        # Ghi kết quả
        results.append([
            name,
            current_link,
            link_status,
            current_title,
            web_title,
            title_eval,
            "\n".join(suggested_links) if suggested_links else "Không có thêm kỷ lục mới"
        ])
        print(f"Hoàn thành kiểm tra: {name}")

print("Đang lưu kết quả ra tab mới...")
df_out = pd.DataFrame(results, columns=[
    "Tên Kỷ lục gia",
    "Link hiện tại",
    "Trạng thái Link",
    "Tiêu đề đang có",
    "Tiêu đề thực tế trên Web",
    "Đánh giá Tiêu đề",
    "Các Kỷ lục/Bài viết đề xuất bổ sung"
])

try:
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df_out.to_excel(writer, sheet_name='Kết quả Kiểm tra Kyluc.vn', index=False)
    print("Xong! Đã lưu tab 'Kết quả Kiểm tra Kyluc.vn'")
except Exception as e:
    print(f"Lỗi ghi file Excel: {e}")
    df_out.to_excel(r'D:\KLG\Danh_Sach_CSDL\KetQuaKiemTra_Kyluc.xlsx', index=False)
    print("Đã lưu ra file riêng KetQuaKiemTra_Kyluc.xlsx")
