import sys
import os
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
from openpyxl import load_workbook
from duckduckgo_search import DDGS
import time
import re
import datetime

file_path = r'D:\KLG\Danh_Sach_CSDL\DỮ LIỆU THÀNH TỰU GAB đối chiếu CSDL ĐÃ CHUẨN HÓA (1).xlsx'

print("Đọc dữ liệu từ file Excel...")
try:
    xl = pd.ExcelFile(file_path)
    df0 = pd.read_excel(xl, sheet_name=xl.sheet_names[0]) # GAB data
    df1 = pd.read_excel(xl, sheet_name=xl.sheet_names[1]) # Master list
    
    # Check if Trang tính 2 exists
    if 'Trang tính 2' in xl.sheet_names:
        df2 = pd.read_excel(xl, sheet_name='Trang tính 2')
        processed_names = df2['Tên Kỷ lục gia'].unique().tolist()
    else:
        df2 = pd.DataFrame(columns=["STT", "Tên Kỷ lục gia", "Tiêu đề thành tựu", "Thời gian (CSDL)", "Trạng thái trên GAB", "Thời gian trên GAB", "Nguồn research"])
        processed_names = []
        
except Exception as e:
    print(f"Lỗi đọc file: {e}")
    sys.exit(1)

all_names = df1[df1['STT'] >= 201]['Tên Kỷ lục gia'].unique().tolist()
names_to_search = [name for name in all_names if name not in processed_names][:20] # Lấy 20 người tiếp theo

if not names_to_search:
    print("Đã hoàn thành toàn bộ danh sách!")
    sys.exit(0)

print(f"Bắt đầu cào dữ liệu cho {len(names_to_search)} người (Batch mới)...")

results_data = []
stt = 200 if df2.empty else df2['STT'].max()

def extract_date(text):
    match = re.search(r'\b(?:0?[1-9]|1[0-2])/(?:19|20)\d{2}\b|\b(?:19|20)\d{2}\b', text)
    if match:
        return match.group(0)
    return 'N/A'

with DDGS() as ddgs:
    for i, name in enumerate(names_to_search):
        stt += 1
        query = f'"{name}" "kỷ lục" -site:kyluc.vn -site:gab.world'
        print(f"[{i+1}/{len(names_to_search)}] Đang tìm kiếm: {name}...")
        
        found_any = False
        try:
            results = ddgs.text(query, max_results=2)
            for res in results:
                title = res.get('title', '')
                desc = res.get('body', '')
                url = res.get('href', '')
                
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
                
            time.sleep(5) # Delay an toàn 5 giây
            
        except Exception as e:
            print(f"Lỗi khi tìm kiếm {name}: {e}")
            time.sleep(10)

        if not found_any:
            results_data.append([stt, name, "Không tìm thấy thông tin trên nguồn ngoài", "N/A", "Nguồn ngoài", "N/A", ""])

# Nối dữ liệu mới vào Trang tính 2
df_new = pd.DataFrame(results_data, columns=df2.columns)
df2_updated = pd.concat([df2, df_new], ignore_index=True)

# Tự động cập nhật Trang tính 3 (Bảng kết luận đối chiếu) dựa trên df2_updated
print("Đang tạo Trang tính 3 đối chiếu...")
conclusions = []
for idx, row in df2_updated.iterrows():
    name = row['Tên Kỷ lục gia']
    title_ngoai = str(row['Tiêu đề thành tựu']).strip()
    time_ngoai = str(row['Thời gian (CSDL)']).strip()
    
    match_df0 = df0[df0['fullName'] == name]
    if len(match_df0) > 0:
        matched_row = match_df0.iloc[0]
        title_gab = str(matched_row['title']).strip()
        time_gab_ms = matched_row['time']
        
        time_gab_str = "N/A"
        if pd.notna(time_gab_ms) and isinstance(time_gab_ms, (int, float)):
            try:
                time_gab_str = datetime.datetime.fromtimestamp(time_gab_ms/1000).strftime('%m/%Y')
            except:
                pass
                
        assessment = []
        action = []
        if title_ngoai == "Không tìm thấy thông tin trên nguồn ngoài":
            assessment.append("❌ Báo chí không đưa tin")
            action.append("Cần xác minh truyền thông")
        else:
            if time_ngoai != 'N/A' and time_ngoai in time_gab_str:
                assessment.append("✅ Thời gian khớp")
            elif time_ngoai == 'N/A':
                assessment.append("ℹ️ Nguồn ngoài không ghi rõ thời gian")
            else:
                assessment.append("⚠️ Lệch thời gian")
                action.append("Đối chiếu lại hồ sơ")
                
            assessment.append("✅ Có thông tin trên báo")
            
        if not action:
            action.append("Thông tin an toàn")
            
        conclusions.append([
            row['STT'], name, title_gab, time_gab_str, title_ngoai, time_ngoai, " | ".join(assessment), ", ".join(action)
        ])

df3 = pd.DataFrame(conclusions, columns=[
    "STT", "Tên Kỷ lục gia", "Tiêu đề (GAB/Kyluc gốc)", "Thời gian (GAB/Kyluc gốc)", 
    "Tiêu đề (Nguồn ngoài)", "Thời gian (Nguồn ngoài)", "Kết luận đánh giá (Chuyên gia)", "Hành động đề xuất"
])

print("Đang ghi vào file Excel...")
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df2_updated.to_excel(writer, sheet_name='Trang tính 2', index=False)
    df3.to_excel(writer, sheet_name='Trang tính 3', index=False)

print(f"XONG BATCH! Đã cập nhật xong Trang tính 2 và 3 cho {len(names_to_search)} người mới.")
