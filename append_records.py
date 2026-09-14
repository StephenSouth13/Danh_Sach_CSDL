import pandas as pd
from datetime import datetime
from openpyxl import load_workbook

file_path = r'D:\KLG\Danh_Sach_CSDL\DỮ LIỆU THÀNH TỰU GAB đối chiếu CSDL ĐÃ CHUẨN HÓA (1).xlsx'
xl = pd.ExcelFile(file_path)

df0 = pd.read_excel(xl, sheet_name=xl.sheet_names[0])

# Find starting index (after LÊ VIẾT HẢI)
idx = df0.index[df0['fullName'] == 'LÊ VIẾT HẢI'].tolist()
start_idx = max(idx) + 1 if idx else 0

df_next = df0.iloc[start_idx:]

stt = 200
current_name = 'LÊ VIẾT HẢI'

new_rows = []

for _, row in df_next.iterrows():
    name = str(row['fullName']).strip()
    if pd.isna(row['fullName']) or name == 'nan' or name == '':
        continue
        
    if name != current_name:
        stt += 1
        current_name = name
        
    title = str(row['title']).strip()
    time_val = row['time']
    time_str = ''
    
    if pd.notna(time_val):
        try:
            if isinstance(time_val, pd.Timestamp):
                time_str = time_val.strftime('%m/%Y')
            elif isinstance(time_val, str) and time_val.isnumeric():
                time_str = datetime.fromtimestamp(int(time_val)/1000).strftime('%m/%Y')
            elif isinstance(time_val, (int, float)) and time_val > 0:
                time_str = datetime.fromtimestamp(time_val/1000).strftime('%m/%Y')
        except Exception:
            pass
            
    url = str(row.get('url', ''))
    if pd.isna(row.get('url')) or url == 'nan':
        url = ''
        
    new_rows.append([stt, name, title, time_str, 'Đã có trên GAB', time_str, url])

# Append to Excel using openpyxl
wb = load_workbook(file_path)
sheet_name = xl.sheet_names[1]
ws = wb[sheet_name]

for r in new_rows:
    ws.append(r)

wb.save(file_path)
print(f"Successfully appended {len(new_rows)} rows. Final STT: {stt}")
