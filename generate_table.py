import pandas as pd
from datetime import datetime

xl = pd.ExcelFile(r'D:\KLG\Danh_Sach_CSDL\DỮ LIỆU THÀNH TỰU GAB đối chiếu CSDL ĐÃ CHUẨN HÓA (1).xlsx')
df0 = pd.read_excel(xl, sheet_name=xl.sheet_names[0])
idx = df0.index[df0['fullName'] == 'LÊ VIẾT HẢI'].tolist()
start_idx = max(idx) + 1
df_next = df0.iloc[start_idx:start_idx+19]
lines = []
stt = 201
current_name = str(df_next.iloc[0]['fullName']).strip()

for _, row in df_next.iterrows():
    name = str(row['fullName']).strip()
    if name != current_name:
        stt += 1
        current_name = name
    title = str(row['title']).strip()
    time_val = row['time']
    time_str = 'N/A'
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
    lines.append(f'| {stt} | {name} | {title} | {time_str} | Cần rà soát GAB | {time_str} |')

with open('md_table.txt', 'w', encoding='utf-8') as f:
    f.write('\n'.join(lines))
