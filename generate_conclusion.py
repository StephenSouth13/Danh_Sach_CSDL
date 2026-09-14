import sys
sys.stdout.reconfigure(encoding='utf-8')
import pandas as pd
from openpyxl import load_workbook
import datetime

file_path = r'D:\KLG\Danh_Sach_CSDL\DỮ LIỆU THÀNH TỰU GAB đối chiếu CSDL ĐÃ CHUẨN HÓA (1).xlsx'

print("Đang đọc dữ liệu từ Excel...")
xl = pd.ExcelFile(file_path)
df0 = pd.read_excel(xl, sheet_name=xl.sheet_names[0])  # Dữ liệu gốc (GAB)
df1 = pd.read_excel(xl, sheet_name=xl.sheet_names[1])  # Dữ liệu đã chuẩn hóa (Trang tính 1)

# Lấy 50 người đầu tiên từ df1 để làm báo cáo đối chiếu
df_sample = df1.head(50).copy()

conclusions = []
for idx, row in df_sample.iterrows():
    name = row['Tên Kỷ lục gia']
    title_1 = str(row['Tiêu đề thành tựu']).strip()
    time_1 = str(row['Thời gian (CSDL)']).strip()
    
    # Tìm kiếm tương ứng bên df0
    match_df0 = df0[df0['fullName'] == name]
    
    if len(match_df0) > 0:
        # Lấy bản ghi đầu tiên khớp
        matched_row = match_df0.iloc[0]
        title_gab = str(matched_row['title']).strip()
        time_gab_ms = matched_row['time']
        
        # Định dạng thời gian GAB
        time_gab_str = "N/A"
        if pd.notna(time_gab_ms) and isinstance(time_gab_ms, (int, float)):
            try:
                time_gab_str = datetime.datetime.fromtimestamp(time_gab_ms/1000).strftime('%m/%Y')
            except:
                pass
                
        # Phân tích đánh giá chuyên gia
        assessment = []
        action = []
        
        # 1. So sánh thời gian
        if time_1 == time_gab_str:
            assessment.append("✅ Thời gian khớp")
        else:
            assessment.append("⚠️ Lệch định dạng thời gian")
            action.append("Cập nhật lại chuẩn MM/YYYY")
            
        # 2. So sánh nội dung
        if title_1 == title_gab:
            assessment.append("✅ Nội dung khớp hoàn toàn")
        else:
            assessment.append("ℹ️ Nội dung có sự sai khác về text")
            action.append("Rà soát câu từ")
            
        if not action:
            action.append("Đạt chuẩn, không cần chỉnh sửa")
            
        conclusions.append([
            row['STT'],
            name,
            title_gab,
            time_gab_str,
            title_1,
            time_1,
            " | ".join(assessment),
            ", ".join(action)
        ])
    else:
        conclusions.append([
            row['STT'],
            name,
            "Không tìm thấy trên GAB",
            "N/A",
            title_1,
            time_1,
            "❌ Thiếu dữ liệu gốc",
            "Thêm mới vào hệ thống"
        ])

# Tạo DataFrame cho Trang tính 3
df3 = pd.DataFrame(conclusions, columns=[
    "STT", 
    "Tên Kỷ lục gia", 
    "Tiêu đề (GAB/Kyluc gốc)", 
    "Thời gian (GAB/Kyluc gốc)", 
    "Tiêu đề (Đã rà soát/Nguồn ngoài)", 
    "Thời gian (Đã rà soát)", 
    "Kết luận đánh giá (Chuyên gia)", 
    "Hành động đề xuất"
])

print("Đang ghi vào Trang tính 3...")
try:
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df3.to_excel(writer, sheet_name='Trang tính 3', index=False)
    print("Thành công! Đã tạo Trang tính 3.")
except Exception as e:
    print(f"Lỗi ghi file: {e}")
    # Lưu ra file riêng nếu bị lỗi
    df3.to_excel('D:\KLG\Danh_Sach_CSDL\Trang_tinh_3_KetLuan.xlsx', index=False)
    print("Đã lưu vào Trang_tinh_3_KetLuan.xlsx")
