import pandas as pd
import pyodbc

# ==================== Bước 1: Đọc file Excel ====================
file_path = r"E:\Scrip\APISIMO\Upload_TKTTT_CaNhan_FULL_TestData.xlsx"

# Chỉ định kiểu dữ liệu cho các cột quan trọng
dtype_spec = {
    "SoCIF": str,
    "SoID": str,
    "MaSoThue": str,
    "SoDienThoaiDangKyDichVu": str,
    "DiaChi": str
}

df = pd.read_excel(file_path, header=0, dtype=dtype_spec)  # Header là dòng đầu tiên

# Loại bỏ cột đầu tiên (STT)
df = df.iloc[:, 1:]

# Gán tên cột dựa trên thứ tự bạn cung cấp
column_names = [
    "SoCIF", "SoID", "LoaiID", "TenKhachHang", "NgaySinh", "GioiTinh",
    "QuocTich", "MaSoThue", "SoDienThoaiDangKyDichVu", "DiaChi",
    "DiaChiKiemSoatTruyCap", "MaSoNhanDangThietBiDiDong", "SoTaiKhoan",
    "LoaiTaiKhoan", "TrangThaiHoatDongTaiKhoan", "NgayMoTaiKhoan",
    "PhuongThucMoTaiKhoan", "NgayXacThucTaiQuay"
]
df.columns = column_names

# ==================== Bước 2: Xử lý dữ liệu ====================
# Chuyển đổi ngày tháng với định dạng rõ ràng
date_columns = ["NgaySinh", "NgayMoTaiKhoan", "NgayXacThucTaiQuay"]
for col in date_columns:
    df[col] = pd.to_datetime(df[col], format="%d/%m/%Y", errors="coerce").dt.strftime("%d/%m/%Y")
    df[col] = df[col].fillna("01/01/1900")  # Thay thế NaT bằng giá trị mặc định

# Chuyển đổi giới tính
df["GioiTinh"] = df["GioiTinh"].map({1: 1, 0: 0, 2: 2}).fillna(2)

# Làm sạch các cột kiểu int
int_columns = [
    "LoaiID", "GioiTinh", "TrangThaiHoatDongTaiKhoan",
    "PhuongThucMoTaiKhoan", "LoaiTaiKhoan"
]
for col in int_columns:
    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

# Xử lý các cột kiểu chuỗi
string_columns = [
    "SoCIF", "SoID", "MaSoThue", "SoDienThoaiDangKyDichVu", "DiaChi"
]

for col in string_columns:
    df[col] = df[col].astype(str).str.strip()  # Loại bỏ khoảng trắng thừa
    df[col] = df[col].replace("", None)       # Thay thế giá trị rỗng bằng NULL

# Kiểm tra và điền giá trị "0" cho cột MaSoThue và SoDienThoaiDangKyDichVu nếu rỗng
for col in ["MaSoThue", "SoDienThoaiDangKyDichVu"]:
    if df[col].isnull().any() or (df[col] == "").any():
        print(f"Cột '{col}' có giá trị rỗng. Đã tự động điền giá trị '0' cho các ô rỗng.")
        df[col] = df[col].fillna("0").replace("", "0")

# Giới hạn độ dài dữ liệu
df["SoCIF"] = df["SoCIF"].str[:36]
df["SoID"] = df["SoID"].str[:15]
df["MaSoThue"] = df["MaSoThue"].str[:13]
df["SoDienThoaiDangKyDichVu"] = df["SoDienThoaiDangKyDichVu"].str[:15]
df["DiaChi"] = df["DiaChi"].str[:300]

# Xử lý giá trị null/NaN
df = df.where(pd.notnull(df), None)

# ==================== Bước 3: Kết nối SQL Server ====================
server = "10.8.103.21"
database = "test"
username = "sa"
password = "q"

conn_str = f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}"
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# ==================== Bước 4: Insert dữ liệu hàng loạt ====================
columns = [
    "SoCIF", "SoID", "LoaiID", "TenKhachHang", "NgaySinh", "GioiTinh",
    "QuocTich", "MaSoThue", "SoDienThoaiDangKyDichVu", "DiaChi",
    "DiaChiKiemSoatTruyCap", "MaSoNhanDangThietBiDiDong", "SoTaiKhoan",
    "LoaiTaiKhoan", "TrangThaiHoatDongTaiKhoan", "NgayMoTaiKhoan",
    "PhuongThucMoTaiKhoan", "NgayXacThucTaiQuay"
]

insert_query = f"""
INSERT INTO DanhSachTKTT ({', '.join(columns)})
VALUES ({', '.join(['?'] * len(columns))})
"""

data = [tuple(row[col] for col in columns) for _, row in df.iterrows()]

try:
    cursor.executemany(insert_query, data)
    conn.commit()
    print(f"Đã insert {len(data)} bản ghi thành công!")
except Exception as e:
    print("Lỗi:", str(e))
    conn.rollback()
finally:
    cursor.close()
    conn.close()