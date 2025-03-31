import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import json
import os
import openpyxl  # Ensure openpyxl is imported

# Hàm chuyển đổi dữ liệu sang JSON theo định dạng simo_001
def convert_to_simo_001(df):
    payload = []
    for _, row in df.iterrows():
        record = {
            "Cif": str(row.get("CIF", "")).strip().zfill(10) if pd.notna(row.get("CIF")) else None,
            "SoID": str(row.get("SoID", "")).strip()[:15] if pd.notna(row.get("SoID")) else None,
            "LoaiID": int(row.get("LoaiID", 0)) if pd.notna(row.get("LoaiID")) else None,
            "TenKhachHang": str(row.get("TenKhachHang", "")).strip()[:150] if pd.notna(row.get("TenKhachHang")) else None,
            "NgaySinh": str(row.get("NgaySinh", "")).strip()[:10] if pd.notna(row.get("NgaySinh")) else None,
            "GioiTinh": int(row.get("GioiTinh", 0)) if pd.notna(row.get("GioiTinh")) else None,
            "MaSoThue": str(row.get("MaSoThue", "")).strip() if pd.notna(row.get("MaSoThue")) else None,
            "SoDienThoaiDangKyDichVu": str(row.get("SoDienThoaiDangKyDichVu", "")).strip()[:15] if pd.notna(row.get("SoDienThoaiDangKyDichVu")) else None,
            "DiaChi": str(row.get("DiaChi", "")).strip()[:300] if pd.notna(row.get("DiaChi")) else None,
            "DiaChiKiemSoatTruyCap": str(row.get("DiaChiKiemSoatTruyCap", "")).strip()[:60] if pd.notna(row.get("DiaChiKiemSoatTruyCap")) else None,
            "MaSoNhanDangThietBiDiDong": str(row.get("MaSoNhanDangThietBiDiDong", "")).strip()[:36] if pd.notna(row.get("MaSoNhanDangThietBiDiDong")) else None,
            "SoTaiKhoan": str(row.get("SoTaiKhoan", "")).strip() if pd.notna(row.get("SoTaiKhoan")) else None,
            "LoaiTaiKhoan": int(row.get("LoaiTaiKhoan", 0)) if pd.notna(row.get("LoaiTaiKhoan")) else None,
            "TrangThaiHoatDongTaiKhoan": int(row.get("TrangThaiHoatDongTaiKhoan", 0)) if pd.notna(row.get("TrangThaiHoatDongTaiKhoan")) else None,
            "NgayMoTaiKhoan": str(row.get("NgayMoTaiKhoan", "")).strip()[:10] if pd.notna(row.get("NgayMoTaiKhoan")) else None,
            "PhuongThucMoTaiKhoan": int(row.get("PhuongThucMoTaiKhoan", 0)) if pd.notna(row.get("PhuongThucMoTaiKhoan")) else None,
            "NgayXacThucTaiQuay": str(row.get("NgayXacThucTaiQuay", "")).strip()[:10] if pd.notna(row.get("NgayXacThucTaiQuay")) else None,
            "QuocTich": str(row.get("QuocTich", "")).strip()[:36] if pd.notna(row.get("QuocTich")) else None
        }
        # Remove keys with None values
        record = {k: v for k, v in record.items() if v is not None}
        if record:  # Only append non-empty records
            payload.append(record)
    return payload

# Hàm chuyển đổi dữ liệu sang JSON theo định dạng simo_003
def convert_to_simo_003(df):
    payload = []
    for _, row in df.iterrows():
        record = {
            "Cif": str(row.get("CIF", "")).strip().zfill(10) if pd.notna(row.get("CIF")) else None,
            "SoTaiKhoan": str(row.get("SoTaiKhoan", "")).strip() if pd.notna(row.get("SoTaiKhoan")) else None,
            "TenKhachHang": str(row.get("TenKhachHang", "")).strip()[:150] if pd.notna(row.get("TenKhachHang")) else None,
            "TrangThaiHoatDongTaiKhoan": int(row.get("TrangThaiHoatDongTaiKhoan", 0)) if pd.notna(row.get("TrangThaiHoatDongTaiKhoan")) else None,
            "NghiNgo": int(row.get("NghiNgo", 0)) if pd.notna(row.get("NghiNgo")) else None,
            "GhiChu": str(row.get("GhiChu", "")).strip()[:500] if pd.notna(row.get("GhiChu")) else None
        }
        # Remove keys with None values
        record = {k: v for k, v in record.items() if v is not None}
        if record:  # Only append non-empty records
            payload.append(record)
    return payload

# Hàm chuyển đổi dữ liệu sang JSON theo định dạng simo_004
def convert_to_simo_004(df):
    payload = []
    for _, row in df.iterrows():
        record = {
            "Cif": str(row.get("CIF", "")).strip().zfill(10) if pd.notna(row.get("CIF")) else None,
            "SoID": str(row.get("SoID", "")).strip()[:15] if pd.notna(row.get("SoID")) else None,
            "LoaiID": int(row.get("LoaiID", 0)) if pd.notna(row.get("LoaiID")) else None,
            "TenKhachHang": str(row.get("TenKhachHang", "")).strip()[:150] if pd.notna(row.get("TenKhachHang")) else None,
            "NgaySinh": str(row.get("NgaySinh", "")).strip()[:10] if pd.notna(row.get("NgaySinh")) else None,
            "GioiTinh": int(row.get("GioiTinh", 0)) if pd.notna(row.get("GioiTinh")) else None,
            "MaSoThue": str(row.get("MaSoThue", "")).strip() if pd.notna(row.get("MaSoThue")) else None,
            "SoDienThoaiDangKyDichVu": str(row.get("SoDienThoaiDangKyDichVu", "")).strip()[:15] if pd.notna(row.get("SoDienThoaiDangKyDichVu")) else None,
            "DiaChi": str(row.get("DiaChi", "")).strip()[:300] if pd.notna(row.get("DiaChi")) else None,
            "DiaChiKiemSoatTruyCap": str(row.get("DiaChiKiemSoatTruyCap", "")).strip()[:60] if pd.notna(row.get("DiaChiKiemSoatTruyCap")) else None,
            "MaSoNhanDangThietBiDiDong": str(row.get("MaSoNhanDangThietBiDiDong", "")).strip()[:36] if pd.notna(row.get("MaSoNhanDangThietBiDiDong")) else None,
            "SoTaiKhoan": str(row.get("SoTaiKhoan", "")).strip() if pd.notna(row.get("SoTaiKhoan")) else None,
            "LoaiTaiKhoan": int(row.get("LoaiTaiKhoan", 0)) if pd.notna(row.get("LoaiTaiKhoan")) else None,
            "TrangThaiHoatDongTaiKhoan": int(row.get("TrangThaiHoatDongTaiKhoan", 0)) if pd.notna(row.get("TrangThaiHoatDongTaiKhoan")) else None,
            "NgayMoTaiKhoan": str(row.get("NgayMoTaiKhoan", "")).strip()[:10] if pd.notna(row.get("NgayMoTaiKhoan")) else None,
            "PhuongThucMoTaiKhoan": int(row.get("PhuongThucMoTaiKhoan", 0)) if pd.notna(row.get("PhuongThucMoTaiKhoan")) else None,
            "NgayXacThucTaiQuay": str(row.get("NgayXacThucTaiQuay", "")).strip()[:10] if pd.notna(row.get("NgayXacThucTaiQuay")) else None,
            "GhiChu": str(row.get("GhiChu", "")).strip()[:500] if pd.notna(row.get("GhiChu")) else None,
            "QuocTich": str(row.get("QuocTich", "")).strip()[:36] if pd.notna(row.get("QuocTich")) else None
        }
        # Remove keys with None values
        record = {k: v for k, v in record.items() if v is not None}
        if record:  # Only append non-empty records
            payload.append(record)
    return payload

# Hàm lưu file JSON vào vị trí người dùng chọn
def save_json_to_location(json_data, service_type):
    try:
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            initialfile=f"payload_{service_type}.json"
        )
        if file_path:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(json_data)
            messagebox.showinfo("Thành công", f"Đã lưu file JSON tại: {file_path}")
    except Exception as e:
        messagebox.showerror("Lỗi", f"Không thể lưu file: {str(e)}")

# Hàm copy JSON vào clipboard
def copy_to_clipboard(json_data):
    root.clipboard_clear()
    root.clipboard_append(json_data)
    root.update()  # Đảm bảo clipboard được cập nhật
    messagebox.showinfo("Thành công", "Đã sao chép JSON vào clipboard!")

# Hàm xử lý file Excel và tạo JSON
def process_excel(file_path, service_type):
    try:
        # Đọc file Excel bằng openpyxl và đảm bảo cột CIF được đọc dưới dạng chuỗi
        df = pd.read_excel(file_path, dtype=str, engine='openpyxl')
        
        # Hiển thị dữ liệu trong bảng trước khi xuất JSON
        display_data(df)

        # Chuyển đổi theo loại dịch vụ
        if service_type == "simo_001":
            payload = convert_to_simo_001(df)
        elif service_type == "simo_003":
            payload = convert_to_simo_003(df)
        elif service_type == "simo_004":
            payload = convert_to_simo_004(df)
        else:
            raise ValueError("Service type không hợp lệ!")

        # Chuyển thành JSON
        json_data = json.dumps(payload, ensure_ascii=False, indent=2)
        
        # Lưu file JSON vào vị trí người dùng chọn
        save_json_to_location(json_data, service_type)
        
        # Hiển thị trên giao diện
        text_area.delete(1.0, tk.END)
        text_area.insert(tk.END, json_data)
        messagebox.showinfo("Thành công", f"Đã tạo file payload_{service_type}.json!")
        
    except Exception as e:
        messagebox.showerror("Lỗi", f"Đã xảy ra lỗi: {str(e)}")

# Hàm hiển thị dữ liệu trong bảng
def display_data(df):
    for widget in data_frame.winfo_children():
        widget.destroy()  # Clear previous data

    # Create a Treeview with a horizontal scrollbar
    tree_frame = tk.Frame(data_frame)
    tree_frame.pack(fill=tk.BOTH, expand=True)

    tree_scroll_x = tk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
    tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

    tree = ttk.Treeview(tree_frame, columns=list(df.columns), show="headings", height=20, xscrollcommand=tree_scroll_x.set)
    tree.pack(fill=tk.BOTH, expand=True)

    tree_scroll_x.config(command=tree.xview)

    # Set column headers
    for col in df.columns:
        tree.heading(col, text=col)
        tree.column(col, width=150, anchor="center")

    # Insert rows into the treeview
    for _, row in df.iterrows():
        tree.insert("", tk.END, values=list(row))

# Hàm chọn file Excel
def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, file_path)

# Hàm chạy chương trình
def run_conversion():
    file_path = entry_file.get()
    service_type = service_var.get()
    
    if not file_path or not os.path.exists(file_path):
        messagebox.showwarning("Cảnh báo", "Vui lòng chọn file Excel hợp lệ!")
        return
    
    if not service_type:
        messagebox.showwarning("Cảnh báo", "Vui lòng chọn loại dịch vụ!")
        return
    
    process_excel(file_path, service_type)

# Tạo giao diện GUI
root = tk.Tk()
root.title("SIMO JSON Converter")
root.geometry("1900x900")  # Update window size

# Nhãn chọn file
label_file = tk.Label(root, text="Chọn file Excel:")
label_file.pack(pady=5)

# Ô nhập file
entry_file = tk.Entry(root, width=100)
entry_file.pack(pady=5)

# Nút chọn file
button_browse = tk.Button(root, text="Browse", command=select_file)
button_browse.pack(pady=5)

# Nhãn chọn dịch vụ
label_service = tk.Label(root, text="Chọn loại dịch vụ:")
label_service.pack(pady=5)

# Dropdown chọn dịch vụ
service_var = tk.StringVar(value="simo_001")
service_options = ["simo_001", "simo_003", "simo_004"]
dropdown_service = tk.OptionMenu(root, service_var, *service_options)
dropdown_service.pack(pady=5)

# Nút chạy
button_run = tk.Button(root, text="Chuyển đổi", command=run_conversion)
button_run.pack(pady=10)

# Frame hiển thị dữ liệu
data_frame = tk.Frame(root)
data_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

# Nút lưu JSON
button_save = tk.Button(root, text="Lưu JSON", command=lambda: save_json_to_location(text_area.get(1.0, tk.END), service_var.get()))
button_save.pack(pady=5)

# Nút sao chép JSON
button_copy = tk.Button(root, text="Sao chép JSON", command=lambda: copy_to_clipboard(text_area.get(1.0, tk.END)))
button_copy.pack(pady=5)

# Text area hiển thị JSON
text_area = tk.Text(root, height=15, width=150)
text_area.pack(pady=10)

# Chạy chương trình
root.mainloop()
