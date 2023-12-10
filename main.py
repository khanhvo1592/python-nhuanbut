from excel_merger import ExcelMerger
from file_dialogs import select_file, select_file_output
from create_pivot import create_pivot
from init import clean_sheets
from generate_individual_sheets import generate_individual_sheets
from check_license import check_license_file
import os
from tkinter import messagebox
from tkinter import Tk

def main():
    # Kiểm tra bản quyền
    license_path = r"D:\license\license.txt.txt"
    license_path2 = r"\\192.168.150.203\truyenthongso\license.txt.txt"
    if not (check_license_file(license_path) or check_license_file(license_path2)):
        # Nếu không có license, hiển thị dialog lỗi và thoát chương trình
        root = Tk()
        root.withdraw()  # Ẩn cửa sổ Tkinter chính
        messagebox.showerror("Lỗi", "Không có giấy phép hợp lệ!")
        root.destroy()
        exit()

    # Run code
    merger = ExcelMerger()
    input_file = select_file()
    output_file = select_file_output()

    # Xử lý trước dữ liệu output
    clean_sheets(output_file)

    # Xử lí excel
    if input_file and output_file:
        merger.merge_sheets(input_file, output_file)
        create_pivot(output_file, output_file)
        generate_individual_sheets(output_file, output_file)

    # Hiển thị dialog xác nhận mở file
    root = Tk()
    root.withdraw()  # Ẩn cửa sổ Tkinter chính
    if messagebox.askyesno("Mở File", "Bạn có muốn mở tập tin output không?"):
        # Mở file nếu người dùng chọn 'Yes'
        os.startfile(output_file)
    root.destroy()

if __name__ == "__main__":
    main()
