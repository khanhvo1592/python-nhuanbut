from excel_merger import ExcelMerger
from file_dialogs import select_file
from create_pivot import create_pivot
from generate_individual_sheets import generate_individual_sheets
from check_license import check_license_file
import os
import shutil
from tkinter import messagebox
from tkinter import Tk
from datetime import datetime
import openpyxl
import tkinter as tk
from tkinter import filedialog
import logging

# Cấu hình logger
logging.basicConfig(filename='app.log', level=logging.INFO, format='%(asctime)s:%(levelname)s:%(message)s')


def create_and_save_excel_file():
    # Create a root window for tkinter and hide it
    root = tk.Tk()
    root.withdraw()

    # Open a dialog to choose a file path
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    if not file_path:
        print("No file selected.")
        return None

    # Create a new Excel file at the chosen path
    wb = openpyxl.Workbook()
    # Get the default sheet (which is typically named "Sheet1")
    default_sheet = wb.active
    # Rename the default sheet to "data"
    default_sheet.title = "data"
    # Save the workbook to the chosen file path
    wb.save(file_path)
    logging.info(f"File has been saved to: {file_path}")
    return file_path


def main():
    # Kiểm tra bản quyền
    license_path = r"D:\license\license.txt.txt"
    license_path2 = r"\\192.168.150.203\truyenthongso\license.txt.txt"
    if not (check_license_file(license_path) or check_license_file(license_path2)):
        # Nếu không có license, hiển thị dialog lỗi và thoát chương trình
        root = Tk()
        root.withdraw()  # Ẩn cửa sổ Tkinter chính
        messagebox.showerror("Lỗi", "Không có giấy phép hợp lệ!")
        logging.critical("License invalid")
        root.destroy()
        exit()

    # Run code
    merger = ExcelMerger()
    input_file = select_file()
    output_file = create_and_save_excel_file()  

    # Xử lí excel
    if input_file and output_file:
        merger.merge_sheets(input_file, output_file)
        create_pivot(output_file, output_file)
        generate_individual_sheets(output_file, output_file)
        logging.info("Excel processing completed")

    # Tạo tên file backup với timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_filename = f"NB_Backup_{timestamp}.xlsx"

    # Đường dẫn file gốc (đảm bảo rằng output_file là một đường dẫn hợp lệ)
    original_file_path = output_file

    # Tạo và lưu file backup mới tại cả hai địa điểm
    try:
        network_backup_path = f"\\\\192.168.150.203\\truyenthongso\\KHANH\\QLNB\\{backup_filename}"
        local_backup_path = f"D:\\license\\{backup_filename}"
        shutil.copy(original_file_path, network_backup_path)
        shutil.copy(original_file_path, local_backup_path)
        logging.info(f"File has been backed up to: {network_backup_path} and {local_backup_path}")
    except Exception as e:
        logging.error(f"An error occurred while copying the file: {e}")

    # Hiển thị dialog xác nhận mở file
    root = Tk()
    root.withdraw()  # Ẩn cửa sổ Tkinter chính
    if messagebox.askyesno("Mở File", "Bạn có muốn mở tập tin output không?"):
        # Mở file nếu người dùng chọn 'Yes'
        os.startfile(output_file)
    root.destroy()

if __name__ == "__main__":
    main()
