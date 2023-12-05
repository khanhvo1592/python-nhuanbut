import tkinter as tk
from tkinter import filedialog
import openpyxl

class ExcelMerger:
    def __init__(self, start_row=12):
        self.start_row = start_row

    def merge_sheets(self, input_file, output_file):
        source_wb = openpyxl.load_workbook(input_file, data_only=True)
        target_wb = openpyxl.Workbook()
        target_ws = target_wb.active

        for sheet in source_wb.sheetnames:
            ws = source_wb[sheet]
            b8_value = ws['B8'].value  # Lấy giá trị từ ô B8
            b9_value = ws['B9'].value  # Lấy giá trị từ ô B9
            for row in ws.iter_rows(min_row=self.start_row, values_only=True):
                if row[2]:  # Kiểm tra dữ liệu ở cột C
                    new_row = [b8_value, b9_value] + list(row)
                    target_ws.append(new_row)

        target_wb.save(output_file)


def select_file():
    root = tk.Tk()
    root.withdraw()  # Ẩn cửa sổ tkinter
    file_path = filedialog.askopenfilename()
    return file_path

def save_file_dialog():
    root = tk.Tk()
    root.withdraw()  # Ẩn cửa sổ tkinter
    file_path = filedialog.asksaveasfilename(
        defaultextension='.xlsx',
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        title="Save merged file as"
    )
    return file_path

# Sử dụng
merger = ExcelMerger()
input_file = select_file()
output_file = save_file_dialog()

if input_file and output_file:
    merger.merge_sheets(input_file, output_file)
