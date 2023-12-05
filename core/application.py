import tkinter as tk
from tkinter import filedialog
import os
from plugins.merge_sheets_plugin import MergeSheetsPlugin

class Application:
    def __init__(self):
        self.plugins = {"merge_sheets": MergeSheetsPlugin(start_row_input=13)}

    def run(self):
        root = tk.Tk()
        root.withdraw()

        # Người dùng chọn file input
        input_file = filedialog.askopenfilename(
            title="Select Excel file", 
            filetypes=(("Excel files", "*.xlsx;*.xls"), ("All files", "*.*"))
        )

        # Người dùng chọn file output đã tồn tại
        output_file = filedialog.askopenfilename(
            title="Select existing output Excel file", 
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*"))
        )

        # Xử lý file nếu cả file input và output đã được chọn
        if input_file and output_file:
            self.plugins["merge_sheets"].process(input_file, output_file)
            self.open_file(output_file)

    def open_file(self, file_path):
        # Mở file bằng chương trình mặc định của hệ thống
        os.startfile(file_path)


