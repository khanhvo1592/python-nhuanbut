import tkinter as tk
from tkinter import filedialog
import os
from plugins.merge_sheets_plugin import MergeSheetsPlugin

class Application:
    def __init__(self):
        self.plugins = {"merge_sheets": MergeSheetsPlugin()}

    def run(self):
        root = tk.Tk()
        root.withdraw()

        input_file = filedialog.askopenfilename(
            title="Select Excel file", 
            filetypes=(("Excel files", "*.xlsx;*.xls"), ("All files", "*.*"))
        )

        output_file = filedialog.asksaveasfilename(
            title="Save merged file as", 
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")), 
            defaultextension=".xlsx"
        )

        if input_file and output_file:
            self.plugins["merge_sheets"].process(input_file, output_file)
            self.open_file(output_file)

    def open_file(self, file_path):
        # Mở file bằng chương trình mặc định của hệ thống
        os.startfile(file_path)  # Chú ý: os.