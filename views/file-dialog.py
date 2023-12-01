import tkinter as tk
from tkinter import filedialog

class FileDialogs:
    @staticmethod
    def open_file_dialog():
        root = tk.Tk()
        root.withdraw()
        return filedialog.askopenfilename(
            title="Select Excel file", 
            filetypes=(("Excel files", "*.xlsx;*.xls"), ("All files", "*.*"))
        )

    @staticmethod
    def save_file_dialog():
        root = tk.Tk()
        root.withdraw()
        return filedialog.asksaveasfilename(
            title="Save merged file as", 
            filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*")), 
            defaultextension=".xlsx"
        )
