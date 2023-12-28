import tkinter as tk
from tkinter import filedialog

def select_file():
    root = tk.Tk()
    root.withdraw()  # Ẩn cửa sổ tkinter
    file_path = filedialog.askopenfilename()
    root.destroy()  # Đóng cửa sổ tkinter sau khi chọn file
    return file_path

# def select_file_output():
#     root = tk.Tk()
#     root.withdraw()  # Ẩn cửa sổ tkinter
#     file_path = filedialog.askopenfilename()
#     root.destroy()  # Đóng cửa sổ tkinter sau khi chọn file
#     return file_path
