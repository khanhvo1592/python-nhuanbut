# check_license.py
import os
from tkinter import messagebox
from tkinter import Tk

def check_license_file(license_file_path):
    # Kiểm tra xem file có tồn tại không
    if os.path.exists(license_file_path):
        # Kiểm tra xem file có dữ liệu không
        if os.path.getsize(license_file_path) > 0:
            return True
        else:
            return False
    else:
        return False