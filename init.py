import openpyxl
import tkinter as tk
from tkinter import filedialog

def create_and_save_excel_file():
    """
    Creates a new Excel file and saves it to the chosen file path.

    Returns:
        str: The file path where the Excel file is saved.
              Returns None if no file is selected.
    """
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

    print(f"File saved to: {file_path}")
    return file_path
