
import openpyxl

def clean_sheets(output_file):
    # Mở workbook hiện có
    book = openpyxl.load_workbook(output_file)
    
    # Xóa tất cả các sheet ngoại trừ sheet "data"
    for sheet_name in book.sheetnames:
        if sheet_name != "data":
            std = book[sheet_name]
            book.remove(std)
            
    # Kiểm tra nếu sheet "data" không tồn tại thì tạo mới
    if "data" not in book.sheetnames:
        book.create_sheet("data")
    else:
        # Nếu sheet "data" tồn tại, xóa sạch nội dung của nó
        sheet = book["data"]
        for row in sheet['A1:Z' + str(sheet.max_row)]:
            for cell in row:
                cell.value = None

    # Lưu workbook
    book.save(output_file)
