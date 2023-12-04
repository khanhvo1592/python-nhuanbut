import openpyxl

class MergeSheetsPlugin:
    def __init__(self, start_row=13):
        self.start_row = start_row

    def process(self, input_file, output_file):
        wb = openpyxl.load_workbook(input_file)
        new_wb = openpyxl.Workbook()
        new_ws = new_wb.active

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for row in ws.iter_rows(min_row=self.start_row, values_only=True):
                if row[2]:  # Check if the value in column C (3rd column) is blank
                    new_ws.append(row)

        new_wb.save(output_file)
