from excel_merger import ExcelMerger
from file_dialogs import select_file, select_file_output

def main():
    merger = ExcelMerger()
    input_file = select_file()
    output_file = select_file_output()

    if input_file and output_file:
        merger.merge_sheets(input_file, output_file)

if __name__ == "__main__":
    main()
