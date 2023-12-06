from excel_merger import ExcelMerger
from file_dialogs import select_file, select_file_output
from create_pivot import create_pivot
from init import clean_sheets
from generate_individual_sheets import generate_individual_sheets


def main():
    merger = ExcelMerger()
    input_file = select_file()
    output_file = select_file_output()
    #xử lý trước dữ liệu output 
    clean_sheets(output_file)
    if input_file and output_file:
        merger.merge_sheets(input_file, output_file)
        create_pivot(output_file, output_file)
        generate_individual_sheets(output_file, output_file)

if __name__ == "__main__":
    main()
