import pandas as pd
import xlrd
import sys

# Internal methods reusage 
sys.path.append("C:/Users/andre/OneDrive/Documents/M.A/DA/IntroductionToDataAnalysis")
import Task1.task1

def all_data(files):
    if files is None:
        raise Exception("Input path is empty")
    
    print("\nConstruct dictionary from files:\n", files)
    column_names = {}
    data = []
    for file in files:
        book = xlrd.open_workbook(file)
        sheet = book.sheet_by_index(0)

        nrows = sheet.nrows
        ncols = sheet.ncols
        print("\nRanges: ", nrows, " ", ncols)
        # Indexes of columns
        for col in range(0, ncols):
            column_names[col] = sheet.cell(0, col).value

        for row in range(1, nrows - 1):

            row_data = []
            for col in range(0, ncols):
                row_data.append(sheet.cell(row, col).value)
            data.append(row_data)

    return column_names, data

def main():
    path_name = "C:/Users/andre/OneDrive/Documents/M.A/DA/IntroductionToDataAnalysis/Task2/Dannye_dlya_Praktiki_2_chast_2"
    
    all_files = Task1.task1.get_all_files(path_name)
    columns, data = all_data(all_files)
    print("\n\t-> Column names and inexes:\n", columns)
    print("\n\t-> All data:\n", pd.Series(data))

if __name__ == "__main__":
    main()