import pandas as pd
import matplotlib.pyplot as plt
import xlrd
import sys

from collections import Counter

# Internal methods reusage 
sys.path.append("C:/Users/andre/OneDrive/Documents/M.A/DA/IntroductionToDataAnalysis")
import Task1.task1

def all_data(files):
    if files is None:
        raise Exception("Input path is empty")
    
    print("\nConstruct dictionary from files:\n", files)
    column_names = {}
    data = []
    severity_list = []
    visibility_list = []
    for file in files:
        book = xlrd.open_workbook(file)
        sheet = book.sheet_by_index(0)

        nrows = sheet.nrows
        ncols = sheet.ncols
        print("\nRanges: ", nrows, " ", ncols)
        # Indexes of columns
        for col in range(0, ncols):
            column_names[sheet.cell(0, col).value] = col

        for row in range(1, nrows - 1):
            severity_list.append(int(sheet.cell(row, int(column_names.get("Severity"))).value))
            if sheet.cell(row, int(column_names.get("Visibility(mi)"))).value != "":
                visibility_list.append(float(sheet.cell(row, int(column_names.get("Visibility(mi)"))).value))
            row_data = []
            for col in range(0, ncols):
                row_data.append(sheet.cell(row, col).value)
            data.append(row_data)

    return column_names, data, severity_list, visibility_list

def main():
    path_name = "C:/Users/andre/OneDrive/Documents/M.A/DA/IntroductionToDataAnalysis/Task2/Dannye_dlya_Praktiki_2_chast_2"
    
    all_files = Task1.task1.get_all_files(path_name)
    columns, data, severity_list, visibility_list = all_data(all_files)
    print("\n\t-> Column names and inexes:\n", columns)
    print("\n\t-> All data:\n", pd.Series(data))
    
    # Sevirity type distribution plot
    sevirity_types_counter = Counter(severity_list)
    plt.bar(sevirity_types_counter.keys(), sevirity_types_counter.values())
    plt.xlabel("Тип тяжести")
    plt.ylabel("Количество наблюдений события")
    plt.title("Распределение событий по типу")
    #plt.show()

    severity_list.sort()
    sevirity_types_counter = Counter(severity_list)
    sum = 0
    summary_distribution = {}
    for count in sevirity_types_counter.values():
        sum += count
        summary_distribution[sum] = sum
    plt.bar(sevirity_types_counter.keys(), summary_distribution)
    plt.xlabel("Тип тяжести по возрастанию")
    plt.ylabel("Накопленное количество наблюдений события")
    plt.title("Распределение событий по типу")
    plt.show()

if __name__ == "__main__":
    main()