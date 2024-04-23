from os import listdir
from os.path import isfile, join

import xlrd

from collections import Counter

import pandas as pd

# Part a, b
def get_all_files(path_name="C:/Users/andre/OneDrive/Documents/M.A/DA/IntroductionToDataAnalysis/Task1/Dannye_dlya_praktiki_1"):
    return [join(path_name, file) for file in listdir(path_name) 
                if isfile(join(path_name, file)) and file.endswith(".xlsx")]

def construct_dict(path_name="C:/Users/andre/OneDrive/Documents/M.A/DA/IntroductionToDataAnalysis/Task1/Dannye_dlya_praktiki_1",
                   sheet_name="Отчет"):
    region_data = {}
    for file in get_all_files(path_name):
        print(file)
        book = xlrd.open_workbook(file)
        sheet = book.sheet_by_name(sheet_name)
        nrows = sheet.nrows
        ncols = sheet.ncols

        # Get the indexes of current years
        current_indexes = []
        year_indexes = {"2015 г.": 0, "2016 г.": 1, "2017 г.": 2, "2018 г.": 3}
        for y in range(1, 3):
            current_indexes.append(year_indexes[sheet.cell(1, y).value])
        print(current_indexes)

        if ncols != 3:
            raise Exception("Unprocessable file ", file)
        print("Num of rows: ", nrows,
              "Num of cols: ", ncols)
        for row in range(3, nrows):
            region_name = sheet.cell(row, 0).value
            if (region_name.find("федеральный округ") == -1):
                if region_data.get(region_name) is None:
                    region_data[region_name] = ["", "", "", ""]
                
                region_data[region_name][current_indexes[0]] = sheet.cell(row, 1).value
                region_data[region_name][current_indexes[1]] = sheet.cell(row, 2).value

    return region_data

# Part c
def process_data(region_data):
    if region_data is None:
        raise Exception("Unprocessable file")
    years_dict = {0: "2015", 1: "2016", 2: "2017", 3: "2018"}
    for year_values in region_data.values():
        index = 0
        count = 0
        avg = 0
        max_value = 0
        max_index = 0
        for year in year_values:
            if year == '':
                index += 1
                continue
            val = float(year)
            # avg
            avg += val
            count += 1
            # max
            if val > max_value:
                max_value = val
                max_index = index
            index += 1
        
        year_values.append(str(avg / count))
        year_values.append(years_dict[max_index])
    
    return region_data

# Part d
def compare_years(region_data):
    if region_data is None:
        raise Exception("Unprocessable file")
    years_count = Counter([value[5] for value in region_data.values()])
    return years_count

# Part e
def fifth(list):
    return float(list[1][4])

def sort_by_average(region_data):
    if region_data is None:
        raise Exception("Unprocessable file")
    return dict(sorted(region_data.items(), key=fifth))

# Main tests
def main():
    print("\n\t\t\t ------ A, B --------")
    data = construct_dict()
    print(pd.Series(data))

    print("\n\t\t\t ------ C --------")
    print(pd.Series(process_data(data)))
    
    print("\n\t\t\t ------ D --------")
    print(pd.Series(compare_years(data)))

    print("\n\t\t\t ------ E --------")
    print(pd.Series(sort_by_average(data)))

if __name__ == "__main__":
    main()