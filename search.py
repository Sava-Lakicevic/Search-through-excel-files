import pandas as pd
import numpy as np
import os
import glob
import sys
import re


def flatten_list(l):
    return [item for sublist in l for item in sublist]

def get_all_excel_files():
    path = os.getcwd()
    excel_files = []
    for folder, _, _ in os.walk(path):
        excel_files.append((glob.glob(os.path.join(folder, "*.xlsx"))))
    return flatten_list(excel_files)

# def main():
#     excel_files = get_all_excel_files()
#     input_string = ''
#     for arg in sys.argv[1:]:
#         input_string += arg + ' '
#     input_string = input_string.strip().lower()
#     for f in excel_files:
#         df = pd.read_excel(f, header=None, sheet_name=None)
#         file_name = f.split("\\")[-2:]
#         for sheet_name, sheet_data in df.items():
#             sheet_data = sheet_data.fillna('-')
#             values = sheet_data.values
#             for row in range(len(values)):
#                 for col in range(len(values[0])):
#                     check_string = str(values[row][col]).lower()
#                     if input_string in check_string:
#                         print(f'File: {file_name}; Sheet: {sheet_name}; Location: ({row}, {col})')

def main():
    excel_files = get_all_excel_files()
    while True:
        input_string = input('What are you looking for: ').strip().lower()
        if input_string == 'kraj rada':
            break
        for f in excel_files:
            file_name = f.split("\\")[-2:]
            try:
                df = pd.read_excel(f, header=None, sheet_name=None)
            except PermissionError:
                print(f'NO PERMISSION TO OPEN {file_name}')
            except ValueError:
                print(f'FAILED TO OPEN {file_name}')
            except:
                print(f'UNKOWN ERROR FOR {file_name}')
            for sheet_name, sheet_data in df.items():
                sheet_data = sheet_data.fillna('-')
                values = sheet_data.values
                for row in range(len(values)):
                    for col in range(len(values[0])):
                        check_string = str(values[row][col]).lower()
                        if input_string in check_string:
                            print(f'File: {file_name}; Sheet: {sheet_name}; Location: ({row}, {col})')


if __name__ == "__main__":
    main()