import pandas as pd
import numpy as np
import os
import glob
import sys
import re

# flatten the list for all excel files
def flatten_list(l):
    return [item for sublist in l for item in sublist]

#uses glob to join path with all files ending with the .xlsx extention
def get_all_excel_files():
    # get current working directory
    path = os.getcwd()
    excel_files = []
    for folder, _, _ in os.walk(path):
        excel_files.append((glob.glob(os.path.join(folder, "*.xlsx"))))
    return flatten_list(excel_files)


def search_through_dataframe(input_string:str, df, file_name):
    for sheet_name, sheet_data in df.items():
        # if you delete NaN values, you delete the entire row (or column, depending on args)
        # if you don't fill the NaN values, you can't compare strings
        sheet_data = sheet_data.fillna('-')
        values = sheet_data.values
        for row in range(len(values)):
            for col in range(len(values[0])):
                check_string = str(values[row][col]).lower()
                # check if the search string is in each cell of the excel file, in order to track the row and column
                if input_string in check_string:
                    print(f'File: {file_name}; Sheet: {sheet_name}; Location: ({row}, {col})')

def main():
    excel_files = get_all_excel_files()
    while True:
        input_string = input('What are you looking for: ').strip().lower()
        if input_string == 'end of work':
            break
        for f in excel_files:
            # take the last 2 inputs of the file path, which inlcudes the containing folder and the file name
            file_name = f.split("\\")[-2:]
            try:
                df = pd.read_excel(f, header=None, sheet_name=None)
                search_through_dataframe(input_string, df, file_name)
            except PermissionError:
                # certain temporary or corrupted files cannot be accessed
                print(f'NO PERMISSION TO OPEN {file_name}')
            except ValueError:
                # certain temporary or corrupted files can be accessed, but throw value error
                print(f'FAILED TO OPEN {file_name}')
            except:
                # handle unexpected errors
                print(f'UNKOWN ERROR FOR {file_name}')
            


if __name__ == "__main__":
    main()
