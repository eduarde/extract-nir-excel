import os
import sys
import pandas as pd


def get_date_file(filename):
    try:
        cols = [8]
        data = pd.read_excel(filename, usecols=cols)
        data.head()
        df = data
        return df.iloc[5][0]
    except IndexError:
        return ""


def read_file(filename):
    cols = [2, 3, 4, 5]
    data = pd.read_excel(filename, usecols=cols)
    data.head()
    df = data 
    df = df.dropna()
    df = df.iloc[2:]
    df['Date'] = pd.to_datetime(get_date_file(filename)).date()
    return df


def write_file(df, filename):
    df.to_excel(filename, sheet_name='Total', index=False)


def main():
    try:
        directory = sys.argv[1]
    except IndexError:
        directory = './files'
    
    result_file_excel = './results.xlsx'
    if os.path.exists(result_file_excel):
        os.remove(result_file_excel)

    df_result = pd.DataFrame()
    for entry in os.scandir(directory):
        if (entry.path.endswith(".xlsx") or entry.path.endswith(".XLS") or entry.path.endswith(".xls") and entry.is_file()):
            df_excel = read_file(entry.path)
            df_result = pd.concat([df_result, df_excel], axis=0)
    write_file(df_result, result_file_excel)


if __name__ == "__main__":
    main()

