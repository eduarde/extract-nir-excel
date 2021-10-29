import os
import sys
import pandas as pd

def read_file(filename):
    cols = [2, 3, 4, 5]
    data = pd.read_excel(filename, usecols=cols)
    data.head()
    df = data
    df = df.dropna()
    df = df.iloc[2:]
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

