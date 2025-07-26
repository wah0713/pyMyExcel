import sys
from tools import ensure_xlsx_suffix
import openpyxl


def read(file_name):
    wd= openpyxl.load_workbook(ensure_xlsx_suffix(file_name))
    ws= wd.active
    for row in ws.iter_rows(values_only=True):
        for cell in row:
            print(f'{cell:<10}',end='\t')
        print('')

if __name__ == "__main__":
    read(sys.argv[1])