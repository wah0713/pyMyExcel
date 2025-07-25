import sys
from tools import ensure_xlsx_suffix
import openpyxl


def read(file_name):
    wd= openpyxl.load_workbook(file_name)
    ws=wd.active
    for row in ws.iter_rows(values_only=True):
        for cell in row:
            print(f'{cell.value:<10}')

if __name__ == "__main__":
    read(ensure_xlsx_suffix(sys.argv[1]))