import sys
import os
from tools import ensure_xlsx_suffix


def open(file_name):
    os.startfile(file_name)

if __name__ == "__main__":
    open(ensure_xlsx_suffix(sys.argv[1]))