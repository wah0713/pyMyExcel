import sys
import os
from tools import ensure_xlsx_suffix


def main(file_name):
    print(file_name)
    os.startfile(ensure_xlsx_suffix(file_name))


if __name__ == "__main__":
    main(sys.argv[1])
