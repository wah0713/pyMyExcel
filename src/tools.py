import os
def ensure_xlsx_suffix(file_name):
    str = file_name
    root, ext = os.path.splitext(file_name)
    if not ext:
        str = root + ".xlsx"
    return str