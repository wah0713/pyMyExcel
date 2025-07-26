import sys
import openpyxl
from tools import ensure_xlsx_suffix


def parse_arguments():
    """解析命令行参数"""
    if len(sys.argv) != 9:
        print(
            "使用方法: python main.py <源文件> <源工作表> <源列> <起始行> <结束行> <目标文件> <目标工作表> <目标列> <目标开始行>"
        )
        print("示例: python main.py A.xlsx Sheet a 1 2 B.xlsx Sheet e 5")
        print("这将从 A.xlsx 工作表的 'Sheet' 表 A 列 1-2 行复制数据")
        print("到 B.xlsx 工作表的 'Sheet' 表 E 列从第5行开始写入")

    return {
        "source_file": ensure_xlsx_suffix(sys.argv[1]),
        "source_sheet": sys.argv[2],
        "source_col": sys.argv[3].upper(),  # 转换为大写字母
        "start_row": int(sys.argv[4]),
        "end_row": int(sys.argv[5]),
        "target_file": ensure_xlsx_suffix(sys.argv[6]),
        "target_sheet": sys.argv[7],
        "target_col": sys.argv[8].upper(),  # 转换为大写字母
        "target_start_row": int(sys.argv[9]),
    }


def copy_excel_data(params):
    """执行Excel数据复制操作"""
    try:
        # 加载源文件
        src_wb = openpyxl.load_workbook(params["source_file"])
        src_ws = src_wb[params["source_sheet"]]

        # 加载目标文件
        tgt_wb = openpyxl.load_workbook(params["target_file"])
        tgt_ws = tgt_wb[params["target_sheet"]]

        # 计算要复制的行数
        row_count = params["end_row"] - params["start_row"] + 1

        target_start_row = params["target_start_row"]

        # 复制数据
        for i in range(row_count):
            src_row = params["start_row"] + i
            tgt_row = target_start_row + i

        # 获取源单元格值
        src_cell = src_ws[f"{params['source_col']}{src_row}"]

        # 写入目标单元格
        print(f"{params['target_col']}{tgt_row}")
        tgt_ws[f"{params['target_col']}{tgt_row}"] = src_cell.value

        # 保存目标文件
        tgt_wb.save(params["target_file"])

        print(f"成功复制 {row_count} 行数据：")
        print(
            f"源文件: {params['source_file']} (工作表: {params['source_sheet']}, 列: {params['source_col']}, 行: {params['start_row']}-{params['end_row']})"
        )
        print(
            f"目标文件: {params['target_file']} (工作表: {params['target_sheet']}, 列: {params['target_col']}, 行: {target_start_row}-{tgt_row})"
        )

    except FileNotFoundError:
        print("错误：找不到指定的Excel文件，请检查路径是否正确")
    except KeyError as e:
        print(f"错误：找不到指定的工作表 - {str(e)}")
    except Exception as e:
        print(f"发生未知错误: {str(e)}")
    finally:
        if "src_wb" in locals():
            src_wb.close()
        if "tgt_wb" in locals():
            tgt_wb.close()


if __name__ == "__main__":
    params = parse_arguments()
    copy_excel_data(params)
# py src/copy_excel_data.py A.xlsx Sheet1 a 1 2 B.xlsx Sheet1 e 5