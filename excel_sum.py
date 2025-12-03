import pandas as pd

def calculate_excel_sum(file_path, sheet_name=None, range_str="A1:A10"):
    """
    统计Excel指定区域的数字和（PyCharm专用，控制台清晰输出）
    :param file_path: Excel文件路径（相对路径/绝对路径均可）
    :param sheet_name: 工作表名称（默认None，自动读取第一个工作表）
    :param range_str: 统计区域（Excel格式，如 "A1:C5"、"B3:E10"，默认"A1:A10"）
    :return: 区域内数字总和（同时在控制台打印结果）
    """
    try:
        # 解析Excel区域（支持单个单元格/连续区域）
        if ":" in range_str:
            start_cell, end_cell = range_str.split(":")
        else:
            start_cell = end_cell = range_str  # 处理单个单元格（如 "F4"）

        # 列字母转数字索引（A→0、B→1...，适配pandas读取逻辑）
        def col_to_num(col_str):
            num = 0
            for c in col_str:
                num = num * 26 + (ord(c.upper()) - ord('A') + 1)
            return num - 1  # 转为0开始的索引

        # 提取起始/结束的列（字母）和行（数字）
        start_col = ''.join([c for c in start_cell if c.isalpha()])
        start_row = int(''.join([c for c in start_cell if c.isdigit()])) - 1  # pandas行索引从0开始
        end_col = ''.join([c for c in end_cell if c.isalpha()])
        end_row = int(''.join([c for c in end_cell if c.isdigit()])) - 1

        # 精准读取指定区域数据
        df = pd.read_excel(
            io=file_path,
            sheet_name=sheet_name,
            usecols=range(col_to_num(start_col), col_to_num(end_col) + 1),  # 列范围
            skiprows=start_row,  # 跳过前面的行
            nrows=end_row - start_row + 1,  # 要读取的行数
            header=None,  # 不设表头，避免数据误判
            engine="openpyxl"  # 解析.xlsx文件（.xls需改xlrd，见注意事项）
        )

        # 清理数据：非数字值（文本、空值）转为NaN并过滤
        df_numeric = df.apply(pd.to_numeric, errors="coerce")  # 非数字→NaN
        df_numeric = df_numeric.dropna(how="all", axis=0).dropna(how="all", axis=1)  # 删除全空行/列

        # 计算总和
        total_sum = df_numeric.sum().sum()
        valid_num_count = df_numeric.notna().sum().sum()  # 有效数字个数

        # PyCharm控制台清晰输出（带分隔线，一目了然）
        print("=" * 60)
        print("📊 Excel 区域求和工具（PyCharm版）")
        print("=" * 60)
        print(f"📁 目标文件：{file_path}")
        print(f"📋 工作表：{sheet_name if sheet_name else '【默认第一个工作表】'}")
        print(f"📍 统计区域：{range_str}")
        print(f"🔢 区域内有效数字个数：{valid_num_count} 个")
        print(f"\n🎉 最终计算结果：{total_sum:.2f}")
        print("=" * 60)

        return total_sum

    # 异常处理（明确提示错误原因，方便排查）
    except FileNotFoundError:
        print(f"❌ 错误：找不到文件「{file_path}」！")
        print("   请检查：1. 文件路径是否正确；2. 文件是否和代码在同一文件夹（相对路径）。")
        return None
    except (ValueError, IndexError):
        print(f"❌ 错误：统计区域「{range_str}」格式无效！")
        print("   正确格式示例：A1:C5（矩形区域）、B3:E10（多行多列）、D3:D15（单列）、F4（单个单元格）")
        return None
    except ModuleNotFoundError:
        print(f"❌ 错误：缺少必要库！请先在PyCharm终端执行安装命令：")
        print("   pip install pandas openpyxl")
        return None
    except Exception as e:
        print(f"❌ 未知错误：{str(e)}")
        return None

# ------------------- 核心参数配置（在这里修改！）-------------------
if __name__ == "__main__":
    # 1. Excel文件路径（关键！按实际情况修改）
    # 相对路径（文件和代码在同一文件夹）：直接写文件名，如 "销售数据.xlsx"
    # 绝对路径（文件在任意位置）：Windows示例 "C:/Users/你的名字/Desktop/销售数据.xlsx"
    #                          Mac示例 "/Users/你的名字/Desktop/销售数据.xlsx"
    EXCEL_FILE = r"D:\WPS云盘\1214901082\WPS云盘\工作\沈飞\订单数据\11月午餐\2025年11月7日午餐.xlsx新.xlsx"  # 👉 这里改你的Excel文件路径

    # 2. 工作表名称（如 "Sheet1"、"销售数据"，不确定就留 None）
    SHEET_NAME = "Sheet1"  # 👉 这里改工作表名称（可选，留None自动读第一个）

    # 3. 要统计的区域（Excel格式，如 "A1:C5"）
    TARGET_RANGE = "K2:AJ1000"  # 👉 这里改你要统计的区域

    # 执行计算（结果直接显示在PyCharm控制台）
    calculate_excel_sum(EXCEL_FILE, SHEET_NAME, TARGET_RANGE)