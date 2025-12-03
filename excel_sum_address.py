import pandas as pd

def calculate_excel_sum_by_address(file_path, address_col, value_range_str, sheet_name=None):
    """
    按地址分组统计Excel指定数值区域的合计（保留自定义数值区域功能）
    :param file_path: Excel文件路径（相对/绝对路径）
    :param address_col: 地址所在列（Excel列字母，如 "D"、"F"）
    :param value_range_str: 要统计的数值区域（Excel格式，如 "A1:C5"、"B3:E10"）
    :param sheet_name: 工作表名称（默认None，读取第一个工作表）
    :return: 按地址分组的合计字典 + 整体总和
    """
    try:
        # ---------------------- 1. 解析数值区域（复制原有逻辑，确保自定义区域生效）----------------------
        if ":" in value_range_str:
            val_start_cell, val_end_cell = value_range_str.split(":")
        else:
            val_start_cell = val_end_cell = value_range_str  # 单个单元格

        # 列字母转数字索引（A→0、B→1...）
        def col_to_num(col_str):
            num = 0
            for c in col_str.upper():
                num = num * 26 + (ord(c) - ord('A') + 1)
            return num - 1  # 转为0开始的索引

        # 解析数值区域的行列范围
        val_start_col = ''.join([c for c in val_start_cell if c.isalpha()])
        val_start_row = int(''.join([c for c in val_start_cell if c.isdigit()])) - 1
        val_end_col = ''.join([c for c in val_end_cell if c.isalpha()])
        val_end_row = int(''.join([c for c in val_end_cell if c.isdigit()])) - 1

        # 解析地址列的数字索引（如 "D" → 3）
        address_col_idx = col_to_num(address_col)

        # ---------------------- 2. 读取数据（同时包含地址列 + 数值区域）----------------------
        # 计算需要读取的所有列（地址列 + 数值区域的列）
        value_col_indices = range(col_to_num(val_start_col), col_to_num(val_end_col) + 1)
        all_col_indices = [address_col_idx] + list(value_col_indices)  # 地址列在前，数值列在后

        df = pd.read_excel(
            io=file_path,
            sheet_name=sheet_name,
            usecols=all_col_indices,  # 只读取需要的列（地址列 + 数值列）
            skiprows=val_start_row,  # 从数值区域的起始行开始读取
            nrows=val_end_row - val_start_row + 1,  # 读取数值区域的总行数
            header=None,  # 不设表头，避免数据误判
            engine="openpyxl"
        )

        # ---------------------- 3. 数据清理与命名 ----------------------
        # 列重命名：第0列=地址列，后面的列=数值列
        col_names = ["地址"] + [f"数值列{i+1}" for i in range(len(value_col_indices))]
        df.columns = col_names

        # 清理数据：地址列去重空值，数值列非数字转NaN
        df = df.dropna(subset=["地址"])  # 删除地址为空的行
        df[col_names[1:]] = df[col_names[1:]].apply(pd.to_numeric, errors="coerce")  # 数值列清理

        # ---------------------- 4. 按地址分组求和 ----------------------
        # 计算每组数值列的总和（按地址分组）
        address_sum = df.groupby("地址")[col_names[1:]].sum().sum(axis=1)  # 每行数值列求和后再按地址汇总
        address_sum = address_sum.round(2)  # 保留2位小数

        # 计算整体总和
        total_sum = address_sum.sum().round(2)

        # ---------------------- 5. PyCharm控制台输出 ----------------------
        print("=" * 70)
        print("📊 Excel 按地址分组求和工具（PyCharm版）")
        print("=" * 70)
        print(f"📁 目标文件：{file_path}")
        print(f"📋 工作表：{sheet_name if sheet_name else '【默认第一个工作表】'}")
        print(f"📍 地址列：{address_col}列")
        print(f"🔢 数值统计区域：{value_range_str}")
        print(f"📈 参与统计的地址数：{len(address_sum)} 个")
        print(f"\n【每个地址合计】")
        print("-" * 50)
        for addr, sum_val in address_sum.items():
            print(f"🏠 {addr}：{sum_val:.2f}")
        print("-" * 50)
        print(f"\n🎉 整体总和：{total_sum:.2f}")
        print("=" * 70)

        return address_sum.to_dict(), total_sum  # 返回字典（地址:合计）和整体总和

    # 异常处理（更贴合新增功能的错误提示）
    except FileNotFoundError:
        print(f"❌ 错误：找不到文件「{file_path}」！请检查路径是否正确。")
        return None, None
    except (ValueError, IndexError):
        print(f"❌ 错误：数值区域「{value_range_str}」或地址列「{address_col}」格式无效！")
        print("   正确示例：数值区域=A1:C5、地址列=D（单个字母，无需加行号）")
        return None, None
    except ModuleNotFoundError:
        print(f"❌ 错误：缺少库！请在PyCharm终端执行：pip install pandas openpyxl")
        return None, None
    except Exception as e:
        print(f"❌ 未知错误：{str(e)}")
        return None, None

# ------------------- 核心参数配置（在这里修改！）-------------------
if __name__ == "__main__":
    # 1. Excel文件路径（必改！相对/绝对路径均可）
    EXCEL_FILE =r"D:\WPS云盘\1214901082\WPS云盘\工作\沈飞\订单数据\11月午餐\11月14日午餐.xlsx" # 示例："C:/Users/张三/Desktop/销售数据.xlsx"（Windows）

    # 2. 地址所在列（必改！Excel列字母，如地址在D列就写"D"）
    ADDRESS_COL = "J"  # 👉 这里改地址列（例："B"、"F"、"H"）

    # 3. 数值统计区域（必改！Excel格式，如 "A1:C5"、"B3:E10"）
    VALUE_RANGE = "K2:AE1000"  # 👉 这里改要统计的数值区域
    # 4. 工作表名称（可选，留None自动读第一个工作表）
    SHEET_NAME = "Sheet1"  # 示例："销售数据"、None

    # 执行统计（结果显示在PyCharm控制台）
    address_totals, overall_total = calculate_excel_sum_by_address(
        file_path=EXCEL_FILE,
        address_col=ADDRESS_COL,
        value_range_str=VALUE_RANGE,
        sheet_name=SHEET_NAME
    )