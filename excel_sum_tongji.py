import pandas as pd
import numpy as np

def calculate_excel_sum_by_address_group(
    file_path, address_col, value_range_str, address_groups, sheet_name=None
):
    """
    自定义地址组求和（忽略大小写，按指定模板输出：文件名+组名+份数）
    :param file_path: Excel文件路径
    :param address_col: 地址所在列（Excel列字母）
    :param value_range_str: 数值统计区域（Excel格式）
    :param address_groups: 自定义地址组（字典：{组名: [地址列表]}）
    :param sheet_name: 工作表名称（默认None）
    :return: 组合计字典 + 整体总和
    """
    try:
        # ---------------------- 1. 解析数值区域 ----------------------
        if ":" in value_range_str:
            val_start_cell, val_end_cell = value_range_str.split(":")
        else:
            val_start_cell = val_end_cell = value_range_str

        def col_to_num(col_str):
            num = 0
            for c in col_str.upper():
                num = num * 26 + (ord(c) - ord('A') + 1)
            return num - 1

        val_start_col = ''.join([c for c in val_start_cell if c.isalpha()])
        val_start_row = int(''.join([c for c in val_start_cell if c.isdigit()])) - 1
        val_end_col = ''.join([c for c in val_end_cell if c.isalpha()])
        val_end_row = int(''.join([c for c in val_end_cell if c.isdigit()])) - 1
        address_col_idx = col_to_num(address_col)

        # ---------------------- 2. 数据读取与清理 ----------------------
        value_col_indices = range(col_to_num(val_start_col), col_to_num(val_end_col) + 1)
        all_col_indices = [address_col_idx] + list(value_col_indices)
        df = pd.read_excel(
            io=file_path,
            sheet_name=sheet_name,
            usecols=all_col_indices,
            skiprows=val_start_row,
            nrows=val_end_row - val_start_row + 1,
            header=None,
            engine="openpyxl"
        )

        # 地址清理（忽略大小写、去空格）
        col_names = ["地址"] + [f"数值列{i+1}" for i in range(len(value_col_indices))]
        df.columns = col_names
        df["地址"] = df["地址"].astype(str).str.lower().str.strip().str.replace("\n", "")
        df = df[df["地址"] != "nan"]

        # 数值列清理（非数字转0）
        df[col_names[1:]] = df[col_names[1:]].apply(pd.to_numeric, errors="coerce").fillna(0)

        # ---------------------- 3. 地址分组与求和 ----------------------
        address_sum = df.groupby("地址")[col_names[1:]].sum().sum(axis=1)
        address_sum = np.round(address_sum, 2).astype(float)

        # 地址组映射（忽略大小写）
        addr_to_group = {}
        for group_name, addrs in address_groups.items():
            clean_addrs = [str(addr).lower().strip().replace("\n", "") for addr in addrs]
            for addr in clean_addrs:
                addr_to_group[addr] = group_name

        # 按组汇总
        group_sum = {}
        other_sum = 0.0
        for addr, sum_val in address_sum.items():
            if addr in addr_to_group:
                group_name = addr_to_group[addr]
                group_sum[group_name] = group_sum.get(group_name, 0.0) + (sum_val if not pd.isna(sum_val) else 0.0)
            else:
                other_sum += sum_val if not pd.isna(sum_val) else 0.0

        # 添加其他组（如有未匹配地址）
        if other_sum > 0:
            group_sum["其他组"] = np.round(other_sum, 2)

        # 按组名顺序输出（保持自定义地址组的顺序，其他组在最后）
        group_order = list(address_groups.keys()) + (["其他组"] if "其他组" in group_sum else [])
        group_sum_ordered = {group: group_sum[group] for group in group_order if group in group_sum}
        overall_total = np.round(sum(group_sum_ordered.values()), 2)

        # ---------------------- 4. 按指定模板输出 ----------------------
        print(f"{file_path.split('/')[-1].split('\\')[-1]}")  # 只显示文件名（不含路径）
        for group_name, sum_val in group_sum_ordered.items():
            print(f"{group_name}：{sum_val:.0f}份")  # 份数保留整数（如需小数，把.0f改成.2f）
        print(f"合计：{overall_total:.0f}份")  # 新增合计行（可选，不需要可删除）

        return group_sum_ordered, overall_total

    # 异常处理（简洁提示）
    except FileNotFoundError:
        print(f"❌ 找不到文件：{file_path}")
        return None, None
    except (ValueError, IndexError):
        print(f"❌ 格式错误：数值区域{value_range_str}或地址列{address_col}无效")
        return None, None
    except ModuleNotFoundError:
        print(f"❌ 缺少库，请执行：pip install pandas openpyxl numpy")
        return None, None
    except Exception as e:
        print(f"❌ 错误：{str(e)}")
        return None, None


# ------------------- 核心参数配置（重点修改这里！）-------------------
if __name__ == "__main__":
    # 1. 基础配置（必改）
    EXCEL_FILE =r"D:\WPS云盘\1214901082\WPS云盘\工作\沈飞\订单数据\11月午餐\1月14日订餐数据.xlsx"  # 你的Excel文件路径（Windows示例："C:/Users/张三/Desktop/销售数据.xlsx"）
    ADDRESS_COL = "J"  # 地址所在列（Excel列字母，如"D"、"F"）
    VALUE_RANGE = "K2:AG1000"  # 数值统计区域（Excel格式，如"A1:C5"、"B3:E10"）
    SHEET_NAME = "Sheet1"  # 工作表名称（留None自动读第一个）

    # 2. 自定义地址组（核心！按你的需求修改组名和地址）
    # 格式：{ "组名1": [地址1, 地址2, ...], "组名2": [地址3, 地址4, ...] }
    CUSTOM_ADDRESS_GROUPS = {
        "兵哥": ["76#","708","709A","902","902","755小门", "654D", "701西", "701北", "701东", "781", "782钛合金", "786南", "786北", "动力水暖", "452#", "570","70号北门线束厂","713a","1机库", "2机库", "3机库", "4机库", "可靠性", "数据中心", "402", "411", "412西", "503", "345厂", "48", "506", "55", "55#", "535", "560南", "570", "70#东北", "70#（29厂）", "73#（45厂）贩卖机旁", "70#西南门（21厂）", "70#西北门（33厂）", "70#西北门（45厂）", "702西", "706", "706A", "709", "709A", "710器材保障部", "712", "713a", "72","72#", "73#", "73", "73#西南门", "74#", "770", "780A", "782数控", "785", "787", "812南", "812北", "85厂房", "94#", "大白楼", "老科技楼", "试飞站白楼", "试飞站汽车队", "试飞站充电班", "试飞站天窗", "培训中心", "器材采购部露天库电动拉门", "档案馆", "706南", "设备维保", "回收科", "试飞站小油库", "设备维保", "766A", "客户服务部", "行政科楼", "培训中心三号楼", "培训中心3号楼", "545", "试飞站篮球架", "545南门玻璃门", "723", "77#北门", "745", "784","73#（45厂）贩卖机旁", "器材动力门" ]
    }

    # 执行统计
    group_totals, overall_total = calculate_excel_sum_by_address_group(
        file_path=EXCEL_FILE,
        address_col=ADDRESS_COL,
        value_range_str=VALUE_RANGE,
        address_groups=CUSTOM_ADDRESS_GROUPS,
        sheet_name=SHEET_NAME
    )
