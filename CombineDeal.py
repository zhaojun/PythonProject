import pandas as pd

# 读取 Excel 文件
file = r'D:\WPS云盘\1214901082\WPS云盘\工作\沈飞\订单数据\11月午餐\1月16日订餐数据.xlsx新.xlsx'
# file = r'D:\餐车午餐账单备份\11月第一周\2025年11月26日午餐.xlsx新.xlsx'
# file2 = r'D:\WPS云盘\1214901082\WPS云盘\工作\沈飞\订单数据\10月午餐\2025年10月29日.xlsx'
file2 = r'D:\餐车账单备份\2026年1月16日.xlsx'
excel_file = pd.ExcelFile(file)

# 获取所有表名
sheet_names = excel_file.sheet_names

# 创建一个空字典，用于存储合并后的数据
combined_data = {}

# 遍历每个工作表
for sheet_name in sheet_names:
    df = excel_file.parse(sheet_name)

    # 将账单类型和订餐内容列转换为字符串类型
    df['账目类型'] = df['账目类型'].astype(str)
    df['订餐内容'] = df['订餐内容'].astype(str)

    # 根据账单类型添加前缀到订餐内容
    df['订餐内容'] = df.apply(lambda row: row['账目类型'] + '：' + row['订餐内容'], axis=1)

    # 按姓名列进行分组，合并订餐内容和金额列
    grouped = df.groupby('唯一ID').agg({
        '订餐内容': lambda x: ', '.join(str(i) for i in x),
        '金额': 'sum'
    }).reset_index()

    # 将结果存储到字典中，键为表名
    combined_data[sheet_name] = grouped

# 创建一个新的 Excel 文件
with pd.ExcelWriter(file2) as writer:
    # 遍历字典，将每个工作表的数据写入新文件
    for sheet_name, data in combined_data.items():
        data.to_excel(writer, sheet_name=sheet_name, index=False)