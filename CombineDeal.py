import pandas as pd

# 读取 Excel 文件
excel_file = pd.ExcelFile('/mnt/七月第一周订单数据.xlsx')

# 获取所有表名
sheet_names = excel_file.sheet_names

# 创建一个空字典，用于存储合并后的数据
combined_data = {}

# 遍历每个工作表
for sheet_name in sheet_names:
    df = excel_file.parse(sheet_name)

    # 将账单类型和订餐内容列转换为字符串类型
    df['账单类型'] = df['账单类型'].astype(str)
    df['订餐内容'] = df['订餐内容'].astype(str)

    # 根据账单类型添加前缀到订餐内容
    df['订餐内容'] = df.apply(lambda row: row['账单类型'] + '：' + row['订餐内容'], axis=1)

    # 按姓名列进行分组，合并订餐内容和金额列
    grouped = df.groupby('姓名').agg({
        '订餐内容': lambda x: ', '.join(str(i) for i in x),
        '金额': 'sum'
    }).reset_index()

    # 将结果存储到字典中，键为表名
    combined_data[sheet_name] = grouped

# 创建一个新的 Excel 文件
with pd.ExcelWriter('/mnt/七月第一周订单数据_合并_带前缀.xlsx') as writer:
    # 遍历字典，将每个工作表的数据写入新文件
    for sheet_name, data in combined_data.items():
        data.to_excel(writer, sheet_name=sheet_name, index=False)