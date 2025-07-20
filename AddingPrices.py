import pandas as pd
from datetime import datetime
#午餐
# BaseDirectory = r'D:\WPS云盘\1214901082\WPS云盘\工作\沈飞\订单数据\7月'
# FileStore = [BaseDirectory+'\\2025年7月18日午餐.xlsx']#,BaseDirectory+'\\2025年7月9日午餐.xlsx']# 读取第一个 Excel 文件
#早餐
BaseDirectory = r'D:\WPS云盘\1214901082\WPS云盘\工作\沈飞\订单数据\7月第四周早餐'
FileStore = [BaseDirectory+'\\0721早餐.xlsx']# 读取第一个 Excel 文件#

file2 = r'D:\WPS云盘\1214901082\WPS云盘\工作\沈飞\订单数据\菜品贴纸打印.xlsx'
for file1 in FileStore:
    # 读取第一个 Excel 文件
    excel_file1 = pd.ExcelFile(file1)

    # 读取第二个 Excel 文件
    excel_file2 = pd.ExcelFile(file2)

    # 获取第一个文档'Sheet1'工作表的数据
    df1 = excel_file1.parse('Sheet1')

    # 遍历不同工作表，合并第二个文档中的菜品价格信息
    dfs = []
    sheet_names = ['Sheet1']
    for sheet_name in sheet_names:
        df = excel_file2.parse(sheet_name)
        # 统一列名
        if '名字' in df.columns:
            df = df.rename(columns={'名字': '名称'})
        if '价格' in df.columns:
            df = df.rename(columns={'价格': '单价'})
        dfs.append(df[['名称', '单价']])
    df2 = pd.concat(dfs, ignore_index=True)

    # 提取 df1 中订餐内容相关列
    food_columns = df1.columns[10:-1]
    df1[food_columns] = df1[food_columns].fillna(0)
    # 初始化总金额列
    df1['金额'] = 0

    # 遍历每个订餐内容相关列
    for col in food_columns:
        # 获取当前菜品的名称和价格
        food_name = col
        food_price = df2[df2['名称'] == food_name]['单价'].values[0] if food_name in df2['名称'].values else 0

        # 计算当前菜品的总金额并累加到总金额列
        df1['金额'] += df1[col] * food_price

    # 将结果保存为 Excel 文件
    NewFile = file1+'新.xlsx'
    df1.to_excel(NewFile, index=False)