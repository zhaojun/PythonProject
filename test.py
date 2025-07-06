import pandas as pd

# 读取第一个 Excel 文件
excel_file1 = pd.ExcelFile(r'C:\Users\Administrator\Desktop\2025年6月30日午餐.xlsx')

# 读取第二个 Excel 文件
excel_file2 = pd.ExcelFile(r'C:\Users\Administrator\Desktop\菜品贴纸打印.xlsx')

# 获取第一个文档'Sheet1'工作表的数据
df1 = excel_file1.parse('Sheet1')

# 遍历不同工作表，合并第二个文档中的菜品价格信息
dfs = []
sheet_names = ['全部菜品', '产品', '早餐', '售价']
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
food_columns = df1.columns[9:-1]

# 获取 df1 中订餐内容列的菜品名称
order_food_names = food_columns.tolist()

# 获取 df2 中的菜品名称
price_food_names = df2['名称'].tolist()

# 找出在 df1 订餐内容中但不在 df2 价格表中的菜品
missing_foods = [food for food in order_food_names if food not in price_food_names]

if missing_foods:
    print('以下菜品在价格表中没有找到价格，请先添加价格：')
    for food in missing_foods:
        print(food)
else:
    print('所有菜品在价格表中都有对应的价格，可以进行金额计算操作。')
