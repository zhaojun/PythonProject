import pandas as pd
from datetime import datetime
#午餐
BaseDirectory = r'D:\WPS云盘\1214901082\WPS云盘\工作\沈飞\订单数据\7月'
FileStore = [BaseDirectory+'\\2025年7月16日午餐.xlsx']

# #早餐
# BaseDirectory = r'D:\WPS云盘\1214901082\WPS云盘\工作\沈飞\订单数据\7月第二周早餐'
# FileStore = [BaseDirectory+'\\0708早餐.xlsx']# 读取第一个 Excel 文件
file2 = r'D:\WPS云盘\1214901082\WPS云盘\工作\沈飞\订单数据\菜品贴纸打印.xlsx'
for file1 in FileStore:
    print(file1)
    excel_file1 = pd.ExcelFile(file1)

    # 读取第二个 Excel 文件
    excel_file2 = pd.ExcelFile(file2)

    # 获取第一个文档'Sheet1'工作表的数据
    df1 = excel_file1.parse('Sheet1')

    # 获取第二个文档'全部菜品'工作表的数据
    df_all_foods = excel_file2.parse('Sheet1')

    # 提取 df1 中订餐内容相关列
    food_columns = df1.columns[10:-1]

    # 获取 df1 中订餐内容列的菜品名称
    order_food_names = food_columns.tolist()

    # 获取 df_all_foods 中的菜品名称
    all_food_names = df_all_foods['名字'].tolist()

    # 找出在 df1 订餐内容中但不在 df_all_foods 中的菜品
    missing_foods = [food for food in order_food_names if food not in all_food_names]

    if missing_foods:
        # 获取当前日期
        current_date = datetime.now().strftime('%Y-%m-%d')
        # 创建新的菜品 DataFrame，价格列留空
        new_foods_df = pd.DataFrame({
            '名字': missing_foods,
            '价格': [None] * len(missing_foods),
            '添加时间': [current_date] * len(missing_foods)
        })

        # 将新菜品添加到'全部菜品'DataFrame
        df_all_foods = pd.concat([df_all_foods, new_foods_df], ignore_index=True)

        # 创建 ExcelWriter 对象
        with pd.ExcelWriter(r'D:\WPS云盘\1214901082\WPS云盘\工作\沈飞\订单数据\菜品贴纸打印-更新.xlsx', engine='openpyxl', mode='w') as writer:
            # 获取所有表名
            sheet_names = excel_file2.sheet_names

            # 遍历每个工作表
            for sheet_name in sheet_names:
                # 获取当前工作表的数据
                df = excel_file2.parse(sheet_name)

                # 如果是'全部菜品'工作表，则写入更新后的数据
                if sheet_name == 'Sheet1':
                    df_all_foods.to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    # 其他工作表保持不变
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f'已将 {len(missing_foods)} 个缺失的菜品添加到"Sheet1"工作表中，请在价格列中填写价格。')
        print('缺失的菜品：')
        for food in missing_foods:
            print(food)
    else:
        print('所有菜品在"全部菜品"工作表中都已存在。')