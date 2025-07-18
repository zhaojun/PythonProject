import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox


def add_missing_dishes():
    # 获取订单文件路径
    order_file = order_file_entry.get()
    # 获取菜品价格文件路径
    price_file = price_file_entry.get()

    if not order_file or not price_file:
        messagebox.showerror("错误", "请选择订单文件和菜品价格文件！")
        return

    try:
        # 读取订单文件
        excel_file1 = pd.ExcelFile(order_file)
        # 读取菜品价格文件
        excel_file2 = pd.ExcelFile(price_file)

        # 获取订单文件 'Sheet1' 工作表的数据
        df1 = excel_file1.parse('Sheet1')
        # 获取菜品价格文件 'Sheet1' 工作表的数据
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

            # 将新菜品添加到 '全部菜品' DataFrame
            df_all_foods = pd.concat([df_all_foods, new_foods_df], ignore_index=True)

            # 创建 ExcelWriter 对象
            with pd.ExcelWriter('菜品贴纸打印-更新.xlsx', engine='openpyxl', mode='w') as writer:
                # 获取所有表名
                sheet_names = excel_file2.sheet_names

                # 遍历每个工作表
                for sheet_name in sheet_names:
                    # 获取当前工作表的数据
                    df = excel_file2.parse(sheet_name)

                    # 如果是 '全部菜品' 工作表，则写入更新后的数据
                    if sheet_name == 'Sheet1':
                        df_all_foods.to_excel(writer, sheet_name=sheet_name, index=False)
                    else:
                        # 其他工作表保持不变
                        df.to_excel(writer, sheet_name=sheet_name, index=False)

            messagebox.showinfo("成功", f'已将 {len(missing_foods)} 个缺失的菜品添加到 "Sheet1" 工作表中，请在价格列中填写价格。\n缺失的菜品：\n{" ".join(missing_foods)}')
        else:
            messagebox.showinfo("提示", '所有菜品在 "全部菜品" 工作表中都已存在。')
    except Exception as e:
        messagebox.showerror("错误", f"处理文件时出现错误：{str(e)}")


def browse_order_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        order_file_entry.delete(0, tk.END)
        order_file_entry.insert(0, file_path)


def browse_price_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        price_file_entry.delete(0, tk.END)
        price_file_entry.insert(0, file_path)


# 创建主窗口
root = tk.Tk()
root.title("添加缺失菜品")

# 创建标签和输入框
order_file_label = tk.Label(root, text="订单文件:")
order_file_label.pack()
order_file_entry = tk.Entry(root, width=50)
order_file_entry.pack()
order_file_button = tk.Button(root, text="浏览...", command=browse_order_file)
order_file_button.pack()

price_file_label = tk.Label(root, text="菜品价格文件:")
price_file_label.pack()
price_file_entry = tk.Entry(root, width=50)
price_file_entry.pack()
price_file_button = tk.Button(root, text="浏览...", command=browse_price_file)
price_file_button.pack()

# 创建处理按钮
process_button = tk.Button(root, text="添加缺失菜品", command=add_missing_dishes)
process_button.pack()

# 运行主循环
root.mainloop()