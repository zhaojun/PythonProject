import pandas as pd

df = pd.read_excel(r"C:\Users\admin\Desktop\沈飞账单-最新.xlsx", engine="openpyxl")

# 连续行号：Excel第3-5行 → pandas索引2-4（切片左闭右开，所以是2:5）
df.drop(index=range(3734, 1050000), axis=0, inplace=True)
df.reset_index(drop=True, inplace=True)
df.to_excel("数据_删除连续行.xlsx", index=False, engine="openpyxl")