import pandas as pd

# ---------------------- 配置参数（只改这部分！）----------------------
excel_path = r"D:\WPS云盘\1214901082\WPS云盘\工作\沈飞\订单数据\11月午餐\2025年11月6日午餐.xlsx"  # 例如："订单数据.xlsx"
sheet_name = "Sheet1"
address_col = "取餐地址"  # J列的表头名称（必须和表格一致）
# 手动列写 K-AH 列的所有菜品表头（若太多，用下面的“批量获取”方法）
dish_cols = ["菜品1", "菜品2", "菜品3", ...]  # 替换为你K-AH列的实际表头
output_path = "地址-菜品求和结果.xlsx"
# ------------------------------------------------------------------------

# 1. 读取 Excel（统一用列名，无混合格式）
df = pd.read_excel(
    excel_path,
    sheet_name=sheet_name,
    usecols=[address_col] + dish_cols  # 只传列名列表，不混列字母
)

# 后续分组求和、保存代码不变（和之前一样）
df[address_col] = df[address_col].str.strip().str.lower()
result = df.groupby(address_col).sum().reset_index()
result.to_excel(output_path, index=False)
print("统计完成！")
