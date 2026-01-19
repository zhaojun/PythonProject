import pandas as pd

# ===================== 【修改这4处，按需调整】 =====================
excel_file_path = r"C:\Users\admin\Desktop\欠款-0114.xlsx"  # 你的Excel文件名
group_col = "地址"                    # 分组列：地址
need_cols = ["姓名", "地址", "合计欠款"]   # 只导出这几列，按需增删，比如加"手机号"
sort_col = "合计欠款"                     # 按欠款金额排序，不需要排序就注释掉下面的sort_values
# ===================================================================

# 1. 读取Excel，只加载需要的列，过滤空数据
df = pd.read_excel(excel_file_path, usecols=need_cols).dropna()

# 2. 按欠款金额【从高到低】排序（升序把ascending改成True即可）
df = df.sort_values(by=sort_col, ascending=False)

# 3. 生成拆分后的Excel
with pd.ExcelWriter(r"C:\Users\admin\Desktop\按地址拆分-欠款-0114.xlsx", engine="openpyxl") as writer:
    for addr_name, group_data in df.groupby(group_col):
        group_data.to_excel(writer, sheet_name=str(addr_name), index=False)

print(f"✅ 定制版拆分完成！")
print(f"📌 导出列：{need_cols}")
print(f"📊 共 {len(df.groupby(group_col))} 个地址，总计 {len(df)} 条欠款记录")