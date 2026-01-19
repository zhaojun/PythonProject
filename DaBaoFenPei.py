import pandas as pd
from datetime import datetime

# 重新读取并预处理数据，聚焦核心信息
df = pd.read_excel(r'C:\Users\admin\Desktop\打包测试.xlsx')
df_clean = df.copy()

# 数据预处理：只保留核心字段，处理缺失值
df_clean = df_clean[['地址', '人次', '打包\n人员']].copy()
df_clean['打包\n人员'] = df_clean['打包\n人员'].fillna('未分配')  # 填充缺失的打包人员
df_clean['人次'] = df_clean['人次'].fillna(0).astype(int)  # 确保餐数为整数
df_clean.columns = ['地址', '餐数', '打包人员']  # 简化列名

# 按打包人员分组，整理每个人的负责信息
packer_groups = df_clean.groupby('打包人员').agg({
    '地址': list,
    '餐数': list
}).reset_index()

# 生成纯文字内容（极简风格，仅含核心信息）
pure_text_content = f"打包人员地址与餐数分配清单\n"
pure_text_content += f"生成时间：{datetime.now().strftime('%Y年%m月%d日')}\n"
pure_text_content += f"=" * 50 + "\n\n"

# 遍历每个打包人员，生成专属文字块
for _, row in packer_groups.iterrows():
    packer_name = row['打包人员']
    addresses = row['地址']
    meal_counts = row['餐数']

    # 跳过未分配（如需保留可删除此判断）
    if packer_name == '未分配':
        continue

    # 计算个人统计信息
    total_address = len(addresses)
    total_meal = sum(meal_counts)

    # 写入个人信息
    pure_text_content += f"【{packer_name}】\n"
    pure_text_content += f"负责地址总数：{total_address}个 | 总餐数：{total_meal}份\n"
    pure_text_content += f"------------------------\n"

    # 写入每个地址的餐数
    for idx, (addr, meal) in enumerate(zip(addresses, meal_counts), 1):
        pure_text_content += f"{idx}. 地址：{addr} | 餐数：{meal}份\n"

    pure_text_content += "\n"  # 人员之间空行分隔

# 保存为纯文字文件（.txt格式，极简无多余格式）
txt_file_path = r'C:\Users\admin\Desktop\打包人员地址餐数纯文字清单.txt'
with open(txt_file_path, 'w', encoding='utf-8') as f:
    f.write(pure_text_content)

# 同时生成「按人员拆分的纯文字文件」（每个人员1个txt，方便单独发送）
split_dir = '/mnt/纯文字拆分文件'
import os

os.makedirs(split_dir, exist_ok=True)

for _, row in packer_groups.iterrows():
    packer_name = row['打包人员']
    addresses = row['地址']
    meal_counts = row['餐数']

    if packer_name == '未分配':
        continue

    # 个人专属纯文字内容
    personal_text = f"{packer_name} 打包地址与餐数清单\n"
    personal_text += f"生成时间：{datetime.now().strftime('%Y年%m月%d日')}\n"
    personal_text += f"=" * 30 + "\n\n"
    personal_text += f"您负责的地址共 {len(addresses)} 个，总餐数 {sum(meal_counts)} 份\n"
    personal_text += f"------------------------\n"

    for idx, (addr, meal) in enumerate(zip(addresses, meal_counts), 1):
        personal_text += f"{idx}. 地址：{addr}\n"
        personal_text += f"   餐数：{meal}份\n"
        personal_text += "\n"  # 地址之间空行，更易读

    # 保存个人文件
    personal_file_path = f"{split_dir}/{packer_name}_地址餐数清单.txt"
    with open(personal_file_path, 'w', encoding='utf-8') as f:
        f.write(personal_text)

# 输出结果提示
print("✅ 纯文字清单生成完成！")
print("\n1. 汇总版纯文字文件（含所有人员）：")
print(f"   文件路径：{txt_file_path}")
print(f"   内容格式：按人员分组，每行仅含「地址+餐数」核心信息\n")

print("2. 拆分版纯文字文件（每人1个文件，方便单独发送）：")
print(f"   保存目录：{split_dir}")
print(f"   包含人员：共 {len(packer_groups[packer_groups['打包人员'] != '未分配'])} 人")
print(f"   每个文件仅含对应人员的地址和餐数，无其他多余信息\n")

# 预览前300字符（展示纯文字风格）
print("📄 纯文字风格预览（汇总版前300字符）：")
print("-" * 50)
print(pure_text_content[:300] + "...")