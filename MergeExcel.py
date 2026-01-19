import pandas as pd
import os


def merge_excel_with_dividers(folder_path, output_file):
    """
    合并指定文件夹中所有Excel文件的Sheet1工作表，
    并在每个文件的数据前添加文件名作为分界线

    参数:
        folder_path: 包含Excel文件的文件夹路径
        output_file: 合并后的输出文件路径
    """
    # 获取文件夹中所有.xlsx和.xls文件
    excel_files = []
    for file in os.listdir(folder_path):
        if file.endswith('.xlsx') and not file.startswith('~$'):  # 排除临时文件
            excel_files.append(os.path.join(folder_path, file))

    # 初始化一个空DataFrame用于存储所有数据
    all_data = pd.DataFrame()

    # 遍历所有Excel文件并读取Sheet1
    for i, file in enumerate(excel_files):
        try:
            # 读取当前文件的Sheet1
            df = pd.read_excel(file, sheet_name='Sheet1', index_col=None)
            file_name = os.path.basename(file)

            # 创建一个分界线行，只在第一个文件前不加分界线
            if i > 0:
                divider = pd.DataFrame([[f"----- 以下数据来自: {file_name} -----"] + [None] * (len(df.columns) - 1)],
                                       columns=df.columns)
                all_data = pd.concat([all_data, divider], ignore_index=True)

            # 添加来源文件列（可选）
            df['来源文件'] = file_name

            # 合并数据
            all_data = pd.concat([all_data, df], ignore_index=True)
            print(f"已处理: {file_name}")
        except Exception as e:
            print(f"处理{os.path.basename(file)}时出错: {str(e)}")

    # 将合并后的数据保存到新的Excel文件
    all_data.to_excel(output_file, index=False)
    print(f"合并完成，已保存至: {output_file}")


# 使用示例
if __name__ == "__main__":
    # 替换为你的Excel文件所在文件夹路径
    input_folder = r"D:\WPS云盘\1214901082\WPS云盘\工作\沈飞\订单数据\11月午餐\新建文件夹"
    # 替换为输出文件路径
    output_file = "1月账单查询文档.xlsx"

    merge_excel_with_dividers(input_folder, output_file)
