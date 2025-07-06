import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from pathlib import Path
from datetime import datetime

from pandas import DataFrame, ExcelFile


class FoodPriceManager:
    def __init__(self, root):
        self.root = root
        self.root.title("菜品价格管理工具")
        self.root.geometry("900x600")
        self.root.minsize(800, 500)

        # 创建主Canvas和Scrollbar
        self.canvas = tk.Canvas(root)
        self.scrollbar = ttk.Scrollbar(root, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 创建Canvas内的Frame
        self.main_frame = ttk.Frame(self.canvas)
        self.canvas_frame = self.canvas.create_window((0, 0), window=self.main_frame, anchor="nw")

        # 配置Canvas滚动区域
        self.main_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        # 创建样式并设置字体
        self.style = ttk.Style()
        self.style.configure('TLabel', font=('SimHei', 10))
        self.style.configure('TButton', font=('SimHei', 10))
        self.style.configure('TCheckbutton', font=('SimHei', 10))
        self.style.configure('Treeview', font=('SimHei', 10))
        self.style.configure('Treeview.Heading', font=('SimHei', 10, 'bold'))

        # 创建文件选择区域
        self.file_frame = ttk.LabelFrame(self.main_frame, text="文件选择", padding="10")
        self.file_frame.pack(fill=tk.X, pady=5, padx=5)

        # 订单文件选择
        ttk.Label(self.file_frame, text="订单文件:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.order_file = tk.StringVar()
        ttk.Entry(self.file_frame, textvariable=self.order_file, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(self.file_frame, text="浏览...", command=lambda: self.browse_file(self.order_file, "订单")).grid(
            row=0, column=2, padx=5, pady=5)

        # 菜品价格文件选择
        ttk.Label(self.file_frame, text="菜品价格文件:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.price_file = tk.StringVar()
        ttk.Entry(self.file_frame, textvariable=self.price_file, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(self.file_frame, text="浏览...", command=lambda: self.browse_file(self.price_file, "菜品价格")).grid(
            row=1, column=2, padx=5, pady=5)

        # 批量处理选项
        self.batch_var = tk.BooleanVar()
        ttk.Checkbutton(self.file_frame, text="批量处理文件夹中的所有订单文件",
                        variable=self.batch_var, command=self.toggle_batch_mode).grid(row=2, column=0, columnspan=3,
                                                                                      sticky=tk.W, pady=5)

        # 文件夹选择（批量模式）
        self.folder_path = tk.StringVar()
        self.folder_entry = ttk.Entry(self.file_frame, textvariable=self.folder_path, width=50, state="disabled")
        self.folder_entry.grid(row=3, column=1, padx=5, pady=5)
        ttk.Button(self.file_frame, text="浏览...", command=self.browse_folder, state="disabled").grid(row=3, column=2,
                                                                                                       padx=5, pady=5)
        ttk.Label(self.file_frame, text="订单文件夹:").grid(row=3, column=0, sticky=tk.W, pady=5)

        # 列设置区域
        self.column_frame = ttk.LabelFrame(self.main_frame, text="列设置", padding="10")
        self.column_frame.pack(fill=tk.X, pady=5, padx=5)

        # 订单文件列设置
        ttk.Label(self.column_frame, text="订单文件:").grid(row=0, column=0, sticky=tk.W, pady=5)

        ttk.Label(self.column_frame, text="菜品名称起始列:").grid(row=1, column=0, sticky=tk.E, pady=2)
        self.order_food_start_col = tk.StringVar(value="10")
        ttk.Entry(self.column_frame, textvariable=self.order_food_start_col, width=5).grid(row=1, column=1, padx=5,
                                                                                           pady=2)

        ttk.Label(self.column_frame, text="菜品名称结束列:").grid(row=1, column=2, sticky=tk.E, pady=2)
        self.order_food_end_col = tk.StringVar(value="-1")
        ttk.Entry(self.column_frame, textvariable=self.order_food_end_col, width=5).grid(row=1, column=3, padx=5,
                                                                                         pady=2)

        # 菜品价格文件列设置
        ttk.Label(self.column_frame, text="菜品价格文件:").grid(row=0, column=4, sticky=tk.W, pady=5)

        ttk.Label(self.column_frame, text="菜品名称列:").grid(row=1, column=4, sticky=tk.E, pady=2)
        self.price_name_col = tk.StringVar(value="名字")
        ttk.Entry(self.column_frame, textvariable=self.price_name_col, width=10).grid(row=1, column=5, padx=5, pady=2)

        ttk.Label(self.column_frame, text="价格列:").grid(row=1, column=6, sticky=tk.E, pady=2)
        self.price_price_col = tk.StringVar(value="价格")
        ttk.Entry(self.column_frame, textvariable=self.price_price_col, width=10).grid(row=1, column=7, padx=5, pady=2)

        # 创建进度区域
        self.progress_frame = ttk.LabelFrame(self.main_frame, text="进度", padding="10")
        self.progress_frame.pack(fill=tk.X, pady=5, padx=5)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.progress_frame, variable=self.progress_var, length=100)
        self.progress_bar.pack(fill=tk.X, padx=5, pady=5)

        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        ttk.Label(self.progress_frame, textvariable=self.status_var).pack(anchor=tk.W, padx=5)

        # 创建结果区域
        self.result_frame = ttk.LabelFrame(self.main_frame, text="检查结果", padding="10")
        self.result_frame.pack(fill=tk.BOTH, expand=True, pady=5, padx=5)

        # 创建结果表格
        columns = ("菜品名称", "来源文件", "状态")
        self.result_tree = ttk.Treeview(self.result_frame, columns=columns, show="headings", height=10)

        for col in columns:
            self.result_tree.heading(col, text=col)
            self.result_tree.column(col, width=300, anchor=tk.CENTER)

        self.result_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 创建统计信息
        self.stats_frame = ttk.Frame(self.result_frame)
        self.stats_frame.pack(fill=tk.X, padx=5, pady=5)

        self.total_foods_var = tk.StringVar()
        self.total_foods_var.set("总检查菜品数: 0")
        ttk.Label(self.stats_frame, textvariable=self.total_foods_var).pack(side=tk.LEFT, padx=10)

        self.missing_foods_var = tk.StringVar()
        self.missing_foods_var.set("缺失价格菜品: 0")
        ttk.Label(self.stats_frame, textvariable=self.missing_foods_var).pack(side=tk.LEFT, padx=10)

        # 创建日志区域
        self.log_frame = ttk.LabelFrame(self.main_frame, text="日志", padding="10")
        self.log_frame.pack(fill=tk.BOTH, expand=True, pady=5, padx=5)

        self.log_text = tk.Text(self.log_frame, wrap=tk.WORD, height=5, font=('SimHei', 10))
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 创建按钮区域
        self.button_frame = ttk.Frame(self.main_frame, padding="10")
        self.button_frame.pack(fill=tk.X, pady=5, padx=5)

        ttk.Button(self.button_frame, text="检查价格", command=self.check_prices).pack(side=tk.RIGHT, padx=5)
        ttk.Button(self.button_frame, text="更新价格文件", command=self.update_price_file).pack(side=tk.RIGHT, padx=5)
        ttk.Button(self.button_frame, text="清空结果", command=self.clear_results).pack(side=tk.RIGHT, padx=5)

        # 绑定鼠标滚轮事件
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        # 绑定关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 初始化
        self.excel_files = []
        self.missing_data = []
        self.price_file_df = None

    def _on_frame_configure(self, event):
        """当Frame大小改变时，更新Canvas的滚动区域"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        """当Canvas大小改变时，调整内部Frame的宽度"""
        width = event.width
        self.canvas.itemconfig(self.canvas_frame, width=width)

    def _on_mousewheel(self, event):
        """处理鼠标滚轮事件"""
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def browse_file(self, var, file_type):
        """浏览并选择文件"""
        file_selected = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
        )
        if file_selected:
            var.set(file_selected)
            self.log(f"已选择{file_type}文件: {os.path.basename(file_selected)}")

    def browse_folder(self):
        """浏览并选择文件夹"""
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.folder_path.set(folder_selected)
            self.log(f"已选择订单文件夹: {os.path.basename(folder_selected)}")

    def toggle_batch_mode(self):
        """切换批量处理模式"""
        if self.batch_var.get():
            self.folder_entry.config(state="normal")
            self.order_file.set("")
            ttk.Entry(self.file_frame, textvariable=self.order_file, width=50, state="disabled").grid(row=0, column=1,
                                                                                                      padx=5, pady=5)
            self.log("已切换到批量处理模式")
        else:
            self.folder_entry.config(state="disabled")
            ttk.Entry(self.file_frame, textvariable=self.order_file, width=50, state="normal").grid(row=0, column=1,
                                                                                                    padx=5, pady=5)
            self.log("已切换到单文件处理模式")

    def log(self, message):
        """添加日志到日志区域"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)

    def clear_results(self):
        """清空结果和日志"""
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        self.log_text.delete(1.0, tk.END)
        self.status_var.set("就绪")
        self.progress_var.set(0)
        self.total_foods_var.set("总检查菜品数: 0")
        self.missing_foods_var.set("缺失价格菜品: 0")
        self.excel_files = []
        self.missing_data = []
        self.price_file_df = None

    def is_xls_file(self, file_path):
        """检查文件是否为 .xls 格式"""
        return file_path.lower().endswith('.xls')

    def is_xlsx_file(self, file_path):
        """检查文件是否为 .xlsx 格式"""
        return file_path.lower().endswith('.xlsx')

    def load_price_file(self) -> tuple[None, None, None] | tuple[DataFrame, ExcelFile, str]:
        """加载菜品价格文件"""
        price_file = self.price_file.get()
        if not price_file:
            messagebox.showerror("错误", "请选择菜品价格文件")
            return None, None, None

        try:
            self.log(f"正在加载菜品价格文件: {os.path.basename(price_file)}")

            # 根据文件类型选择引擎
            engine = 'openpyxl' if self.is_xlsx_file(price_file) else 'xlrd'

            # 读取 Excel 文件
            #excel_file = pd.ExcelFile(price_file)
            excel_file = pd.ExcelFile(r"C:\Users\Administrator\Desktop\菜品贴纸打印.xlsx")
            # 直接获取'Sheet1'工作表的数据
            try:
                price_df = excel_file.parse('Sheet1', engine='openpyxl')
            except Exception as e:
                error_msg = f"错误：无法加载价格文件的'Sheet1'工作表: {str(e)}"
                self.log(error_msg)
                messagebox.showerror("错误", error_msg)
                return None, None, None

            # 检查DataFrame是否为空
            if price_df.empty:
                error_msg = "错误：价格文件的'Sheet1'工作表为空"
                self.log(error_msg)
                messagebox.showerror("错误", error_msg)
                return None, None, None

            # 检查菜品名称列是否存在
            if self.price_name_col.get() not in price_df.columns:
                error_msg = f"错误：在价格文件的'Sheet1'工作表中未找到'{self.price_name_col.get()}'列"
                self.log(error_msg)
                messagebox.showerror("错误", error_msg)
                return None, None, None

            # 检查价格列是否存在
            if self.price_price_col.get() not in price_df.columns:
                error_msg = f"错误：在价格文件的'Sheet1'工作表中未找到'{self.price_price_col.get()}'列"
                self.log(error_msg)
                messagebox.showerror("错误", error_msg)
                return None, None, None

            self.log(f"已加载菜品价格文件，共 {len(price_df)} 条记录")
            return price_df, excel_file, 'Sheet1'

        except Exception as e:
            error_msg = f"加载菜品价格文件时出错: {str(e)}"
            self.log(error_msg)
            messagebox.showerror("错误", error_msg)
            return None, None, None

    def check_prices(self):
        """检查订单中的菜品价格"""
        # 加载价格文件
        result = self.load_price_file()


        self.price_file_df, excel_file_obj, price_sheet = result

        # 获取订单文件列表
        if self.batch_var.get():
            folder_path = self.folder_path.get()
            if not folder_path:
                messagebox.showerror("错误", "请选择订单文件夹")
                return

            # 查找所有 Excel 文件
            self.excel_files = []
            for ext in ['.xlsx', '.xls']:
                self.excel_files.extend(Path(folder_path).rglob(f"*{ext}"))

            if not self.excel_files:
                messagebox.showwarning("警告", f"在文件夹 {folder_path} 中未找到Excel文件")
                self.status_var.set("就绪")
                return

            self.log(f"找到 {len(self.excel_files)} 个订单文件")

        else:
            file_path = self.order_file.get()
            if not file_path:
                messagebox.showerror("错误", "请选择订单文件")
                return

            if not (self.is_xls_file(file_path) or self.is_xlsx_file(file_path)):
                messagebox.showerror("错误", "请选择有效的Excel文件 (.xlsx 或 .xls)")
                return

            self.excel_files = [file_path]

        # 清空结果
        self.clear_results()

        # 检查价格
        self.status_var.set("正在检查价格...")
        self.root.update_idletasks()

        total_foods = 0
        missing_count = 0

        # 获取价格文件中的菜品名称列表
        if self.price_file_df is None:
            self.log("错误：价格文件数据为空")
            self.status_var.set("就绪")
            return

        all_food_names = self.price_file_df[self.price_name_col.get()].tolist()

        file_count = len(self.excel_files)
        for i, file_path in enumerate(self.excel_files):
            self.progress_var.set((i / file_count) * 100)
            self.status_var.set(f"正在处理文件 {i + 1}/{file_count}: {os.path.basename(file_path)}")
            self.root.update_idletasks()

            # 检查单个文件
            try:
                # 根据文件类型选择引擎
                engine = 'openpyxl' if self.is_xlsx_file(file_path) else 'xlrd'

                # 读取 Excel 文件
                order_excel_file = pd.ExcelFile(file_path)

                # 获取'Sheet1'工作表的数据
                df_order = order_excel_file.parse('Sheet1', engine=engine)

                if df_order.empty:
                    self.log(f"警告：订单文件 '{os.path.basename(file_path)}' 的'Sheet1'工作表为空")
                    continue

                # 获取菜品名称列范围
                try:
                    start_col = int(self.order_food_start_col.get())
                    end_col = int(self.order_food_end_col.get())
                except ValueError:
                    self.log(f"错误：列号必须为整数")
                    continue

                # 验证列索引是否有效
                if start_col < 0 or start_col >= len(df_order.columns):
                    self.log(f"错误：订单文件的列索引超出范围（总列数：{len(df_order.columns)}）")
                    continue

                if end_col != -1 and (end_col <= start_col or end_col > len(df_order.columns)):
                    self.log(f"错误：订单文件的结束列索引无效（总列数：{len(df_order.columns)}）")
                    continue

                # 提取订餐内容相关列
                if end_col == -1:
                    food_columns = df_order.columns[start_col:]
                else:
                    food_columns = df_order.columns[start_col:end_col]

                if not food_columns.empty:
                    # 获取订餐内容列的菜品名称
                    order_food_names = food_columns.tolist()

                    # 找出在订单中但不在价格文件中的菜品
                    missing_foods = [food for food in order_food_names if food not in all_food_names]

                    total_foods += len(order_food_names)
                    missing_count += len(missing_foods)

                    # 记录缺失的菜品
                    for food in missing_foods:
                        self.missing_data.append({
                            "菜品名称": food,
                            "来源文件": os.path.basename(file_path),
                            "状态": "价格缺失"
                        })

                    if missing_foods:
                        self.log(f"文件 {os.path.basename(file_path)} 中发现 {len(missing_foods)} 个价格缺失的菜品")
                    else:
                        self.log(f"文件 {os.path.basename(file_path)} 中的所有菜品均有价格记录")
                else:
                    self.log(f"警告：订单文件 '{os.path.basename(file_path)}' 中未找到菜品列")

            except Exception as e:
                self.log(f"处理文件 {os.path.basename(file_path)} 时发生错误: {str(e)}")

        # 显示结果
        self.display_results()

        # 更新统计信息
        self.total_foods_var.set(f"总检查菜品数: {total_foods}")
        self.missing_foods_var.set(f"缺失价格菜品: {missing_count}")

        self.status_var.set(f"完成! 共检查 {total_foods} 个菜品，发现 {missing_count} 个价格缺失")
        self.progress_var.set(100)

        # 滚动到底部，确保按钮可见
        self.canvas.yview_moveto(1.0)

        if missing_count > 0:
            messagebox.showinfo("检查完成",
                                f"共检查 {total_foods} 个菜品，发现 {missing_count} 个价格缺失\n请查看结果表格获取详细信息")
        else:
            messagebox.showinfo("检查完成", f"恭喜！所有菜品均有价格记录")

    def display_results(self):
        """显示检查结果"""
        # 清空现有结果
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        # 显示缺失数据
        for data in self.missing_data:
            self.result_tree.insert("", tk.END, values=(
                data["菜品名称"],
                data["来源文件"],
                data["状态"]
            ))

    def update_price_file(self):
        """更新菜品价格文件，添加缺失的菜品"""
        if not self.missing_data:
            messagebox.showinfo("提示", "没有需要添加的菜品")
            return

        if self.price_file_df is None:
            messagebox.showerror("错误", "请先检查价格，加载菜品价格文件")
            return

        # 确认操作
        if not messagebox.askyesno("确认", f"确定要将 {len(self.missing_data)} 个缺失的菜品添加到价格文件中吗？"):
            return

        try:
            # 获取当前日期
            current_date = datetime.now().strftime('%Y-%m-%d')

            # 提取缺失的菜品名称
            missing_food_names = [item["菜品名称"] for item in self.missing_data]

            # 创建新的菜品 DataFrame
            new_foods_df = pd.DataFrame({
                self.price_name_col.get(): missing_food_names,
                self.price_price_col.get(): [None] * len(missing_food_names),
                '添加时间': [current_date] * len(missing_food_names)
            })

            # 将新菜品添加到价格 DataFrame
            self.price_file_df = pd.concat([self.price_file_df, new_foods_df], ignore_index=True)

            # 选择保存位置
            price_file_path = self.price_file.get()
            default_dir = os.path.dirname(price_file_path)
            default_name = os.path.splitext(os.path.basename(price_file_path))[0] + "-更新.xlsx"

            save_path = filedialog.asksaveasfilename(
                initialdir=default_dir,
                initialfile=default_name,
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )

            if not save_path:
                return

            # 创建 ExcelWriter 对象
            with pd.ExcelWriter(save_path, engine='openpyxl', mode='w') as writer:
                # 获取价格文件的所有表名
                price_file = self.price_file.get()
                engine = 'openpyxl' if self.is_xlsx_file(price_file) else 'xlrd'
                excel_file = pd.ExcelFile(price_file)
                sheet_names = excel_file.sheet_names

                # 遍历每个工作表
                for sheet_name in sheet_names:
                    if sheet_name == 'Sheet1':
                        # 写入更新后的价格数据
                        self.price_file_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    else:
                        # 其他工作表保持不变
                        df = excel_file.parse(sheet_name, engine=engine)
                        df.to_excel(writer, sheet_name=sheet_name, index=False)

            self.log(f'已将 {len(missing_food_names)} 个缺失的菜品添加到"Sheet1"工作表中，请在价格列中填写价格。')
            self.log(f'更新后的文件已保存至: {save_path}')
            messagebox.showinfo("成功",
                                f'已将 {len(missing_food_names)} 个缺失的菜品添加到价格文件中\n请在价格列中填写价格。\n更新后的文件已保存至: {save_path}')

        except Exception as e:
            self.log(f"更新价格文件时出错: {str(e)}")
            messagebox.showerror("错误", f"更新价格文件时出错: {str(e)}")

    def on_closing(self):
        """关闭窗口时确认"""
        if messagebox.askokcancel("退出", "确定要退出吗?"):
            self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = FoodPriceManager(root)
    root.mainloop()