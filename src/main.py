# Copyright 2023 agenius666
# GitHub: https://github.com/agenius666/Audit-Inspector-Toolkit
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
import time
import chardet
import pyarrow
import webbrowser
import requests
import threading
import sqlite3


class ExcelLikeApp:
    def __init__(self, root):
        """
        初始化应用程序
        :param root: Tkinter 根窗口
        """
        self.root = root
        self.root.title("Audit Inspector Toolkit - 1.1.0")

        # 中英文表名映射
        self.table_name_mapping = {
            "序时账": "journal",
            "科目余额表": "balance",
            "凭证": "voucher"
        }

        # 定义数据保存路径
        self.data_dir = "saved_data"
        if not os.path.exists(self.data_dir):
            os.makedirs(self.data_dir)

        # 初始化数据库
        self.db_path = os.path.join(self.data_dir, "data.db")
        self.init_db()

        # 初始化当前表格名称
        self.current_sheet_name = None

        # 初始化每个表的筛选状态
        self.filter_states = {
            "序时账": {
                "filter_history": [],  # 筛选历史记录
                "filtered_data_cache": None,  # 当前筛选结果缓存
                "filter_entries": {}  # 筛选框输入框状态
            },
            "科目余额表": {
                "filter_history": [],
                "filtered_data_cache": None,
                "filter_entries": {}
            },
            "凭证": {
                "filter_history": [],
                "filtered_data_cache": None,
                "filter_entries": {}
            }
        }

        # 初始化空数据
        self.sheets = {
            "序时账": pd.DataFrame(
                columns=["日期", "凭证字号", "科目编码", "科目名称", "辅助核算", "摘要", "借方", "贷方", "数量", "外币"]
            ),
            "凭证": pd.DataFrame(columns=["日期", "凭证字号", "摘要", "科目名称", "借方", "贷方", "数量/外币"]),
            "科目余额表": pd.DataFrame(
                columns=["科目编码", "科目名称", "期初借方余额", "期初贷方余额", "本期借方发生额", "本期贷方发生额",
                         "期末借方余额", "期末贷方余额"]
            ),
        }

        # 初始化 trees 和 filter_frames
        self.trees = {}
        self.filter_frames = {}
        self.filter_entries = {}

        # 初始化筛选历史记录
        self.filter_history = []
        self.filtered_data_cache = None

        # 创建按钮栏
        self.create_buttons()

        # 创建每个sheet的界面
        self.create_sheets_ui()

        # 初始加载数据
        self.load_from_db("序时账", limit=100, offset=0)  # 加载前 100 行序时账
        self.load_from_db("科目余额表")  # 全量加载科目余额表

        # 绑定快捷键
        self.root.bind(
            "<Control-c>",
            lambda event: self.copy_selection(self.trees[self.notebook.tab(self.notebook.select(), "text")]),
        )
        self.root.bind(
            "<Control-v>",
            lambda event: self.paste_selection(self.trees[self.notebook.tab(self.notebook.select(), "text")]),
        )

    def copy_selection(self, tree):
        """
        复制选中的数据到剪贴板
        :param tree: Treeview 控件
        """
        selected_items = tree.selection()
        if selected_items:
            # 获取所有选中行的数据
            copied_data = []
            for item in selected_items:
                values = tree.item(item, 'values')
                copied_data.append("\t".join(map(str, values)))

            # 将多行数据复制到剪贴板
            self.root.clipboard_clear()
            self.root.clipboard_append("\n".join(copied_data))

    def paste_selection(self, tree):
        """
        将剪贴板中的数据粘贴到 Treeview 中
        :param tree: Treeview 控件
        """
        try:
            clipboard_data = self.root.clipboard_get()
            # 按行分割剪贴板中的数据
            rows = clipboard_data.split("\n")
            for row in rows:
                values = row.split("\t")
                tree.insert("", "end", values=values)
        except tk.TclError:
            messagebox.showerror("错误", "剪贴板中没有数据或数据格式不正确！")

    def create_buttons(self):
        """
        创建按钮栏
        """
        # 创建上传按钮的容器
        upload_frame = ttk.Frame(self.root)
        upload_frame.pack(side=tk.TOP, fill=tk.X, pady=5)

        # 添加上传序时账按钮
        upload_journal_button = ttk.Button(upload_frame, text="上传序时账", command=lambda: self.upload_file("序时账"))
        upload_journal_button.pack(side=tk.LEFT, padx=5)

        # 添加上传科目余额表按钮
        upload_balance_button = ttk.Button(upload_frame, text="上传科目余额表", command=lambda: self.upload_file("科目余额表"))
        upload_balance_button.pack(side=tk.LEFT, padx=5)

        # 添加上传数据库按钮
        upload_db_button = ttk.Button(upload_frame, text="上传数据库", command=self.upload_db)
        upload_db_button.pack(side=tk.LEFT, padx=5)

        # 添加分隔线
        separator = ttk.Separator(upload_frame, orient="vertical")
        separator.pack(side=tk.LEFT, padx=5, fill=tk.Y)

        # 添加清空按钮
        clear_journal_button = ttk.Button(upload_frame, text="清空序时账", command=lambda: self.clear_sheet("序时账"))
        clear_journal_button.pack(side=tk.LEFT, padx=5)

        clear_balance_button = ttk.Button(upload_frame, text="清空科目余额表", command=lambda: self.clear_sheet("科目余额表"))
        clear_balance_button.pack(side=tk.LEFT, padx=5)

        # 添加分隔线
        separator = ttk.Separator(upload_frame, orient="vertical")
        separator.pack(side=tk.LEFT, padx=5, fill=tk.Y)

        # 恢复原始序时账按钮
        restore_journal_button = ttk.Button(upload_frame, text="恢复序时账", command=self.restore_journal)
        restore_journal_button.pack(side=tk.LEFT, padx=5)

        # 恢复科目余额表按钮
        restore_balance_button = ttk.Button(upload_frame, text="恢复科目余额表", command=self.restore_balance)
        restore_balance_button.pack(side=tk.LEFT, padx=5)

        # 添加分隔线
        separator = ttk.Separator(upload_frame, orient="vertical")
        separator.pack(side=tk.LEFT, padx=5, fill=tk.Y)

        # 添加保存按钮
        save_journal_button = ttk.Button(upload_frame, text="保存序时账", command=lambda: self.save_sheet("序时账"))
        save_journal_button.pack(side=tk.LEFT, padx=5)

        save_balance_button = ttk.Button(upload_frame, text="保存科目余额表", command=lambda: self.save_sheet("科目余额表"))
        save_balance_button.pack(side=tk.LEFT, padx=5)

        save_db_button = ttk.Button(upload_frame, text="保存数据库", command=self.save_to_db_from_ui)
        save_db_button.pack(side=tk.LEFT, padx=5)

        # 新增按钮（放到最右边）
        data_validation_button = ttk.Button(upload_frame, text="数据校验", command=self.data_validation)
        data_validation_button.pack(side=tk.RIGHT, padx=5)

        # 创建第二排按钮的容器
        second_row_frame = ttk.Frame(self.root)
        second_row_frame.pack(side=tk.TOP, fill=tk.X, pady=5)

        # 将恢复筛选和清空筛选按钮放到第二排的最右边
        clear_filter_button = ttk.Button(second_row_frame, text="清空筛选", command=self.clear_filter)
        clear_filter_button.pack(side=tk.RIGHT, padx=5)

        restore_filter_button = ttk.Button(second_row_frame, text="恢复筛选", command=self.restore_last_filter)
        restore_filter_button.pack(side=tk.RIGHT, padx=5)

    def create_sheets_ui(self):
        """
        创建每个sheet的界面，并为序时账的Treeview绑定滚动事件
        """
        # 创建Notebook（选项卡）
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # 创建每个sheet的界面
        for sheet_name, df in self.sheets.items():
            # 创建选项卡的容器
            tab_frame = ttk.Frame(self.notebook)
            self.notebook.add(tab_frame, text=sheet_name)

            # 绑定选项卡切换事件
            self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

            # 创建筛选框的容器（放在选项卡的下方）
            filter_frame = ttk.Frame(tab_frame)
            filter_frame.pack(fill=tk.X, pady=5)
            self.filter_frames[sheet_name] = filter_frame

            # 创建Treeview控件
            tree = ttk.Treeview(tab_frame, columns=list(df.columns), show="headings")
            tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

            # 添加垂直滚动条
            scrollbar = ttk.Scrollbar(tab_frame, orient=tk.VERTICAL, command=tree.yview)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            tree.configure(yscrollcommand=scrollbar.set)

            # 设置列标题
            for col in df.columns:
                tree.heading(col, text=col)
                tree.column(col, width=100)

            # 保存Treeview控件以便后续更新
            self.trees[sheet_name] = tree

            # 绑定右键单击事件（仅对科目余额表和序时账）
            if sheet_name == "科目余额表":
                tree.bind("<Button-3>", self.show_detail_journal)  # 右键单击事件
            elif sheet_name == "序时账":
                tree.bind("<Button-3>", self.show_voucher_details)  # 右键单击事件

                # 绑定滚动事件（仅对序时账）
                def on_scroll(event, tree=tree):
                    """
                    当用户滚动到底部时加载下一部分数据
                    """
                    # 检查是否滚动到底部
                    if tree.yview()[1] == 1.0:  # 滚动条到底部
                        # 计算当前已加载的行数
                        offset = len(self.sheets["序时账"])
                        # 加载下一部分数据
                        self.load_from_db("序时账", limit=100, offset=offset)

                # 绑定滚动事件
                tree.bind("<MouseWheel>", on_scroll)  # Windows 和 macOS
                tree.bind("<Button-4>", on_scroll)  # Linux（向上滚动）
                tree.bind("<Button-5>", on_scroll)  # Linux（向下滚动）

            elif sheet_name == "凭证":
                tree.bind("<Button-3>", lambda event: self.notebook.select(0))  # 右键切换到序时账选项卡

            # 初始化筛选框
            self.update_filter_entries(sheet_name)

    def on_tab_changed(self, event):
        """
        选项卡切换事件处理
        """
        # 获取当前选中的选项卡名称
        selected_tab_index = self.notebook.index(self.notebook.select())
        self.current_sheet_name = self.notebook.tab(selected_tab_index, "text")

    def create_progress_window(self, title):
        """
        创建进度条窗口
        :param title: 窗口标题
        :return: 返回进度条窗口、进度条和计时器标签
        """
        # 创建进度条窗口
        progress_window = tk.Toplevel(self.root)
        progress_window.title(title)
        progress_window.geometry("300x100")

        # 添加进度条
        progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=250, mode="determinate")
        progress_bar.pack(pady=10)

        # 添加计时器标签
        timer_label = ttk.Label(progress_window, text="耗时: 0.00 秒")
        timer_label.pack(pady=5)

        return progress_window, progress_bar, timer_label

    def update_progress(self, progress_window, progress_bar, timer_label, start_time, current, total):
        """
        更新进度条
        :param progress_window: 进度条窗口
        :param progress_bar: 进度条
        :param timer_label: 计时器标签
        :param start_time: 开始时间
        :param current: 当前进度
        :param total: 总进度
        """
        # 更新进度条
        progress_bar["value"] = (current + 1) / total * 100
        timer_label.config(text=f"耗时: {time.time() - start_time:.2f} 秒")
        progress_window.update()
        time.sleep(0.01)  # 模拟处理延迟

    def init_db(self):
        """
        初始化数据库，仅创建数据表（如果表不存在）
        """
        # 如果数据库文件已存在，先删除它
        if os.path.exists(self.db_path):
            os.remove(self.db_path)

        # 创建新的数据库文件
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()

            # 创建序时账表
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS journal (
                    日期 TEXT,
                    凭证字号 TEXT,
                    科目编码 TEXT,
                    科目名称 TEXT,
                    辅助核算 TEXT,
                    摘要 TEXT,
                    借方 REAL,
                    贷方 REAL,
                    数量 REAL,
                    外币 REAL
                )
            ''')

            # 创建科目余额表
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS balance (
                    科目编码 TEXT,
                    科目名称 TEXT,
                    期初借方余额 REAL,
                    期初贷方余额 REAL,
                    本期借方发生额 REAL,
                    本期贷方发生额 REAL,
                    期末借方余额 REAL,
                    期末贷方余额 REAL
                )
            ''')

            conn.commit()

    def upload_file(self, sheet_name):
        """
        上传文件并写入初始化的数据库（data.db），显示进度条
        """
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("Parquet files", "*.parquet")]
        )
        if not file_path:
            return

        # 创建进度条窗口
        progress_window, progress_bar, timer_label = self.create_progress_window("上传文件")

        try:
            # 模拟上传过程
            total_steps = 10
            start_time = time.time()
            for i in range(total_steps):
                self.update_progress(progress_window, progress_bar, timer_label, start_time, i, total_steps)
                time.sleep(0.1)  # 模拟上传延迟

            # 读取文件
            if file_path.endswith('.csv'):
                with open(file_path, 'rb') as f:
                    raw_data = f.read()
                    result = chardet.detect(raw_data)
                    encoding = result['encoding']
                df = pd.read_csv(file_path, encoding=encoding)
            elif file_path.endswith('.parquet'):
                df = pd.read_parquet(file_path)
            else:
                df = pd.read_excel(file_path)

            # 将数据写入初始化的数据库（data.db）
            with sqlite3.connect(self.db_path) as conn:
                table_name = self.table_name_mapping.get(sheet_name)
                if not table_name:
                    messagebox.showerror("错误", f"不支持的文件类型：{sheet_name}")
                    return

                # 写入数据
                df.to_sql(table_name, conn, if_exists='replace', index=False)

                # 为每一列创建索引
                cursor = conn.cursor()
                for col in df.columns:
                    index_name = f"idx_{table_name}_{col}"
                    cursor.execute(f'CREATE INDEX IF NOT EXISTS {index_name} ON {table_name} ({col})')
                conn.commit()

            # 更新 Treeview
            if sheet_name == "序时账":
                self.load_from_db("序时账", limit=100, offset=0)  # 加载前 100 行
            else:
                self.load_from_db(sheet_name)  # 全量加载

            messagebox.showinfo("成功", f"{sheet_name}上传完成！")

        except Exception as e:
            messagebox.showerror("错误", f"上传{sheet_name}时出错: {e}")
        finally:
            progress_window.destroy()  # 关闭进度条窗口

    def upload_db(self):
        """
        上传指定路径的数据库文件，并检查每一列的索引，显示进度条
        """
        file_path = filedialog.askopenfilename(filetypes=[("SQLite files", "*.db")])
        if not file_path:
            return

        # 创建进度条窗口
        progress_window, progress_bar, timer_label = self.create_progress_window("上传数据库")

        try:
            # 模拟上传过程
            total_steps = 10
            start_time = time.time()
            for i in range(total_steps):
                self.update_progress(progress_window, progress_bar, timer_label, start_time, i, total_steps)
                time.sleep(0.1)  # 模拟上传延迟

            # 检查指定路径的数据库文件
            with sqlite3.connect(file_path) as conn:
                cursor = conn.cursor()

                # 检查每一列的索引
                for table_name in ["journal", "balance"]:
                    cursor.execute(f"PRAGMA table_info({table_name})")
                    columns = [row[1] for row in cursor.fetchall()]  # 获取列名
                    for col in columns:
                        index_name = f"idx_{table_name}_{col}"
                        cursor.execute(f'CREATE INDEX IF NOT EXISTS {index_name} ON {table_name} ({col})')
                conn.commit()

            # 将指定路径的数据库文件设置为当前数据库
            self.db_path = file_path

            # 更新 Treeview
            self.load_from_db("序时账", limit=100, offset=0)  # 分页加载序时账（前 100 行）
            self.load_from_db("科目余额表")  # 全量加载科目余额表

            messagebox.showinfo("成功", "数据库上传完成！")

        except Exception as e:
            messagebox.showerror("错误", f"上传数据库时出错: {e}")
        finally:
            progress_window.destroy()  # 关闭进度条窗口

    def load_from_db(self, sheet_name, limit=None, offset=None):
        """
        从当前数据库加载数据
        :param sheet_name: 表名（如 "序时账" 或 "科目余额表"）
        :param limit: 每次加载的行数（仅对序时账有效）
        :param offset: 起始行数（仅对序时账有效）
        :return: 返回一个 DataFrame，即使为空也不会返回 None
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                table_name = self.table_name_mapping.get(sheet_name)
                if not table_name:
                    messagebox.showerror("错误", f"未找到表名映射：{sheet_name}")
                    return pd.DataFrame()  # 返回空 DataFrame 而不是 None

                query = f'SELECT * FROM {table_name}'

                # 如果是序时账，分页加载
                if sheet_name == "序时账" and limit is not None and offset is not None:
                    query += f' LIMIT {limit} OFFSET {offset}'

                df = pd.read_sql(query, conn)

                # 更新当前表格数据
                if sheet_name == "序时账":
                    if offset == 0:  # 第一次加载
                        self.sheets[sheet_name] = df
                    else:  # 追加数据
                        self.sheets[sheet_name] = pd.concat([self.sheets[sheet_name], df], ignore_index=True)
                else:  # 科目余额表，全量加载
                    self.sheets[sheet_name] = df

                # 更新 Treeview
                tree = self.trees[sheet_name]
                if sheet_name == "序时账" and offset == 0:  # 第一次加载时清空 Treeview
                    for item in tree.get_children():
                        tree.delete(item)
                for _, row in df.iterrows():
                    tree.insert("", "end", values=list(row))

                return df  # 确保返回 DataFrame

        except Exception as e:
            messagebox.showerror("错误", f"从数据库加载数据时出错: {e}")
            return pd.DataFrame()  # 返回空 DataFrame 而不是 None

    def clear_sheet(self, sheet_name):
        """
        清空指定表的数据
        :param sheet_name: 表名
        """
        # 创建进度条窗口
        progress_window, progress_bar, timer_label = self.create_progress_window(f"清空 {sheet_name}")

        try:
            # 模拟清空过程
            total_steps = 10
            start_time = time.time()
            for i in range(total_steps):
                self.update_progress(progress_window, progress_bar, timer_label, start_time, i, total_steps)

            # 清空数据
            self.sheets[sheet_name] = pd.DataFrame(columns=self.sheets[sheet_name].columns)

            # 清空Treeview
            tree = self.trees[sheet_name]
            for item in tree.get_children():
                tree.delete(item)

            # 更新筛选框
            self.update_filter_entries(sheet_name)

        except Exception as e:
            messagebox.showerror("错误", f"清空 {sheet_name} 时出错: {e}")
        finally:
            progress_window.destroy()

    def restore_journal(self, limit=100, offset=0):
        """
        恢复序时账数据（分页加载）
        :param limit: 每次加载的行数
        :param offset: 起始行数
        """
        try:
            # 从数据库中加载原始序时账数据（分页加载）
            original_journal = self.load_from_db("序时账", limit=limit, offset=offset)

            # 检查返回值是否为空 DataFrame
            if original_journal.empty:
                if offset == 0:  # 第一次加载时提示
                    messagebox.showwarning("警告", "未找到序时账数据！")
                return  # 直接返回，避免后续逻辑执行

            # 更新筛选框
            self.update_filter_entries("序时账")

            # 切换到序时账选项卡
            self.notebook.select(self.trees["序时账"].master)

        except Exception as e:
            messagebox.showerror("错误", f"恢复序时账时出错: {e}")

    def restore_balance(self):
        """
        恢复科目余额表数据（全量加载）
        """
        try:
            # 从数据库中加载原始科目余额表数据（全量加载）
            original_balance = self.load_from_db("科目余额表")

            # 检查返回值是否为空 DataFrame
            if original_balance.empty:
                messagebox.showwarning("警告", "未找到科目余额表数据！")
                return

            # 更新筛选框
            self.update_filter_entries("科目余额表")

            # 切换到科目余额表选项卡
            self.notebook.select(self.trees["科目余额表"].master)

        except Exception as e:
            messagebox.showerror("错误", f"恢复科目余额表时出错: {e}")

    def save_sheet(self, sheet_name):
        """
        将数据库中的表保存为指定格式的文件
        :param sheet_name: 表名（如 "序时账" 或 "科目余额表"）
        """
        # 弹出文件保存对话框
        file_path = filedialog.asksaveasfilename(
            defaultextension=".parquet",
            filetypes=[("Parquet files", "*.parquet"), ("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
            title=f"保存 {sheet_name}"
        )
        if not file_path:
            return  # 用户取消选择

        # 创建进度条窗口
        progress_window, progress_bar, timer_label = self.create_progress_window(f"保存 {sheet_name}")

        try:
            # 获取数据库中的表名
            table_name = self.table_name_mapping.get(sheet_name)
            if not table_name:
                messagebox.showerror("错误", f"未找到表名映射：{sheet_name}")
                return

            # 模拟保存过程
            total_steps = 10
            start_time = time.time()
            for i in range(total_steps):
                self.update_progress(progress_window, progress_bar, timer_label, start_time, i, total_steps)

            # 从数据库中读取数据并保存为指定格式
            with sqlite3.connect(self.db_path) as conn:
                df = pd.read_sql(f"SELECT * FROM {table_name}", conn)

                if file_path.endswith('.parquet'):
                    df.to_parquet(file_path, index=False)
                elif file_path.endswith('.xlsx'):
                    df.to_excel(file_path, index=False)
                elif file_path.endswith('.csv'):
                    df.to_csv(file_path, index=False, encoding='utf-8-sig')

            messagebox.showinfo("成功", f"{sheet_name} 已保存！")

        except Exception as e:
            messagebox.showerror("错误", f"保存 {sheet_name} 时出错: {e}")
        finally:
            progress_window.destroy()

    def save_to_db_from_ui(self):
        """
        将当前的数据库文件保存到用户指定的位置
        """
        # 弹出文件保存对话框
        file_path = filedialog.asksaveasfilename(
            defaultextension=".db",
            filetypes=[("SQLite files", "*.db")],
            title="保存数据库"
        )
        if not file_path:
            return  # 用户取消选择

        # 创建进度条窗口
        progress_window, progress_bar, timer_label = self.create_progress_window("保存数据库")

        try:
            # 模拟保存过程
            total_steps = 10
            start_time = time.time()
            for i in range(total_steps):
                self.update_progress(progress_window, progress_bar, timer_label, start_time, i, total_steps)

            # 复制当前的数据库文件到用户指定的位置
            import shutil
            shutil.copy(self.db_path, file_path)

            messagebox.showinfo("成功", f"数据库已保存至：{file_path}")

        except Exception as e:
            messagebox.showerror("错误", f"保存数据库时出错: {e}")
        finally:
            progress_window.destroy()

    def data_validation(self):
        """
        数据校验：检查科目余额表与序时账的金额是否一致
        """
        # 创建进度条窗口
        progress_window, progress_bar, timer_label = self.create_progress_window("数据校验")

        # 在单独的线程中执行数据校验
        def validate_data():
            try:
                with sqlite3.connect(self.db_path) as conn:
                    cursor = conn.cursor()

                    # 检查数据库中是否存在序时账和科目余额表
                    cursor.execute(
                        "SELECT name FROM sqlite_master WHERE type='table' AND name IN ('journal', 'balance')")
                    tables = cursor.fetchall()
                    if len(tables) != 2:
                        messagebox.showerror("错误", "数据库中没有找到序时账或科目余额表！")
                        return

                    # 创建临时表存储序时账的汇总数据
                    cursor.execute('''
                        CREATE TEMPORARY TABLE temp_journal_summary AS
                        SELECT 科目编码, SUM(借方) AS 序时账借方金额, SUM(贷方) AS 序时账贷方金额
                        FROM journal
                        GROUP BY 科目编码
                    ''')

                    # 为临时表创建索引以加快查询速度
                    cursor.execute('CREATE INDEX idx_temp_journal_subject_code ON temp_journal_summary (科目编码)')

                    # 读取科目余额表数据
                    cursor.execute('''
                        SELECT 科目编码, 科目名称, 本期借方发生额, 本期贷方发生额
                        FROM balance
                    ''')
                    balance_data = cursor.fetchall()

                    # 计算校验值
                    discrepancies = []
                    for row in balance_data:
                        subject_code = row[0]
                        subject_name = row[1]
                        debit_balance = row[2]
                        credit_balance = row[3]

                        # 从临时表中获取序时账的金额
                        cursor.execute('''
                            SELECT 序时账借方金额, 序时账贷方金额
                            FROM temp_journal_summary
                            WHERE 科目编码 = ?
                        ''', (subject_code,))
                        journal_row = cursor.fetchone()
                        journal_debit, journal_credit = journal_row if journal_row else (0, 0)

                        # 计算差异
                        debit_diff = debit_balance - journal_debit
                        credit_diff = credit_balance - journal_credit
                        total_diff = debit_diff + credit_diff

                        if abs(total_diff) > 1e-3:  # 浮点数精度问题
                            discrepancies.append((subject_code, subject_name, total_diff))

                    # 检查是否有差异
                    if not discrepancies:
                        messagebox.showinfo("成功", "科目余额表与序时账金额核对一致！")
                    else:
                        # 显示差异信息
                        discrepancy_message = "以下科目余额表与序时账金额不一致：\n\n"
                        for subject_code, subject_name, diff in discrepancies:
                            discrepancy_message += f"科目编码: {subject_code}, 科目名称: {subject_name}, 差异金额: {diff:.2f}\n"
                        messagebox.showerror("错误", discrepancy_message)

                    # 删除临时表
                    cursor.execute("DROP TABLE temp_journal_summary")

            except Exception as e:
                messagebox.showerror("错误", f"数据校验时出错: {e}")
            finally:
                # 关闭进度条窗口
                progress_window.destroy()

        # 启动线程
        threading.Thread(target=validate_data).start()

    def update_treeview(self, tree, df):
        """
        更新 Treeview 中的数据
        :param tree: Treeview 控件
        :param df: 要显示的数据（Pandas DataFrame）
        """
        # 清空 Treeview
        for item in tree.get_children():
            tree.delete(item)

        # 插入数据
        for _, row in df.iterrows():
            tree.insert("", "end", values=list(row))

    def apply_filter_from_entry(self, tree, col, entry):
        """
        根据输入框的内容筛选数据
        :param tree: Treeview 控件
        :param col: 列名
        :param entry: 输入框
        """
        # 获取当前表格的中文名称
        sheet_name = self.notebook.tab(self.notebook.select(), "text")

        # 获取输入框的内容
        filter_text = entry.get().strip()

        if not filter_text:
            messagebox.showwarning("警告", "请输入筛选条件！")
            return

        # 检查当前表的筛选缓存
        if (
                self.filter_states[sheet_name]["filtered_data_cache"] is None
                or self.filter_states[sheet_name]["filtered_data_cache"].empty
        ):
            # 如果是第一次筛选，从数据库中加载数据
            self.apply_first_filter(tree, col, filter_text)
        else:
            # 否则，在内存中进行筛选
            self.apply_nth_filter(tree, col, filter_text)

    def restore_last_filter(self):
        """
        恢复上一次筛选
        """
        try:
            # 获取当前表格的中文名称
            sheet_name = self.notebook.tab(self.notebook.select(), "text")

            # 检查当前表的筛选历史记录
            if len(self.filter_states[sheet_name]["filter_history"]) <= 1:
                messagebox.showwarning("警告", "没有可恢复的筛选记录！")
                return

            # 恢复到上一次的筛选结果
            self.filter_states[sheet_name]["filtered_data_cache"] = self.filter_states[sheet_name]["filter_history"][
                -2]  # 使用倒数第二条记录
            self.filter_states[sheet_name]["filter_history"].pop()  # 移除当前筛选结果

            # 更新 Treeview
            self.update_treeview(self.trees[sheet_name], self.filter_states[sheet_name]["filtered_data_cache"])

        except Exception as e:
            messagebox.showerror("错误", f"恢复上一次筛选时出错: {e}")

    def clear_filter(self):
        """
        清空筛选，恢复原始数据（仅对当前 sheet）
        """
        try:
            # 获取当前表格的中文名称
            sheet_name = self.notebook.tab(self.notebook.select(), "text")
            if not sheet_name:
                raise ValueError("未找到当前选中的选项卡！")

            # 清空当前 sheet 的筛选缓存和历史记录
            self.filter_states[sheet_name]["filtered_data_cache"] = None
            self.filter_states[sheet_name]["filter_history"] = []

            # 根据表格类型加载数据
            if sheet_name == "序时账":
                # 序时账加载前 100 行
                self.load_from_db(sheet_name, limit=100, offset=0)
            elif sheet_name == "科目余额表":
                # 科目余额表全量加载
                self.load_from_db(sheet_name)
            else:
                # 其他表格（如凭证）全量加载
                self.load_from_db(sheet_name)

            # 更新 Treeview
            tree = self.trees[sheet_name]
            for item in tree.get_children():
                tree.delete(item)

            # 插入原始数据
            for _, row in self.sheets[sheet_name].iterrows():
                tree.insert("", "end", values=list(row))

            # 更新筛选框
            self.update_filter_entries(sheet_name)

        except Exception as e:
            messagebox.showerror("错误", f"清空筛选时出错: {e}")

    def update_filter_entries(self, sheet_name, filtered_df=None):
        """
        更新筛选框（仅保留输入框筛选）
        :param sheet_name: 表名（例如 "序时账" 或 "科目余额表"）
        :param filtered_df: 筛选后的数据（可选）
        """
        # 清空筛选框容器
        filter_frame = self.filter_frames[sheet_name]
        for widget in filter_frame.winfo_children():
            widget.destroy()

        # 获取当前表格的数据
        if filtered_df is None:
            df = self.sheets[sheet_name]
        else:
            df = filtered_df

        # 使用 grid 布局来排列筛选框
        for col_idx, col in enumerate(df.columns):
            # 创建输入框（文字筛选框）
            filter_entry = ttk.Entry(filter_frame)
            filter_entry.grid(row=0, column=col_idx, padx=5, pady=5, sticky="ew")  # 第一行

            # 绑定输入框的回车事件
            filter_entry.bind("<Return>", lambda event, c=col, e=filter_entry: self.apply_filter_from_entry(
                self.trees[sheet_name], c, e))

            # 设置列权重，使筛选框自适应宽度
            filter_frame.columnconfigure(col_idx, weight=1)

        # 保存筛选框的输入框到当前表的筛选状态
        self.filter_states[sheet_name]["filter_entries"] = {col: filter_entry for col_idx, col in enumerate(df.columns)}

    def apply_first_filter(self, tree, col, filter_text):
        """
        第一次筛选（从数据库中加载数据）
        :param tree: Treeview 控件
        :param col: 列名
        :param filter_text: 筛选条件
        """
        # 获取当前表格的中文名称
        sheet_name = self.notebook.tab(self.notebook.select(), "text")
        table_name = self.table_name_mapping.get(sheet_name)
        if not table_name:
            messagebox.showerror("错误", f"未找到表名映射：{sheet_name}")
            return

        # 使用 SQL 进行指定列的筛选
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()

            # 构建 SQL 查询
            query = f"SELECT * FROM {table_name} WHERE {col} LIKE ?"
            cursor.execute(query, (f"%{filter_text}%",))  # 使用 LIKE 进行模糊匹配
            filtered_data = cursor.fetchall()

        # 获取列名
        columns = self.sheets[sheet_name].columns

        # 将筛选结果转换为 Pandas DataFrame
        if filtered_data:
            filtered_df = pd.DataFrame(filtered_data, columns=columns)
        else:
            filtered_df = pd.DataFrame(columns=columns)

        # 保存筛选结果到当前表的筛选状态
        self.filter_states[sheet_name]["filtered_data_cache"] = filtered_df
        self.filter_states[sheet_name]["filter_history"].append(filtered_df.copy())  # 将结果添加到历史记录

        # 更新 Treeview
        self.update_treeview(tree, filtered_df)

    def apply_nth_filter(self, tree, col, filter_text):
        """
        第 N 次筛选（在内存中进行）
        :param tree: Treeview 控件
        :param col: 列名
        :param filter_text: 筛选条件
        """
        # 获取当前表格的中文名称
        sheet_name = self.notebook.tab(self.notebook.select(), "text")

        # 检查当前表的筛选缓存
        if (
                self.filter_states[sheet_name]["filtered_data_cache"] is None
                or self.filter_states[sheet_name]["filtered_data_cache"].empty
        ):
            messagebox.showwarning("警告", "请先进行第一次筛选！")
            return

        # 确保列名存在
        if col not in self.filter_states[sheet_name]["filtered_data_cache"].columns:
            messagebox.showerror("错误", f"列名 '{col}' 不存在！")
            return

        # 将指定列转换为字符串类型
        self.filter_states[sheet_name]["filtered_data_cache"][col] = \
        self.filter_states[sheet_name]["filtered_data_cache"][col].astype(str)

        # 在内存中进行筛选
        try:
            filtered_df = self.filter_states[sheet_name]["filtered_data_cache"][
                self.filter_states[sheet_name]["filtered_data_cache"][col].str.contains(filter_text, case=False,
                                                                                        na=False)
            ]
        except Exception as e:
            messagebox.showerror("错误", f"筛选时出错: {e}")
            return

        # 保存筛选结果到当前表的筛选状态
        self.filter_states[sheet_name]["filtered_data_cache"] = filtered_df
        self.filter_states[sheet_name]["filter_history"].append(filtered_df.copy())  # 将结果添加到历史记录

        # 更新 Treeview
        self.update_treeview(tree, filtered_df)

    def show_detail_journal(self, event):
        """
        在科目余额表右键时，从数据库中筛选出与选中科目相关的序时账数据，并显示在序时账 Treeview 中。
        """
        try:

            # 获取选中的行
            selected_item = self.trees["科目余额表"].selection()
            if not selected_item:
                return

            # 获取科目编码
            item_values = self.trees["科目余额表"].item(selected_item, "values")
            subject_code = str(item_values[0]).strip()  # 将科目编码转换为字符串并去除空格

            if not subject_code:
                messagebox.showwarning("警告", "未选中有效的科目编码！")
                return

            # 使用 SQL 查询从数据库中筛选数据
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()

                # 构建 SQL 查询
                query = '''
                    SELECT * FROM journal
                    WHERE 科目编码 = ?
                '''
                cursor.execute(query, (subject_code,))
                filtered_data = cursor.fetchall()

            # 检查是否有匹配的数据
            if not filtered_data:
                messagebox.showinfo("提示", f"未找到科目编码为 {subject_code} 的明细账！")
                return

            # 将筛选结果转换为 Pandas DataFrame
            columns = ["日期", "凭证字号", "科目编码", "科目名称", "辅助核算", "摘要", "借方", "贷方", "数量", "外币"]
            filtered_df = pd.DataFrame(filtered_data, columns=columns)

            # 清空当前序时账 Treeview
            tree = self.trees["序时账"]
            for item in tree.get_children():
                tree.delete(item)

            # 插入筛选后的数据
            for _, row in filtered_df.iterrows():
                tree.insert("", "end", values=list(row))

            # 切换到序时账选项卡
            self.notebook.select(self.trees["序时账"].master)

            # 标记为已筛选状态
            self.is_filtered = True

        except Exception as e:
            messagebox.showerror("错误", f"显示明细账时出错: {e}")

    def show_voucher_details(self, event):
        """
        在序时账右键时，根据选中的凭证编号，从数据库中筛选出相同凭证编号的记录，
        并将结果交给 Pandas 处理，放入凭证 Sheet 中。
        """
        try:
            # 获取选中的行
            selected_item = self.trees["序时账"].selection()
            if not selected_item:
                return

            # 获取选中行的数据
            item_values = self.trees["序时账"].item(selected_item, "values")
            selected_voucher = item_values[1]  # 凭证字号
            selected_date = item_values[0]  # 日期

            # 检查凭证字号和日期是否为空
            if not selected_voucher or not selected_date:
                messagebox.showwarning("警告", "未选中有效的凭证！")
                return

            # 使用 SQL 查询从数据库中筛选数据
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()

                # 构建 SQL 查询
                query = '''
                    SELECT * FROM journal
                    WHERE 凭证字号 = ? AND 日期 = ?
                '''
                cursor.execute(query, (selected_voucher, selected_date))
                filtered_data = cursor.fetchall()

            # 检查是否有匹配的数据
            if not filtered_data:
                messagebox.showwarning("警告", "未找到符合条件的凭证记录！")
                return

            # 将筛选结果转换为 Pandas DataFrame
            columns = ["日期", "凭证字号", "科目编码", "科目名称", "辅助核算", "摘要", "借方", "贷方", "数量", "外币"]
            filtered_df = pd.DataFrame(filtered_data, columns=columns)

            # 清空凭证表
            self.sheets["凭证"] = pd.DataFrame(
                columns=["日期", "凭证字号", "摘要", "科目名称", "借方", "贷方", "数量", "外币"])

            # 将筛选后的数据插入凭证表
            new_rows = []
            for _, row in filtered_df.iterrows():
                project_name = row.get("科目名称", "").strip()
                if project_name:
                    project_name = f"【{project_name}】"

                new_row = {
                    "日期": pd.to_datetime(row["日期"]),  # 统一为 datetime 类型
                    "凭证字号": str(row["凭证字号"]),  # 统一为字符串类型
                    "摘要": str(row["摘要"]),  # 统一为字符串类型
                    "科目名称": f"{str(row['科目名称'])}{project_name}",  # 统一为字符串类型
                    "借方": float(row["借方"]) if row["借方"] != 0 else "",  # 统一为浮点数类型
                    "贷方": float(row["贷方"]) if row["贷方"] != 0 else "",  # 统一为浮点数类型
                    "数量": row["数量"] if row["数量"] != 0 else "",  # 保持原始值
                    "外币": row["外币"] if row["外币"] != 0 else ""  # 保持原始值
                }
                new_rows.append(new_row)

            # 使用 pd.concat 添加新行
            if new_rows:
                self.sheets["凭证"] = pd.concat([self.sheets["凭证"], pd.DataFrame(new_rows)], ignore_index=True)

            # 将 "借方" 和 "贷方" 列中的空字符串替换为 0
            self.sheets["凭证"]["借方"] = self.sheets["凭证"]["借方"].replace("", 0).astype(float)
            self.sheets["凭证"]["贷方"] = self.sheets["凭证"]["贷方"].replace("", 0).astype(float)

            # 计算借方和贷方的合计
            total_debit = self.sheets["凭证"]["借方"].sum()
            total_credit = self.sheets["凭证"]["贷方"].sum()

            # 添加合计行
            total_row = {
                "摘要": "合     计",
                "借方": total_debit,
                "贷方": total_credit
            }
            self.sheets["凭证"] = pd.concat([self.sheets["凭证"], pd.DataFrame([total_row])], ignore_index=True)

            # 检查借贷是否平衡
            if abs(total_debit - total_credit) > 1e-6:  # 浮点数精度问题
                messagebox.showerror("错误", "数据有误，请检查！详见【凭证】")

            # 更新凭证 Treeview
            tree = self.trees["凭证"]
            for item in tree.get_children():
                tree.delete(item)

            for _, row in self.sheets["凭证"].iterrows():
                tree.insert("", "end", values=list(row))

            # 切换到凭证选项卡
            self.notebook.select(self.trees["凭证"].master)

        except Exception as e:
            messagebox.showerror("错误", f"显示凭证信息时出错: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("800x600")  # 设置初始窗口大小
    app = ExcelLikeApp(root)
    root.mainloop()
