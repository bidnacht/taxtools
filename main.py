import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import numpy as np


class TaxDataProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("成效统计工具")
        self.root.geometry("600x500")

        # 存储数据 - 减少内存占用，使用更小的数据类型
        self.main_table = None
        self.reduce_table = None
        self.previous_table = None
        self.current_table = None

        # 缓存替换字典 - 优化为元组列表便于迭代
        self.replacement_patterns = [
            ('武汉东湖新技术开发区税务', '东湖'),
            ('武汉化学工业区税务', '化工'),
            ('武汉化学工业区税务分', '化工'),
            ('武汉市东西湖区税务', '东西湖'),
            ('武汉市新洲区税务', '新洲'),
            ('武汉市武昌区税务', '武昌'),
            ('武汉市汉阳区税务', '汉阳'),
            ('武汉市江夏区税务', '江夏'),
            ('武汉市江岸区税务', '江岸'),
            ('武汉市江汉区税务', '江汉'),
            ('武汉市洪山区税务', '洪山'),
            ('武汉市硚口区税务', '硚口'),
            ('武汉市蔡甸区税务', '蔡甸'),
            ('武汉市青山区税务', '青山'),
            ('武汉市黄陂区税务', '黄陂'),
            ('武汉经济技术开发区（汉南区）税务', '武经'),
            ('武汉长江新区税务', '长江新区'),
            ('武汉市东湖生态旅游风景区税务', '洪山'),
            ('第一税务分', '一分局'),
            ('第二税务分', '二分局'),
        ]

        # 转换为字典用于快速查找（如果需要的话）
        self.replacement_dict = dict(self.replacement_patterns)

        # 定义区局排序表 - 使用元组更节省内存
        self.district_order = (
            '江岸', '江汉', '硚口', '汉阳', '武昌', '青山',
            '东湖', '武经', '洪山', '东西湖', '蔡甸', '江夏',
            '黄陂', '新洲', '长江新区', '化工', '一分局', '二分局'
        )

        # 创建区局集合用于快速查找
        self.district_set = set(self.district_order)

        # 创建UI
        self.create_widgets()

    def create_widgets(self):
        # 标题
        title_label = tk.Label(self.root, text="税务数据处理系统", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)

        # 主表选择
        main_frame = tk.Frame(self.root)
        main_frame.pack(pady=5, fill="x", padx=20)

        tk.Label(main_frame, text="选择主表:").pack(side="left")
        self.main_file_label = tk.Label(main_frame, text="未选择文件", fg="red")
        self.main_file_label.pack(side="left", padx=10)
        tk.Button(main_frame, text="选择文件", command=self.select_main_file).pack(side="right")

        # 调减增收表选择
        reduce_frame = tk.Frame(self.root)
        reduce_frame.pack(pady=5, fill="x", padx=20)

        tk.Label(reduce_frame, text="选择调减增收表:").pack(side="left")
        self.reduce_file_label = tk.Label(reduce_frame, text="未选择文件", fg="red")
        self.reduce_file_label.pack(side="left", padx=10)
        tk.Button(reduce_frame, text="选择文件", command=self.select_reduce_file).pack(side="right")

        # 往年入库表选择
        previous_frame = tk.Frame(self.root)
        previous_frame.pack(pady=5, fill="x", padx=20)

        tk.Label(previous_frame, text="选择往年入库表:").pack(side="left")
        self.previous_file_label = tk.Label(previous_frame, text="未选择文件", fg="red")
        self.previous_file_label.pack(side="left", padx=10)
        tk.Button(previous_frame, text="选择文件", command=self.select_previous_file).pack(side="right")

        # 今年入库表选择
        current_frame = tk.Frame(self.root)
        current_frame.pack(pady=5, fill="x", padx=20)

        tk.Label(current_frame, text="选择今年入库表:").pack(side="left")
        self.current_file_label = tk.Label(current_frame, text="未选择文件", fg="red")
        self.current_file_label.pack(side="left", padx=10)
        tk.Button(current_frame, text="选择文件", command=self.select_current_file).pack(side="right")

        # 处理按钮
        process_btn = tk.Button(self.root, text="开始处理数据", command=self.process_data,
                                bg="green", fg="white", font=("Arial", 12, "bold"))
        process_btn.pack(pady=20)

        # 进度条
        self.progress = ttk.Progressbar(self.root, orient="horizontal", length=400, mode="determinate")
        self.progress.pack(pady=10)

        # 状态标签
        self.status_label = tk.Label(self.root, text="请选择所有需要的Excel文件", font=("Arial", 10))
        self.status_label.pack(pady=10)

        # 日志文本框
        self.log_text = tk.Text(self.root, height=10, width=70)
        self.log_text.pack(pady=10, padx=20, fill="both", expand=True)

        # 添加滚动条
        scrollbar = tk.Scrollbar(self.log_text)
        scrollbar.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.log_text.yview)

    def log_message(self, message):
        """优化日志记录，减少UI更新频率"""
        self.log_text.insert("end", f"{message}\n")
        # 每记录5条消息更新一次UI，提高性能
        try:
            # 直接获取行数信息
            line_no = int(self.log_text.index('end-1c').split('.')[0])
            if line_no % 5 == 0:
                self.log_text.see("end")
                self.root.update_idletasks()
        except (ValueError, IndexError):
            # 如果无法获取行号，则每次都更新
            self.log_text.see("end")
            self.root.update_idletasks()
    def select_main_file(self):
        filename = filedialog.askopenfilename(title="选择主表", filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.main_file_label.config(text=os.path.basename(filename), fg="green")
            self.main_table_path = filename

    def select_reduce_file(self):
        filename = filedialog.askopenfilename(title="选择调减增收表", filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.reduce_file_label.config(text=os.path.basename(filename), fg="green")
            self.reduce_table_path = filename

    def select_previous_file(self):
        filename = filedialog.askopenfilename(title="选择往年入库表", filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.previous_file_label.config(text=os.path.basename(filename), fg="green")
            self.previous_table_path = filename

    def select_current_file(self):
        filename = filedialog.askopenfilename(title="选择今年入库表", filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.current_file_label.config(text=os.path.basename(filename), fg="green")
            self.current_table_path = filename

    def load_and_preprocess_table(self, file_path):
        """加载并预处理表格，删除前两行，优化内存使用"""
        try:
            # 使用dtype参数减少内存使用，使用更高效的引擎
            df = pd.read_excel(file_path, header=2, engine='openpyxl')
            # 将所有可能的数值列转换为浮点型
            df = self.convert_all_numeric_columns(df)
            # 自动优化数值列的数据类型
            df = self.optimize_dataframe_dtypes(df)

            return df
        except Exception as e:
            try:
                df = pd.read_excel(file_path, header=2)
                df = self.convert_all_numeric_columns(df)
                df = self.optimize_dataframe_dtypes(df)
                return df
            except Exception as e:
                raise Exception(f"读取文件失败: {str(e)}")

    def convert_all_numeric_columns(self, df):
        """优化DataFrame的数据类型以减少内存使用"""
        # 转换数值列
        self.log_message("开始转换数值列格式...")
        numeric_patterns = [
            '入库', '收入', '合计', '金额', '税额', '税款', '金额',
            '调减', '增收', '统计', '金额', '总计', '合计', '总计'
        ]
        converted_count = 0
        for col in df.columns:
            col_type = df[col].dtype
            # 如果已经是数值类型，跳过
            if pd.api.types.is_numeric_dtype(df[col]):
                continue
            # 检查列名是否包含数值相关的关键词
            is_numeric_col = False
            if isinstance(col, str):
                for pattern in numeric_patterns:
                    if pattern in col:
                        is_numeric_col = True
                        break

            # 尝试转换为数值类型
            try:
                # 先尝试直接转换
                converted = pd.to_numeric(df[col], errors='coerce')

                # 如果转换成功（非空值比例超过一定阈值）
                not_null_count = converted.notna().sum()
                if not_null_count > 0 and not_null_count / len(df) > 0.1:
                    df[col] = converted
                    converted_count += 1
                    self.log_message(f"  转换列 '{col}': {col_type} -> float64")
            except Exception as e:
                # 转换失败，保持原样
                pass

        self.log_message(f"数值列格式转换完成，共转换{converted_count}列")
        return df

    def optimize_dataframe_dtypes(self, df):
        """优化DataFrame的数据类型以减少内存使用"""
        # 转换数值列
        for col in df.columns:
            col_type = df[col].dtype

            # 如果是object类型，尝试转换为数值类型
            if col_type == 'object':
                try:
                    # 尝试转换为数值类型
                    converted = pd.to_numeric(df[col], errors='ignore')
                    if converted.dtype != 'object':  # 如果转换成功
                        df[col] = converted
                        # 转换为float32以确保小数点
                        df[col] = df[col].astype('float32')
                except:
                    pass

            # 优化数值类型 - 确保所有数值列都是浮点数以保留小数
            if pd.api.types.is_numeric_dtype(df[col]):
                # 对于浮点数列，直接使用float32或float64
                if pd.api.types.is_float_dtype(df[col]):
                    # 对于汇总表，使用float64确保精度
                    if col in ['应对入库合计', '调减收入合计', '往年入库合计', '往年调减收入合计',
                               '今年入库金额', '今年非税收入', '今年调减收入', '今年入库统计总额']:
                        df[col] = df[col].astype('float64')
                    else:
                        df[col] = df[col].astype('float32')
                # 对于整数列，如果有小数可能，转换为float32
                elif pd.api.types.is_integer_dtype(df[col]):
                    # 检查列名是否可能包含金额数据
                    money_keywords = ['金额', '收入', '税额', '合计', '总计', '税款', '入库']
                    col_name = str(col).lower()
                    if any(keyword in col_name for keyword in money_keywords):
                        df[col] = df[col].astype('float32')

        return df
    def filter_and_sort_by_district_order(self, df, district_col='所属区局'):
        """过滤并按照排序表顺序排列数据，优化性能"""
        if district_col not in df.columns:
            self.log_message(f"警告: 数据中缺少'{district_col}'列，无法按区局排序")
            return df

        # 记录过滤前的数据量
        before_count = len(df)

        # 使用集合进行快速查找，过滤掉不在排序表中的区局
        df_filtered = df[df[district_col].isin(self.district_set)].copy()

        # 记录过滤掉的数据量
        filtered_count = before_count - len(df_filtered)
        if filtered_count > 0:
            filtered_districts = df.loc[~df[district_col].isin(self.district_set), district_col].unique()
            self.log_message(f"过滤掉{filtered_count}条数据，涉及区局: {', '.join(filtered_districts)}")

        # 创建排序映射字典 - 使用字典推导式
        order_dict = {district: i for i, district in enumerate(self.district_order)}

        # 添加排序列
        df_filtered['_order'] = df_filtered[district_col].map(order_dict)

        # 按照排序列排序
        df_filtered = df_filtered.sort_values('_order')

        # 删除临时排序列并释放内存
        df_filtered = df_filtered.drop('_order', axis=1)

        self.log_message(f"按区局排序完成，剩余{len(df_filtered)}条数据")
        return df_filtered

    def process_main_table(self, df):
        """处理主表，优化字符串处理性能"""
        self.log_message("开始处理主表...")

        # 1. 处理R列和S列
        r_col = "当前处理机关（团队/网格）/人员"
        s_col = "风险应对部门"

        # 使用fillna原地操作，避免创建新列
        if r_col in df.columns and s_col in df.columns:
            df[r_col] = df[r_col].fillna(df[s_col])

        # 2. 使用向量化操作拆分R列
        if r_col in df.columns:
            # 使用str.split的expand参数，但限制分割次数
            split_result = df[r_col].str.split("局", n=3, expand=True)

            # 使用iloc直接获取需要的列，避免循环
            if split_result.shape[1] >= 4:
                df['split_col2'] = split_result.iloc[:, 1]
                df['split_col3'] = split_result.iloc[:, 2]
            elif split_result.shape[1] >= 3:
                df['split_col2'] = split_result.iloc[:, 1]
                df['split_col3'] = None
            else:
                df['split_col2'] = None
                df['split_col3'] = None

        # 3. 去掉第二列为空的数据行
        if 'split_col2' in df.columns:
            df = df.dropna(subset=['split_col2'])

        # 4. 处理"武汉市税务"的数据
        if 'split_col2' in df.columns and 'split_col3' in df.columns:
            wuhan_mask = df['split_col2'] == "武汉市税务"
            df.loc[wuhan_mask, 'split_col2'] = df.loc[wuhan_mask, 'split_col3']

        # 5. 创建新S列并命名为"所属区局"
        if 'split_col2' in df.columns:
            df['所属区局'] = df['split_col2']

        # 6. 应用替换规则 - 使用向量化操作
        if '所属区局' in df.columns:
            # 构建替换函数
            def replace_district(district):
                if pd.isna(district):
                    return district
                # 使用循环查找，但比多次调用map快
                for pattern, replacement in self.replacement_patterns:
                    if pattern in str(district):
                        return replacement
                return district

            # 使用apply但避免lambda
            df['所属区局'] = df['所属区局'].apply(replace_district)

        # 7. 处理"应对入库合计"列 - 转换为数值格式
        if '应对入库合计' in df.columns:
            df['应对入库合计'] = pd.to_numeric(df['应对入库合计'], errors='coerce')
            nan_count = df['应对入库合计'].isna().sum()
            if nan_count > 0:
                self.log_message(f"警告: '应对入库合计'列有{nan_count}个值无法转换为数字，已设为NaN")

        # 删除拆分过程中产生的中间列
        cols_to_drop = [col for col in ['split_col2', 'split_col3'] if col in df.columns]
        if cols_to_drop:
            df = df.drop(cols_to_drop, axis=1)

        # 优化数据类型
        df = self.optimize_dataframe_dtypes(df)

        self.log_message(f"主表处理完成，剩余{len(df)}行数据")
        return df

    def merge_additional_data(self, main_df):
        """合并调减收入和往年入库数据，优化合并性能"""
        self.log_message("开始合并调减收入和往年入库数据...")

        # 合并调减收入数据
        if hasattr(self, 'reduce_table_path') and self.reduce_table_path:
            try:
                reduce_df = self.load_and_preprocess_table(self.reduce_table_path)

                # 检查必要的列是否存在
                required_cols = ['任务批次名称', '纳税人名称', '统计调减收入']
                if all(col in reduce_df.columns for col in required_cols):
                    # 重命名列
                    reduce_df = reduce_df.rename(columns={'统计调减收入': '调减收入'})

                    # 对调减收入表进行去重处理，只保留第一条记录
                    # 使用keep='first'参数
                    reduce_df_unique = reduce_df.drop_duplicates(
                        subset=['任务批次名称', '纳税人名称'],
                        keep='first'
                    )

                    # 检查是否有重复记录
                    if len(reduce_df) != len(reduce_df_unique):
                        duplicate_count = len(reduce_df) - len(reduce_df_unique)
                        self.log_message(f"注意: 调减收入表中有{duplicate_count}条重复记录，已去除重复")

                    # 优化合并：只选择需要的列
                    merge_cols = ['任务批次名称', '纳税人名称', '调减收入']
                    main_df = pd.merge(
                        main_df,
                        reduce_df_unique[merge_cols],
                        on=['任务批次名称', '纳税人名称'],
                        how='left',
                        copy=False  # 避免不必要的数据复制
                    )

                    # 验证合并结果
                    merged_rows = main_df['调减收入'].notna().sum()
                    self.log_message(f"调减收入数据合并完成，成功匹配{merged_rows}条记录")

                    # 释放内存
                    del reduce_df, reduce_df_unique
                else:
                    missing_cols = [col for col in required_cols if col not in reduce_df.columns]
                    self.log_message(f"警告: 调减增收表中缺少列: {missing_cols}")
            except Exception as e:
                self.log_message(f"合并调减收入数据时出错: {str(e)}")

        # 合并往年入库数据
        if hasattr(self, 'previous_table_path') and self.previous_table_path:
            try:
                previous_df = self.load_and_preprocess_table(self.previous_table_path)

                # 检查必要的列是否存在
                required_cols = ['任务批次名称', '纳税人名称', '入库金额', '统计调减收入']
                if all(col in previous_df.columns for col in required_cols):
                    # 重命名列
                    previous_df = previous_df.rename(columns={
                        '入库金额': '往年入库',
                        '统计调减收入': '往年调减收入'  # 新增：重命名统计调减收入列
                    })

                    # 对往年入库表进行去重处理，只保留第一条记录
                    previous_df_unique = previous_df.drop_duplicates(
                        subset=['任务批次名称', '纳税人名称'],
                        keep='first'
                    )

                    # 检查是否有重复记录
                    if len(previous_df) != len(previous_df_unique):
                        duplicate_count = len(previous_df) - len(previous_df_unique)
                        self.log_message(f"注意: 往年入库表中有{duplicate_count}条重复记录，已去除重复")

                    # 优化合并
                    merge_cols = ['任务批次名称', '纳税人名称', '往年入库', '往年调减收入']  # 新增：包含往年调减收入列
                    main_df = pd.merge(
                        main_df,
                        previous_df_unique[merge_cols],
                        on=['任务批次名称', '纳税人名称'],
                        how='left',
                        copy=False
                    )

                    # 验证合并结果
                    merged_rows = main_df['往年入库'].notna().sum()
                    merged_reduce_rows = main_df['往年调减收入'].notna().sum()
                    self.log_message(
                        f"往年入库数据合并完成，成功匹配{merged_rows}条入库记录，{merged_reduce_rows}条调减记录")

                    # 释放内存
                    del previous_df, previous_df_unique
                else:
                    missing_cols = [col for col in required_cols if col not in previous_df.columns]
                    self.log_message(f"警告: 往年入库表中缺少列: {missing_cols}")
            except Exception as e:
                self.log_message(f"合并往年入库数据时出错: {str(e)}")

        # 去除重复行
        original_rows = len(main_df)
        main_df = main_df.drop_duplicates()

        if len(main_df) != original_rows:
            self.log_message(f"注意: 去除{original_rows - len(main_df)}条重复主表记录")

        # 优化数据类型
        main_df = self.optimize_dataframe_dtypes(main_df)

        self.log_message(f"合并完成，主表最终有{len(main_df)}条记录")
        return main_df

    def process_current_table(self):
        """处理今年入库表 - 按照新规则处理税务机关列，优化性能"""
        self.log_message("开始处理今年入库表...")

        if not hasattr(self, 'current_table_path') or not self.current_table_path:
            self.log_message("警告: 未选择今年入库表")
            return None

        current_df = self.load_and_preprocess_table(self.current_table_path)

        # 1. 检查必要的列是否存在
        tax_authority_col = None
        possible_cols = ['税务机关', '当前处理机关（团队/网格）/人员', '风险应对部门']

        for col in possible_cols:
            if col in current_df.columns:
                tax_authority_col = col
                break

        if tax_authority_col is None:
            self.log_message("错误: 今年入库表中找不到税务机关相关的列")
            return None

        # 2. 处理税务机关列
        r_col = tax_authority_col

        # 如果找到的是"当前处理机关（团队/网格）/人员"列且有"风险应对部门"列，则进行填充
        s_col = "风险应对部门"
        if r_col == "当前处理机关（团队/网格）/人员" and s_col in current_df.columns:
            current_df[r_col] = current_df[r_col].fillna(current_df[s_col])

        # 使用向量化操作拆分列
        if r_col in current_df.columns:
            split_result = current_df[r_col].str.split("局", n=3, expand=True)

            # 使用iloc直接获取需要的列
            if split_result.shape[1] >= 4:
                current_df['split_col2'] = split_result.iloc[:, 1]
                current_df['split_col3'] = split_result.iloc[:, 2]
            elif split_result.shape[1] >= 3:
                current_df['split_col2'] = split_result.iloc[:, 1]
                current_df['split_col3'] = None
            else:
                current_df['split_col2'] = None
                current_df['split_col3'] = None

        # 去掉第二列为空的数据行
        if 'split_col2' in current_df.columns:
            current_df = current_df.dropna(subset=['split_col2'])

        # 处理"武汉市税务"的数据
        if 'split_col2' in current_df.columns and 'split_col3' in current_df.columns:
            wuhan_mask = current_df['split_col2'] == "武汉市税务"
            current_df.loc[wuhan_mask, 'split_col2'] = current_df.loc[wuhan_mask, 'split_col3']

        # 创建所属区局列
        if 'split_col2' in current_df.columns:
            current_df['所属区局'] = current_df['split_col2']

        # 应用替换规则
        if '所属区局' in current_df.columns:
            # 使用与主表相同的替换函数
            def replace_district(district):
                if pd.isna(district):
                    return district
                for pattern, replacement in self.replacement_patterns:
                    if pattern in str(district):
                        return replacement
                return district

            current_df['所属区局'] = current_df['所属区局'].apply(replace_district)

        # 删除中间列
        cols_to_drop = [col for col in ['split_col2', 'split_col3'] if col in current_df.columns]
        if cols_to_drop:
            current_df = current_df.drop(cols_to_drop, axis=1)

        # 3. 处理数值列
        # 定义需要处理的列
        num_cols = ['入库金额', '其中：非税收入', '统计调减收入']

        # 批量转换数值列
        for col in num_cols:
            if col in current_df.columns:
                current_df[col] = pd.to_numeric(current_df[col], errors='coerce').fillna(0)
            else:
                current_df[col] = 0
                self.log_message(f"警告: 今年入库表中缺少'{col}'列，已设为0")

        # 计算统计总额 - 不再减去非税收入
        current_df['今年入库统计总额'] = current_df['入库金额'] - current_df['统计调减收入']

        # 优化数据类型
        current_df = self.optimize_dataframe_dtypes(current_df)

        # 记录处理结果
        unique_districts = current_df['所属区局'].nunique()
        self.log_message(f"今年入库表处理完成：")
        self.log_message(f"  - 共{len(current_df)}行数据")
        self.log_message(f"  - 涉及{unique_districts}个不同的区局")
        self.log_message(f"  - 入库金额总计: {current_df['入库金额'].sum():.2f}")
        self.log_message(f"  - 其中：非税收入总计: {current_df['其中：非税收入'].sum():.2f}")
        self.log_message(f"  - 统计调减收入总计: {current_df['统计调减收入'].sum():.2f}")
        self.log_message(f"  - 今年入库统计总额: {current_df['今年入库统计总额'].sum():.2f}")

        return current_df

    def create_summary_sheet(self, main_df, current_df, title="统计汇总"):
        """创建汇总表，使用更高效的分组和聚合方法"""
        self.log_message(f"创建{title}表...")

        # 检查必要的列是否存在
        if '所属区局' not in main_df.columns:
            raise Exception("主表中缺少必要列: 所属区局")
        if '应对入库合计' not in main_df.columns:
            raise Exception("主表中缺少必要列: 应对入库合计")

        # 获取所有区局（合并主表和今年入库表的区局）
        main_districts = set(main_df['所属区局'].unique())
        if current_df is not None:
            current_districts = set(current_df['所属区局'].unique())
        else:
            current_districts = set()

        all_districts = sorted(main_districts.union(current_districts))

        # 预计算主表分组结果，避免重复分组
        main_grouped = main_df.groupby('所属区局')

        # 使用列表推导式提高效率
        summary_data = []

        # 创建当前数据的分组（如果存在）
        if current_df is not None and '所属区局' in current_df.columns:
            current_grouped = current_df.groupby('所属区局')
        else:
            current_grouped = None

        # 处理排序表中的区局
        for district in self.district_order:
            if district not in all_districts:
                continue

            # 获取主表该区局的数据
            if district in main_grouped.groups:
                main_district_data = main_grouped.get_group(district)

                # 计算主表各项合计
                response_total = main_district_data['应对入库合计'].sum()
                task_count = len(main_district_data)

                # 调减收入合计
                reduce_total = main_district_data['调减收入'].sum() if '调减收入' in main_district_data.columns else 0

                # 往年入库合计
                previous_total = main_district_data['往年入库'].sum() if '往年入库' in main_district_data.columns else 0

                # 往年调减收入合计 - 新增
                previous_reduce_total = main_district_data[
                    '往年调减收入'].sum() if '往年调减收入' in main_district_data.columns else 0
            else:
                response_total = 0
                task_count = 0
                reduce_total = 0
                previous_total = 0
                previous_reduce_total = 0  # 新增

            # 获取今年入库表该区局的数据
            if current_grouped is not None and district in current_grouped.groups:
                current_district_data = current_grouped.get_group(district)

                current_rows = len(current_district_data)
                current_in_total = current_district_data[
                    '入库金额'].sum() if '入库金额' in current_district_data.columns else 0
                current_reduce_total = current_district_data[
                    '统计调减收入'].sum() if '统计调减收入' in current_district_data.columns else 0
                current_taxfree_total = current_district_data[
                    '其中：非税收入'].sum() if '其中：非税收入' in current_district_data.columns else 0
                current_total = current_district_data[
                    '今年入库统计总额'].sum() if '今年入库统计总额' in current_district_data.columns else 0
            else:
                current_rows = 0
                current_in_total = 0
                current_reduce_total = 0
                current_taxfree_total = 0
                current_total = 0

            # 添加到汇总数据
            summary_data.append({
                '所属区局': district,
                '应对入库合计': response_total,
                '调减收入合计': reduce_total,
                '往年入库合计': previous_total,
                '往年调减收入合计': previous_reduce_total,  # 新增
                '今年入库金额': current_in_total,
                '今年非税收入': current_taxfree_total,
                '今年调减收入': current_reduce_total,
                '今年入库统计总额': current_total,
                '总任务条数': task_count,
                '今年入库条数': current_rows
            })

        # 处理不在排序表中的区局
        remaining_districts = [d for d in all_districts if d not in self.district_order]
        if remaining_districts:
            self.log_message(f"注意: 以下区局不在排序表中，将放在最后: {remaining_districts}")

            for district in remaining_districts:
                # 获取主表该区局的数据
                if district in main_grouped.groups:
                    main_district_data = main_grouped.get_group(district)

                    response_total = main_district_data['应对入库合计'].sum()
                    task_count = len(main_district_data)
                    reduce_total = main_district_data[
                        '调减收入'].sum() if '调减收入' in main_district_data.columns else 0
                    previous_total = main_district_data[
                        '往年入库'].sum() if '往年入库' in main_district_data.columns else 0
                    previous_reduce_total = main_district_data[
                        '往年调减收入'].sum() if '往年调减收入' in main_district_data.columns else 0  # 新增
                else:
                    response_total = 0
                    task_count = 0
                    reduce_total = 0
                    previous_total = 0
                    previous_reduce_total = 0  # 新增

                # 获取今年入库表该区局的数据
                if current_grouped is not None and district in current_grouped.groups:
                    current_district_data = current_grouped.get_group(district)

                    current_rows = len(current_district_data)
                    current_in_total = current_district_data[
                        '入库金额'].sum() if '入库金额' in current_district_data.columns else 0
                    current_reduce_total = current_district_data[
                        '统计调减收入'].sum() if '统计调减收入' in current_district_data.columns else 0
                    current_taxfree_total = current_district_data[
                        '其中：非税收入'].sum() if '其中：非税收入' in current_district_data.columns else 0
                    current_total = current_district_data[
                        '今年入库统计总额'].sum() if '今年入库统计总额' in current_district_data.columns else 0
                else:
                    current_rows = 0
                    current_in_total = 0
                    current_reduce_total = 0
                    current_taxfree_total = 0
                    current_total = 0

                summary_data.append({
                    '所属区局': district,
                    '应对入库合计': response_total,
                    '调减收入合计': reduce_total,
                    '往年入库合计': previous_total,
                    '往年调减收入合计': previous_reduce_total,  # 新增
                    '今年入库金额': current_in_total,
                    '今年非税收入': current_taxfree_total,
                    '今年调减收入': current_reduce_total,
                    '今年入库统计总额': current_total,
                    '总任务条数': task_count,
                    '今年入库条数': current_rows
                })

        # 创建汇总DataFrame
        summary_df = pd.DataFrame(summary_data)

        # 添加总计行
        total_row = {
            '所属区局': '总计',
            '应对入库合计': summary_df['应对入库合计'].sum(),
            '调减收入合计': summary_df['调减收入合计'].sum(),
            '往年入库合计': summary_df['往年入库合计'].sum(),
            '往年调减收入合计': summary_df['往年调减收入合计'].sum(),  # 新增
            '今年入库金额': summary_df['今年入库金额'].sum(),
            '今年非税收入': summary_df['今年非税收入'].sum(),
            '今年调减收入': summary_df['今年调减收入'].sum(),
            '今年入库统计总额': summary_df['今年入库统计总额'].sum(),
            '总任务条数': summary_df['总任务条数'].sum(),
            '今年入库条数': summary_df['今年入库条数'].sum()
        }
        summary_df = pd.concat([summary_df, pd.DataFrame([total_row])], ignore_index=True)

        # 优化数据类型 - 确保数值列使用浮点数类型以保留小数
        numeric_cols = [
            '应对入库合计', '调减收入合计', '往年入库合计', '往年调减收入合计',
            '今年入库金额', '今年非税收入', '今年调减收入', '今年入库统计总额'
        ]

        for col in numeric_cols:
            if col in summary_df.columns:
                summary_df[col] = pd.to_numeric(summary_df[col], errors='coerce').astype('float64')

        self.log_message(f"{title}表创建完成")
        return summary_df

    def create_no_first_bureau_sheet(self, main_df, current_df):
        """创建无一分局分摊表，优化字符串处理和分组性能"""
        self.log_message("创建无一分局分摊表...")

        # 检查必要的列是否存在
        if '所属区局' not in main_df.columns:
            raise Exception("主表中缺少必要列: 所属区局")
        if '应对入库合计' not in main_df.columns:
            raise Exception("主表中缺少必要列: 应对入库合计")
        if '主管税务机关' not in main_df.columns:
            self.log_message("警告: 主表中缺少'主管税务机关'列，无法进行一分局分摊")
            return None

        # 创建副本，但只复制必要的列以减少内存
        needed_cols = ['所属区局', '应对入库合计', '主管税务机关']
        if '调减收入' in main_df.columns:
            needed_cols.append('调减收入')
        if '往年入库' in main_df.columns:
            needed_cols.append('往年入库')
        if '往年调减收入' in main_df.columns:  # 新增
            needed_cols.append('往年调减收入')

        df_copy = main_df[needed_cols].copy()

        # 预编译替换模式，提高性能
        def convert_tax_authority(authority_str):
            if pd.isna(authority_str):
                return None
            authority_str = str(authority_str)
            # 直接使用字符串查找
            for pattern, replacement in self.replacement_patterns:
                if pattern in authority_str:
                    return replacement
            return None

        # 对主管税务机关列进行转换
        df_copy['主管税务机关_简称'] = df_copy['主管税务机关'].apply(convert_tax_authority)

        # 统计原始一分局的数据
        first_bureau_mask = df_copy['所属区局'] == '一分局'
        first_bureau_count = first_bureau_mask.sum()
        first_bureau_total = df_copy.loc[first_bureau_mask, '应对入库合计'].sum()

        self.log_message(f"原始数据中一分局记录数: {first_bureau_count}")
        self.log_message(f"原始数据中一分局应对入库合计总计: {first_bureau_total:.2f}")

        # 创建新区局列
        df_copy['所属区局_新'] = df_copy['所属区局']

        # 对于一分局的记录，使用主管税务机关_简称替换所属区局
        df_copy.loc[first_bureau_mask, '所属区局_新'] = df_copy.loc[first_bureau_mask, '主管税务机关_简称']

        # 处理主管税务机关_简称为空的情况
        null_auth_mask = first_bureau_mask & df_copy['主管税务机关_简称'].isna()
        null_auth_count = null_auth_mask.sum()
        if null_auth_count > 0:
            self.log_message(f"警告: {null_auth_count}条一分局记录的主管税务机关为空或无法识别")
            df_copy.loc[null_auth_mask, '所属区局_新'] = '未知区局'

        # 按新的区局分组统计
        new_grouped = df_copy.groupby('所属区局_新')
        all_new_districts = set(new_grouped.groups.keys())

        # 创建今年入库表分组（如果存在）
        if current_df is not None and '所属区局' in current_df.columns:
            current_grouped = current_df.groupby('所属区局')
        else:
            current_grouped = None

        summary_data = []

        # 先处理排序表中的区局
        for district in self.district_order:
            if district not in all_new_districts:
                continue

            district_data = new_grouped.get_group(district)

            # 计算主表各项合计
            response_total = district_data['应对入库合计'].sum()
            task_count = len(district_data)

            # 调减收入合计
            reduce_total = district_data['调减收入'].sum() if '调减收入' in district_data.columns else 0

            # 往年入库合计
            previous_total = district_data['往年入库'].sum() if '往年入库' in district_data.columns else 0

            # 往年调减收入合计 - 新增
            previous_reduce_total = district_data[
                '往年调减收入'].sum() if '往年调减收入' in district_data.columns else 0

            # 统计原始为一分局的记录数
            first_bureau_count_in_district = (district_data['所属区局'] == '一分局').sum()  # 修正变量名
            first_bureau_total_in_district = district_data.loc[
                district_data['所属区局'] == '一分局', '应对入库合计'].sum()

            # 获取今年入库表该区局的数据
            if current_grouped is not None and district in current_grouped.groups:
                current_district_data = current_grouped.get_group(district)

                current_rows = len(current_district_data)
                current_in_total = current_district_data[
                    '入库金额'].sum() if '入库金额' in current_district_data.columns else 0
                current_reduce_total = current_district_data[
                    '统计调减收入'].sum() if '统计调减收入' in current_district_data.columns else 0
                current_taxfree_total = current_district_data[
                    '其中：非税收入'].sum() if '其中：非税收入' in current_district_data.columns else 0
                current_total = current_district_data[
                    '今年入库统计总额'].sum() if '今年入库统计总额' in current_district_data.columns else 0
            else:
                current_rows = 0
                current_in_total = 0
                current_reduce_total = 0
                current_taxfree_total = 0
                current_total = 0

            # 添加到汇总数据
            summary_data.append({
                '所属区局': district,
                '应对入库合计': response_total,
                '调减收入合计': reduce_total,
                '往年入库合计': previous_total,
                '往年调减收入合计': previous_reduce_total,  # 新增
                '今年入库金额': current_in_total,
                '今年非税收入': current_taxfree_total,
                '今年调减收入': current_reduce_total,
                '今年入库统计总额': current_total,
                '总任务条数': task_count,
                '今年入库条数': current_rows,
                '一分局分摊记录数': first_bureau_count_in_district,
                '一分局分摊金额': first_bureau_total_in_district
            })

        # 处理不在排序表中的区局（包括'未知区局'）
        remaining_districts = [d for d in all_new_districts if d not in self.district_order]

        if remaining_districts:
            self.log_message(f"注意: 以下区局不在排序表中，将放在最后: {remaining_districts}")

            for district in remaining_districts:
                district_data = new_grouped.get_group(district)

                # 计算主表各项合计
                response_total = district_data['应对入库合计'].sum()
                task_count = len(district_data)

                # 调减收入合计
                reduce_total = district_data['调减收入'].sum() if '调减收入' in district_data.columns else 0

                # 往年入库合计
                previous_total = district_data['往年入库'].sum() if '往年入库' in district_data.columns else 0

                # 往年调减收入合计 - 新增
                previous_reduce_total = district_data[
                    '往年调减收入'].sum() if '往年调减收入' in district_data.columns else 0

                # 统计原始为一分局的记录数
                first_bureau_count_in_district = (district_data['所属区局'] == '一分局').sum()  # 修正变量名
                first_bureau_total_in_district = district_data.loc[
                    district_data['所属区局'] == '一分局', '应对入库合计'].sum()

                # 获取今年入库表该区局的数据
                if current_grouped is not None and district in current_grouped.groups:
                    current_district_data = current_grouped.get_group(district)

                    current_rows = len(current_district_data)
                    current_in_total = current_district_data[
                        '入库金额'].sum() if '入库金额' in current_district_data.columns else 0
                    current_reduce_total = current_district_data[
                        '统计调减收入'].sum() if '统计调减收入' in current_district_data.columns else 0
                    current_taxfree_total = current_district_data[
                        '其中：非税收入'].sum() if '其中：非税收入' in current_district_data.columns else 0
                    current_total = current_district_data[
                        '今年入库统计总额'].sum() if '今年入库统计总额' in current_district_data.columns else 0
                else:
                    current_rows = 0
                    current_in_total = 0
                    current_reduce_total = 0
                    current_taxfree_total = 0
                    current_total = 0

                # 添加到汇总数据
                summary_data.append({
                    '所属区局': district,
                    '应对入库合计': response_total,
                    '调减收入合计': reduce_total,
                    '往年入库合计': previous_total,
                    '往年调减收入合计': previous_reduce_total,  # 新增
                    '今年入库金额': current_in_total,
                    '今年非税收入': current_taxfree_total,
                    '今年调减收入': current_reduce_total,
                    '今年入库统计总额': current_total,
                    '总任务条数': task_count,
                    '今年入库条数': current_rows,
                    '一分局分摊记录数': first_bureau_count_in_district,
                    '一分局分摊金额': first_bureau_total_in_district
                })

        # 创建汇总DataFrame
        summary_df = pd.DataFrame(summary_data)

        # 添加总计行
        total_row = {
            '所属区局': '总计',
            '应对入库合计': summary_df['应对入库合计'].sum(),
            '调减收入合计': summary_df['调减收入合计'].sum(),
            '往年入库合计': summary_df['往年入库合计'].sum(),
            '往年调减收入合计': summary_df['往年调减收入合计'].sum(),  # 新增
            '今年入库金额': summary_df['今年入库金额'].sum(),
            '今年非税收入': summary_df['今年非税收入'].sum(),
            '今年调减收入': summary_df['今年调减收入'].sum(),
            '今年入库统计总额': summary_df['今年入库统计总额'].sum(),
            '总任务条数': summary_df['总任务条数'].sum(),
            '今年入库条数': summary_df['今年入库条数'].sum(),
            '一分局分摊记录数': summary_df['一分局分摊记录数'].sum(),
            '一分局分摊金额': summary_df['一分局分摊金额'].sum()
        }
        summary_df = pd.concat([summary_df, pd.DataFrame([total_row])], ignore_index=True)

        # 优化数据类型 - 确保数值列使用浮点数类型以保留小数
        numeric_cols = [
            '应对入库合计', '调减收入合计', '往年入库合计', '往年调减收入合计',
            '今年入库金额', '今年非税收入', '今年调减收入', '今年入库统计总额',
            '一分局分摊金额'
        ]

        for col in numeric_cols:
            if col in summary_df.columns:
                summary_df[col] = pd.to_numeric(summary_df[col], errors='coerce').astype('float64')

        self.log_message("无一分局分摊表创建完成")
        self.log_message(f"分摊后涉及区局数: {len(all_new_districts)}")
        self.log_message(f"一分局记录总数: {first_bureau_count}")
        self.log_message(f"成功分摊记录数: {first_bureau_count - null_auth_count}")
        self.log_message(f"未识别主管税务机关记录数: {null_auth_count}")

        # 释放内存
        del df_copy

        return summary_df

    def process_data(self):
        """主处理函数，优化内存管理和性能"""
        try:
            self.progress['value'] = 0
            self.status_label.config(text="开始处理数据...")
            self.log_text.delete(1.0, tk.END)  # 清空日志
            self.log_message("=" * 50)
            self.log_message("开始数据处理流程")

            # 检查文件是否选择
            if not hasattr(self, 'main_table_path'):
                messagebox.showerror("错误", "请先选择主表")
                return

            # 1. 加载主表
            self.progress['value'] = 10
            self.status_label.config(text="加载主表...")
            main_df = self.load_and_preprocess_table(self.main_table_path)
            self.log_message(f"主表加载成功，共{len(main_df)}行数据")
            main_df = self.convert_all_numeric_columns(main_df)
            # 2. 处理主表
            self.progress['value'] = 30
            self.status_label.config(text="处理主表...")
            main_df = self.process_main_table(main_df)

            # 3. 合并调减收入和往年入库数据
            self.progress['value'] = 50
            self.status_label.config(text="合并附加数据...")
            main_df = self.merge_additional_data(main_df)

            # 4. 过滤主表，删除不在排序表中的数据，并按照排序表顺序排列
            self.progress['value'] = 60
            self.status_label.config(text="过滤和排序主表数据...")
            main_df = self.filter_and_sort_by_district_order(main_df)

            # 5. 处理今年入库表
            self.progress['value'] = 70
            self.status_label.config(text="处理今年入库表...")
            current_df = self.process_current_table()

            # 6. 过滤和排序今年入库表
            if current_df is not None:
                self.progress['value'] = 75
                self.status_label.config(text="过滤和排序今年入库表...")
                current_df = self.filter_and_sort_by_district_order(current_df)

            # 7. 创建汇总表
            self.progress['value'] = 80
            self.status_label.config(text="创建汇总表...")
            summary_df = self.create_summary_sheet(main_df, current_df, "统计汇总")

            # 8. 创建无一分局分摊表
            self.progress['value'] = 85
            self.status_label.config(text="创建无一分局分摊表...")
            no_first_bureau_df = self.create_no_first_bureau_sheet(main_df, current_df)

            # 9. 在保存之前，删除主表中所有不需要的列
            self.progress['value'] = 90
            self.status_label.config(text="清理和保存结果...")

            # 清理主表
            main_df_columns_to_keep = [
                '任务批次名称', '纳税人名称', '所属区局', '应对入库合计',
                '调减收入', '往年入库'
            ]

            # 确保只保留存在的列
            existing_columns = [col for col in main_df_columns_to_keep if col in main_df.columns]

            # 如果有其他必要的列需要保留，可以在这里添加
            if '主管税务机关' in main_df.columns:
                existing_columns.append('主管税务机关')

            # 删除不需要的列，只保留必要的列
            columns_to_drop = [col for col in main_df.columns if col not in existing_columns]

            if columns_to_drop:
                self.log_message(f"删除主表中不需要的列: {', '.join(columns_to_drop[:10])}..." +
                                 (f"等共{len(columns_to_drop)}列" if len(columns_to_drop) > 10 else ""))
                main_df = main_df[existing_columns]
                self.log_message(f"主表清理完成，保留{len(existing_columns)}列")
            else:
                self.log_message("主表无需清理，所有列都是必要的")
            # 9. 保存结果
            self.progress['value'] = 90
            self.status_label.config(text="保存结果...")

            output_file = filedialog.asksaveasfilename(
                title="保存处理结果",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )

            if output_file:
                # 使用更高效的写入方式
                with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                    main_df.to_excel(writer, sheet_name='新主表', index=False)
                    if current_df is not None:
                        current_df.to_excel(writer, sheet_name='今年入库', index=False)
                    summary_df.to_excel(writer, sheet_name='统计汇总', index=False)
                    if no_first_bureau_df is not None:
                        no_first_bureau_df.to_excel(writer, sheet_name='无一分局分摊', index=False)

                self.progress['value'] = 100
                self.status_label.config(text="处理完成！")
                self.log_message(f"处理完成！结果已保存到: {output_file}")
                messagebox.showinfo("完成", f"数据处理完成！\n结果已保存到: {output_file}")

                # 强制垃圾回收
                import gc
                gc.collect()

        except Exception as e:
            self.log_message(f"处理过程中出现错误: {str(e)}")
            messagebox.showerror("错误", f"处理过程中出现错误:\n{str(e)}")
            self.status_label.config(text="处理失败")
        finally:
            # 确保UI更新
            self.root.update()


if __name__ == "__main__":
    root = tk.Tk()
    app = TaxDataProcessor(root)
    root.mainloop()
