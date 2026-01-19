import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import warnings
import matplotlib
import pandas as pd
import numpy as np

matplotlib.use('TkAgg')  # 使用TkAgg后端确保与Tkinter兼容
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import seaborn as sns
import threading
import traceback

warnings.filterwarnings('ignore')

# 设置matplotlib中文字体
try:
    plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'DejaVu Sans']
    plt.rcParams['axes.unicode_minus'] = False
    print("中文字体设置成功")
except Exception as e:
    print(f"字体设置失败: {e}")

sns.set_palette("husl")


class ModernTaskAllocator:
    def __init__(self, history_df=None, new_tasks_df=None, current_time=None):
        """
        初始化任务分配器
        """
        self.history_df = history_df.copy() if history_df is not None else None
        self.new_tasks_df = new_tasks_df.copy() if new_tasks_df is not None else None
        self.current_time = current_time or datetime.now()

        if history_df is not None and new_tasks_df is not None:
            self._preprocess_data()
            self.personnel_stats = self._calculate_personnel_stats()
            self.dept_stats = self._calculate_department_stats()
            self.system_load = self._calculate_system_load()
        else:
            self.personnel_stats = {}
            self.dept_stats = {}
            self.system_load = {}

    def set_data(self, history_df, new_tasks_df):
        """设置数据并重新计算"""
        self.history_df = history_df.copy()
        self.new_tasks_df = new_tasks_df.copy()
        self._preprocess_data()
        self.personnel_stats = self._calculate_personnel_stats()
        self.dept_stats = self._calculate_department_stats()
        self.system_load = self._calculate_system_load()

    def _preprocess_data(self):
        """数据预处理"""
        # 检查必要字段
        required_cols = ['nsrsbh', 'nsrmc', 'wchj_lz_rwpcmc', 'ydcljg', 'xfsj', 'yjwcsj']
        for col in required_cols:
            if col not in self.history_df.columns:
                raise ValueError(f"历史数据缺少字段: {col}")
            if col not in self.new_tasks_df.columns:
                raise ValueError(f"新任务数据缺少字段: {col}")

        # 转换时间字段
        time_columns = ['xfsj', 'fksj', 'yjwcsj']
        for col in time_columns:
            if col in self.history_df.columns:
                self.history_df[col] = pd.to_datetime(self.history_df[col], errors='coerce')
            if col in self.new_tasks_df.columns:
                self.new_tasks_df[col] = pd.to_datetime(self.new_tasks_df[col], errors='coerce')

        # 处理缺失值
        self.history_df['ydczry_mc'] = self.history_df['ydczry_mc'].fillna('未知')

        # 创建任务唯一标识
        self.history_df['task_id'] = (
                self.history_df['nsrsbh'].astype(str) + '_' +
                self.history_df['wchj_lz_rwpcmc'].astype(str)
        )
        self.new_tasks_df['task_id'] = (
                self.new_tasks_df['nsrsbh'].astype(str) + '_' +
                self.new_tasks_df['wchj_lz_rwpcmc'].astype(str)
        )

        # 计算处理时长（仅对已完成任务）
        mask = ~self.history_df['fksj'].isna()
        if mask.any():
            self.history_df.loc[mask, 'process_duration'] = (
                                                                    self.history_df.loc[mask, 'fksj'] -
                                                                    self.history_df.loc[mask, 'xfsj']
                                                            ).dt.total_seconds() / 3600

        # 标记进行中的任务
        self.history_df['is_ongoing'] = self.history_df['fksj'].isna()

    def _calculate_personnel_stats(self):
        """计算人员统计数据"""
        personnel_stats = {}
        all_personnel = self.history_df['ydczry_mc'].dropna().unique()

        for person in all_personnel:
            try:
                person_data = self.history_df[self.history_df['ydczry_mc'] == person]
                if person_data.empty:
                    continue

                dept = person_data['ydcljg'].iloc[0] if not person_data.empty else '未知'

                # 计算最大并发任务数（基于历史数据的时间窗口分析）
                person_data_sorted = person_data.sort_values('xfsj')
                max_concurrent = 0

                if len(person_data_sorted) > 0:
                    for idx, row in person_data_sorted.iterrows():
                        check_time = row['xfsj']
                        if pd.isna(check_time):
                            continue

                        # 计算在此时刻该人员正在处理的任务数
                        concurrent = person_data_sorted[
                            (person_data_sorted['xfsj'] <= check_time) &
                            (
                                    (person_data_sorted['fksj'].isna()) |
                                    (person_data_sorted['fksj'] > check_time)
                            )
                            ].shape[0]

                        max_concurrent = max(max_concurrent, concurrent)

                # 计算处理效率
                completed_tasks = person_data[~person_data['fksj'].isna()]
                avg_process_time = None
                completion_rate = 0

                if not completed_tasks.empty and 'process_duration' in completed_tasks.columns:
                    durations = completed_tasks['process_duration'].dropna()
                    if not durations.empty:
                        avg_process_time = durations.mean()
                    completion_rate = len(completed_tasks) / len(person_data)

                # 当前任务数
                current_tasks = person_data[person_data['is_ongoing']].shape[0]

                personnel_stats[person] = {
                    'department': dept,
                    'max_concurrent_tasks': max(max_concurrent, 1),
                    'avg_process_time': avg_process_time or 24,
                    'completion_rate': completion_rate,
                    'current_tasks': current_tasks,
                    'historical_tasks_count': len(person_data),
                    'completed_tasks_count': len(completed_tasks)
                }
            except Exception as e:
                print(f"计算人员 {person} 统计信息失败: {e}")
                continue

        return personnel_stats

    def _calculate_department_stats(self):
        """计算机关统计数据（基于时间点的并发任务数）"""
        dept_stats = {}

        for dept in self.history_df['ydcljg'].unique():
            try:
                dept_data = self.history_df[self.history_df['ydcljg'] == dept]
                if dept_data.empty:
                    continue

                dept_personnel = dept_data['ydczry_mc'].unique()

                # 计算机关在历史时间点的最大并发任务数
                dept_data_sorted = dept_data.sort_values('xfsj')
                max_concurrent_dept = 0

                if len(dept_data_sorted) > 0:
                    # 分析每个时间点的并发任务数
                    for idx, row in dept_data_sorted.iterrows():
                        check_time = row['xfsj']
                        if pd.isna(check_time):
                            continue

                        # 计算该时间点该机关正在处理的任务数
                        concurrent_dept = dept_data_sorted[
                            (dept_data_sorted['xfsj'] <= check_time) &
                            (
                                    (dept_data_sorted['fksj'].isna()) |
                                    (dept_data_sorted['fksj'] > check_time)
                            )
                            ].shape[0]

                        max_concurrent_dept = max(max_concurrent_dept, concurrent_dept)

                # 计算机关处理效率
                completed_tasks = dept_data[~dept_data['fksj'].isna()]
                avg_process_time = None
                if not completed_tasks.empty and 'process_duration' in completed_tasks.columns:
                    durations = completed_tasks['process_duration'].dropna()
                    if not durations.empty:
                        avg_process_time = durations.mean()

                # 当前机关未完成任务数
                ongoing_tasks = dept_data[dept_data['is_ongoing']]

                dept_stats[dept] = {
                    'personnel_count': len(dept_personnel),
                    'max_concurrent_tasks': max(max_concurrent_dept, 1),
                    'avg_process_time': avg_process_time or 24,
                    'ongoing_tasks': len(ongoing_tasks),
                    'completion_rate': len(completed_tasks) / len(dept_data) if len(dept_data) > 0 else 0,
                    'personnel_list': list(dept_personnel)
                }
            except Exception as e:
                print(f"计算机关 {dept} 统计信息失败: {e}")
                continue

        return dept_stats

    def _calculate_system_load(self):
        """计算系统整体负载（基于最大并发能力）"""
        # 计算各机关当前负载率
        dept_load = {}
        total_current_tasks = 0
        total_capacity = 0

        for dept, stats in self.dept_stats.items():
            try:
                # 计算机关当前任务数（正在进行中的任务）
                current_tasks = stats['ongoing_tasks']

                # 计算机关最大并发能力
                capacity = stats['max_concurrent_tasks']

                # 计算负载率
                load_percentage = (current_tasks / capacity * 100) if capacity > 0 else 0

                dept_load[dept] = {
                    'load_percentage': load_percentage,
                    'current_tasks': current_tasks,
                    'capacity': capacity
                }

                total_current_tasks += current_tasks
                total_capacity += capacity
            except Exception as e:
                print(f"计算机关 {dept} 负载失败: {e}")
                continue

        # 计算系统整体负载
        system_load_percentage = (total_current_tasks / total_capacity * 100) if total_capacity > 0 else 0

        return {
            'system_load_percentage': system_load_percentage,
            'dept_load': dept_load,
            'total_current_tasks': total_current_tasks,
            'total_capacity': total_capacity
        }

    def _calculate_task_urgency(self, task_row):
        """计算任务紧急程度（0-100分）"""
        try:
            urgency_score = 0

            # 获取规定完成时间
            yjwcsj = task_row['yjwcsj']
            if pd.isna(yjwcsj):
                return 50  # 无规定时间，默认中等紧急

            # 计算剩余时间（小时）
            time_left = (yjwcsj - self.current_time).total_seconds() / 3600

            # 时间紧迫性评分（权重40%）
            if time_left <= 0:
                time_score = 100  # 已逾期
            elif time_left <= 24:
                time_score = 90  # 24小时内
            elif time_left <= 72:
                time_score = 70  # 3天内
            elif time_left <= 168:
                time_score = 40  # 7天内
            else:
                time_score = 20  # 7天以上

            # 机关负载调整（权重30%）
            dept = task_row['ydcljg']
            if dept in self.system_load['dept_load']:
                dept_load_pct = self.system_load['dept_load'][dept]['load_percentage']
                if dept_load_pct > 85:
                    load_score = 100  # 负载很高
                elif dept_load_pct > 70:
                    load_score = 80
                elif dept_load_pct > 50:
                    load_score = 60
                elif dept_load_pct > 30:
                    load_score = 40
                else:
                    load_score = 20
            else:
                load_score = 50

            # 历史延期风险（权重30%）
            # 根据机关历史处理时间和剩余时间比例评估
            if dept in self.dept_stats:
                avg_process_time = self.dept_stats[dept]['avg_process_time']
                remaining_hours = max(time_left, 1)
                risk_ratio = avg_process_time / remaining_hours
                if risk_ratio > 2:
                    risk_score = 100  # 风险很高
                elif risk_ratio > 1.5:
                    risk_score = 80
                elif risk_ratio > 1:
                    risk_score = 60
                elif risk_ratio > 0.5:
                    risk_score = 40
                else:
                    risk_score = 20
            else:
                risk_score = 50

            # 综合评分
            urgency_score = (time_score * 0.4) + (load_score * 0.3) + (risk_score * 0.3)
            return min(100, max(0, urgency_score))

        except Exception as e:
            print(f"计算任务紧急度失败: {e}")
            return 50

    def _calculate_person_suitability(self, person, task_urgency):
        """计算人员适合度（0-100分）"""
        try:
            if person not in self.personnel_stats:
                return 0

            stats = self.personnel_stats[person]

            # 当前负载评分（权重40%）
            current_load = stats['current_tasks']
            max_capacity = stats['max_concurrent_tasks']
            load_ratio = current_load / max_capacity if max_capacity > 0 else 1
            load_score = 100 * (1 - min(load_ratio, 1))

            # 处理效率评分（权重30%）
            avg_time = stats['avg_process_time']
            efficiency_score = max(0, 100 - (avg_time / 48 * 100))  # 假设48小时为基准

            # 任务紧急度匹配（权重30%）
            # 高紧急任务分配给效率高且当前任务少的人
            if task_urgency > 80:  # 高紧急度
                match_score = efficiency_score * 0.7 + load_score * 0.3
            elif task_urgency > 60:  # 中高紧急度
                match_score = efficiency_score * 0.5 + load_score * 0.5
            else:  # 低紧急度，侧重负载均衡
                match_score = load_score * 0.7 + efficiency_score * 0.3

            suitability_score = (load_score * 0.4) + (efficiency_score * 0.3) + (match_score * 0.3)
            return min(100, max(0, suitability_score))

        except Exception as e:
            print(f"计算人员 {person} 适合度失败: {e}")
            return 0

    def allocate_tasks(self):
        """执行任务分配"""
        if self.system_load['system_load_percentage'] > 90:
            print("系统负载过高，暂缓分配")
            return None

        # 计算新任务的紧急程度
        self.new_tasks_df['urgency_score'] = self.new_tasks_df.apply(
            self._calculate_task_urgency, axis=1
        )

        # 按紧急程度排序（紧急的优先分配）
        new_tasks_sorted = self.new_tasks_df.sort_values('urgency_score', ascending=False)

        allocations = []
        unallocated_tasks = []
        total_tasks = len(new_tasks_sorted)

        # 使用itertuples替代iterrows以提升性能
        for idx, task in enumerate(new_tasks_sorted.itertuples(index=False), 1):
            try:
                task_id = task.task_id
                dept = task.ydcljg
                urgency = task.urgency_score

                # 检查机关是否存在
                if dept not in self.dept_stats:
                    unallocated_tasks.append({
                        'task_id': task_id,
                        'reason': f'机关"{dept}"在历史数据中不存在',
                        'urgency_score': urgency
                    })
                    continue

                # 获取机关内人员
                dept_personnel = self.dept_stats[dept]['personnel_list']
                if not dept_personnel:
                    unallocated_tasks.append({
                        'task_id': task_id,
                        'reason': f'机关"{dept}"无可用人员',
                        'urgency_score': urgency
                    })
                    continue

                # 计算每个人员的适合度
                candidate_scores = {}
                for person in dept_personnel:
                    if person not in self.personnel_stats:
                        continue

                    # 检查是否达到最大任务数
                    if self.personnel_stats[person]['current_tasks'] >= \
                            self.personnel_stats[person]['max_concurrent_tasks']:
                        continue

                    suitability = self._calculate_person_suitability(person, urgency)
                    if suitability > 0:
                        candidate_scores[person] = suitability

                if not candidate_scores:
                    unallocated_tasks.append({
                        'task_id': task_id,
                        'reason': f'机关"{dept}"所有人员任务已满或无法匹配',
                        'urgency_score': urgency
                    })
                    continue

                # 选择最适合的人员
                best_person = max(candidate_scores.items(), key=lambda x: x[1])[0]

                # 计算预计完成时间
                avg_process_time = self.personnel_stats[best_person]['avg_process_time']
                estimated_completion = self.current_time + timedelta(hours=avg_process_time)

                # 记录分配
                allocation = {
                    'task_id': task_id,
                    'nsrsbh': task.nsrsbh,
                    'nsrmc': task.nsrmc,
                    'wchj_lz_rwpcmc': task.wchj_lz_rwpcmc,
                    'assigned_department': dept,
                    'assigned_person': best_person,
                    'urgency_score': urgency,
                    'person_suitability_score': candidate_scores[best_person],
                    'estimated_completion': estimated_completion,
                    'current_load': self.personnel_stats[best_person]['current_tasks'],
                    'max_capacity': self.personnel_stats[best_person]['max_concurrent_tasks']
                }

                allocations.append(allocation)

                # 更新人员当前任务数
                self.personnel_stats[best_person]['current_tasks'] += 1

                # 更新机关负载
                if dept in self.system_load['dept_load']:
                    self.system_load['dept_load'][dept]['current_tasks'] += 1
                    self.system_load['dept_load'][dept]['load_percentage'] = (
                            self.system_load['dept_load'][dept]['current_tasks'] /
                            self.system_load['dept_load'][dept]['capacity'] * 100
                    )

                self.system_load['total_current_tasks'] += 1
                self.system_load['system_load_percentage'] = (
                        self.system_load['total_current_tasks'] /
                        self.system_load['total_capacity'] * 100
                )

            except Exception as e:
                print(f"分配任务 {task_id if 'task_id' in locals() else '未知'} 失败: {e}")
                unallocated_tasks.append({
                    'task_id': task_id if 'task_id' in locals() else '未知',
                    'reason': f'分配过程出错: {str(e)}',
                    'urgency_score': urgency if 'urgency' in locals() else 0
                })
                continue

        return {
            'allocations': pd.DataFrame(allocations) if allocations else pd.DataFrame(),
            'unallocated_tasks': pd.DataFrame(unallocated_tasks) if unallocated_tasks else pd.DataFrame(),
            'system_status': self.system_load
        }


class TaskAllocatorUI:
    def __init__(self, root):
        self.root = root
        self.root.title("智能任务分配系统")
        self.root.geometry("1400x800")

        # 设置样式
        self.setup_styles()

        # 初始化分配器
        self.allocator = None
        self.history_df = None
        self.new_tasks_df = None
        self.result = None

        # 创建UI
        self.create_widgets()

        # 延迟加载示例数据，避免阻塞UI
        self.root.after(100, self.load_sample_data_async)

        # 确保窗口显示
        self.root.update()

    def setup_styles(self):
        """设置样式"""
        style = ttk.Style()
        style.theme_use('clam')

        # 自定义颜色
        self.bg_color = "#f5f7fa"
        self.sidebar_color = "#2c3e50"
        self.accent_color = "#3498db"
        self.success_color = "#27ae60"
        self.warning_color = "#f39c12"
        self.danger_color = "#e74c3c"

        # 配置样式
        style.configure("TFrame", background=self.bg_color)
        style.configure("TLabel", background=self.bg_color, font=("微软雅黑", 10))
        style.configure("Title.TLabel", font=("微软雅黑", 16, "bold"))
        style.configure("Header.TLabel", font=("微软雅黑", 12, "bold"))
        style.configure("TButton", padding=6)

        self.root.configure(bg=self.bg_color)

    def create_widgets(self):
        """创建所有UI部件"""
        # 主容器
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 侧边栏
        self.create_sidebar(main_container)

        # 主内容区
        self.create_main_content(main_container)

    def create_sidebar(self, parent):
        """创建侧边栏"""
        sidebar = ttk.Frame(parent, width=200)
        sidebar.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        sidebar.pack_propagate(False)

        # 标题
        title_label = ttk.Label(sidebar, text="任务分配系统",
                                style="Title.TLabel")
        title_label.pack(pady=20)

        # 按钮区域
        btn_frame = ttk.Frame(sidebar)
        btn_frame.pack(pady=20, padx=10)

        # 按钮列表
        buttons = [
            ("加载历史数据", self.load_history_data),
            ("加载新任务", self.load_new_tasks),
            ("数据预览", self.preview_data),
            ("智能分配", self.run_allocation),
            ("生成报告", self.generate_report),
            ("导出结果", self.export_results),
            ("重置系统", self.reset_system),
            ("使用帮助", self.show_help)
        ]

        for text, command in buttons:
            btn = ttk.Button(btn_frame, text=text, command=command)
            btn.pack(fill=tk.X, pady=5)

        # 状态显示
        status_frame = ttk.Frame(sidebar)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=10)

        self.status_label = ttk.Label(status_frame, text="就绪")
        self.status_label.pack()

    def create_main_content(self, parent):
        """创建主内容区"""
        notebook = ttk.Notebook(parent)
        notebook.pack(fill=tk.BOTH, expand=True)

        # 仪表板标签页
        self.create_dashboard_tab(notebook)

        # 分配结果标签页
        self.create_allocation_tab(notebook)

        # 数据分析标签页
        self.create_analysis_tab(notebook)

        # 系统设置标签页
        self.create_settings_tab(notebook)

    def create_dashboard_tab(self, notebook):
        """创建仪表板标签页"""
        dashboard_frame = ttk.Frame(notebook)
        notebook.add(dashboard_frame, text="仪表板")

        # 顶部状态卡片
        self.create_status_cards(dashboard_frame)

        # 图表区域
        chart_frame = ttk.Frame(dashboard_frame)
        chart_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # 左图表
        left_chart_frame = ttk.Frame(chart_frame)
        left_chart_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        self.load_chart_canvas = tk.Canvas(left_chart_frame, bg="white", height=300)
        self.load_chart_canvas.pack(fill=tk.BOTH, expand=True)

        # 右图表
        right_chart_frame = ttk.Frame(chart_frame)
        right_chart_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))

        self.efficiency_chart_canvas = tk.Canvas(right_chart_frame, bg="white", height=300)
        self.efficiency_chart_canvas.pack(fill=tk.BOTH, expand=True)

        # 底部信息
        info_frame = ttk.Frame(dashboard_frame)
        info_frame.pack(fill=tk.X, pady=10)

        self.info_text = scrolledtext.ScrolledText(info_frame, height=10,
                                                   font=("微软雅黑", 10))
        self.info_text.pack(fill=tk.BOTH, expand=True)

    def create_status_cards(self, parent):
        """创建状态卡片"""
        cards_frame = ttk.Frame(parent)
        cards_frame.pack(fill=tk.X, pady=(0, 10))

        # 卡片数据
        self.cards_data = {
            "system_load": {"title": "系统负载", "value": "0%", "color": self.success_color},
            "total_tasks": {"title": "总任务数", "value": "0", "color": self.accent_color},
            "allocated": {"title": "已分配", "value": "0", "color": self.success_color},
            "pending": {"title": "待分配", "value": "0", "color": self.warning_color},
            "personnel": {"title": "处理人员", "value": "0", "color": self.accent_color},
            "departments": {"title": "机关数量", "value": "0", "color": self.accent_color}
        }

        self.cards = {}
        for i, (key, data) in enumerate(self.cards_data.items()):
            card = tk.Frame(cards_frame, bg="white", relief=tk.RAISED, bd=1)
            card.grid(row=0, column=i, padx=5, pady=5, sticky="nsew")
            cards_frame.grid_columnconfigure(i, weight=1)

            # 标题
            title_label = tk.Label(card, text=data["title"], bg="white",
                                   font=("微软雅黑", 10))
            title_label.pack(pady=(10, 0))

            # 值
            value_label = tk.Label(card, text=data["value"], bg="white",
                                   font=("微软雅黑", 24, "bold"), fg=data["color"])
            value_label.pack(pady=(5, 10))

            self.cards[key] = value_label

    def create_allocation_tab(self, notebook):
        """创建分配结果标签页"""
        allocation_frame = ttk.Frame(notebook)
        notebook.add(allocation_frame, text="分配结果")

        # 控制按钮
        control_frame = ttk.Frame(allocation_frame)
        control_frame.pack(fill=tk.X, pady=5)

        ttk.Button(control_frame, text="执行分配",
                   command=self.run_allocation).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="导出Excel",
                   command=self.export_results).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="清空结果",
                   command=self.clear_results).pack(side=tk.LEFT, padx=5)

        # 分配结果表格
        self.create_allocation_table(allocation_frame)

    def create_allocation_table(self, parent):
        """创建分配结果表格"""
        # 创建Treeview
        columns = ("任务ID", "户名", "机关", "分配人员", "紧急度", "适合度", "预计完成")
        self.allocation_tree = ttk.Treeview(parent, columns=columns, show="headings", height=15)

        # 设置列标题
        for col in columns:
            self.allocation_tree.heading(col, text=col)
            self.allocation_tree.column(col, width=100)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL,
                                  command=self.allocation_tree.yview)
        self.allocation_tree.configure(yscrollcommand=scrollbar.set)

        # 布局
        self.allocation_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 详情显示
        detail_frame = ttk.Frame(parent)
        detail_frame.pack(fill=tk.X, pady=10)

        self.detail_text = scrolledtext.ScrolledText(detail_frame, height=8,
                                                     font=("微软雅黑", 10))
        self.detail_text.pack(fill=tk.BOTH, expand=True)

    def create_analysis_tab(self, notebook):
        """创建数据分析标签页"""
        analysis_frame = ttk.Frame(notebook)
        notebook.add(analysis_frame, text="数据分析")

        # 分析选项
        options_frame = ttk.Frame(analysis_frame)
        options_frame.pack(fill=tk.X, pady=5)

        self.analysis_var = tk.StringVar(value="load_distribution")
        analyses = [
            ("负载分布", "load_distribution"),
            ("机关效率", "dept_efficiency"),  # 改为机关效率
            ("人员配置", "personnel_config"),  # 改为人员配置和负载
            ("时间趋势", "time_trend")
        ]

        for text, value in analyses:
            ttk.Radiobutton(options_frame, text=text, variable=self.analysis_var,
                            value=value, command=self.update_analysis).pack(side=tk.LEFT, padx=10)

        # 图表区域
        self.analysis_canvas = tk.Canvas(analysis_frame, bg="white", height=400)
        self.analysis_canvas.pack(fill=tk.BOTH, expand=True, pady=10)

    def create_settings_tab(self, notebook):
        """创建系统设置标签页"""
        settings_frame = ttk.Frame(notebook)
        notebook.add(settings_frame, text="系统设置")

        # 参数设置
        params_frame = ttk.LabelFrame(settings_frame, text="分配参数", padding=10)
        params_frame.pack(fill=tk.X, pady=5, padx=10)

        # 负载阈值
        ttk.Label(params_frame, text="系统负载阈值(%):").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.load_threshold_var = tk.StringVar(value="90")
        ttk.Entry(params_frame, textvariable=self.load_threshold_var, width=10).grid(row=0, column=1, pady=5)

        # 权重设置
        ttk.Label(params_frame, text="紧急度权重(时间:负载:风险):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.weights_var = tk.StringVar(value="0.4:0.3:0.3")
        ttk.Entry(params_frame, textvariable=self.weights_var, width=15).grid(row=1, column=1, pady=5)

        # 时间阈值
        ttk.Label(params_frame, text="紧急任务阈值(小时):").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.urgent_threshold_var = tk.StringVar(value="24")
        ttk.Entry(params_frame, textvariable=self.urgent_threshold_var, width=10).grid(row=2, column=1, pady=5)

        # 保存设置按钮
        ttk.Button(params_frame, text="保存设置",
                   command=self.save_settings).grid(row=3, column=0, columnspan=2, pady=10)

        # 系统信息
        info_frame = ttk.LabelFrame(settings_frame, text="系统信息", padding=10)
        info_frame.pack(fill=tk.X, pady=5, padx=10)

        ttk.Label(info_frame, text="版本: 2.0").pack(anchor=tk.W)
        ttk.Label(info_frame, text="开发者: 智能分配系统").pack(anchor=tk.W)
        ttk.Label(info_frame, text="更新时间: 2024").pack(anchor=tk.W)

    def load_sample_data(self):
        """加载示例数据"""
        try:
            # 创建示例数据
            history_data = {
                'nsrsbh': ['001', '002', '003', '004', '005'],
                'nsrmc': ['公司A', '公司B', '公司C', '公司D', '公司E'],
                'wchj_lz_rwpcmc': ['任务1', '任务2', '任务3', '任务4', '任务5'],
                'ydczry_mc': ['张三', '李四', '张三', '王五', '李四'],
                'ydcljg': ['机关1', '机关1', '机关2', '机关2', '机关1'],
                'xfsj': ['2023-01-01', '2023-01-02', '2023-01-03', '2023-01-04', '2023-01-05'],
                'fksj': ['2023-01-02', None, '2023-01-05', None, None],
                'yjwcsj': ['2023-01-03', '2023-01-10', '2023-01-06', '2023-01-08', '2023-01-12']
            }

            new_tasks_data = {
                'nsrsbh': ['006', '007', '008'],
                'nsrmc': ['公司F', '公司G', '公司H'],
                'wchj_lz_rwpcmc': ['任务6', '任务7', '任务8'],
                'ydczry_mc': [None, None, None],
                'ydcljg': ['机关1', '机关2', '机关1'],
                'xfsj': ['2023-01-06', '2023-01-06', '2023-01-06'],
                'fksj': [None, None, None],
                'yjwcsj': ['2023-01-07', '2023-01-09', '2023-01-15']
            }

            self.history_df = pd.DataFrame(history_data)
            self.new_tasks_df = pd.DataFrame(new_tasks_data)

            self.update_status("示例数据加载成功！", "success")
        except Exception as e:
            self.update_status(f"加载示例数据失败: {str(e)}", "error")

    def load_sample_data_async(self):
        """异步加载示例数据，避免阻塞UI"""
        def load_data():
            try:
                # 创建示例数据
                history_data = {
                    'nsrsbh': ['001', '002', '003', '004', '005'],
                    'nsrmc': ['公司A', '公司B', '公司C', '公司D', '公司E'],
                    'wchj_lz_rwpcmc': ['任务1', '任务2', '任务3', '任务4', '任务5'],
                    'ydczry_mc': ['张三', '李四', '张三', '王五', '李四'],
                    'ydcljg': ['机关1', '机关1', '机关2', '机关2', '机关1'],
                    'xfsj': ['2023-01-01', '2023-01-02', '2023-01-03', '2023-01-04', '2023-01-05'],
                    'fksj': ['2023-01-02', None, '2023-01-05', None, None],
                    'yjwcsj': ['2023-01-03', '2023-01-10', '2023-01-06', '2023-01-08', '2023-01-12']
                }

                new_tasks_data = {
                    'nsrsbh': ['006', '007', '008'],
                    'nsrmc': ['公司F', '公司G', '公司H'],
                    'wchj_lz_rwpcmc': ['任务6', '任务7', '任务8'],
                    'ydczry_mc': [None, None, None],
                    'ydcljg': ['机关1', '机关2', '机关1'],
                    'xfsj': ['2023-01-06', '2023-01-06', '2023-01-06'],
                    'fksj': [None, None, None],
                    'yjwcsj': ['2023-01-07', '2023-01-09', '2023-01-15']
                }

                self.history_df = pd.DataFrame(history_data)
                self.new_tasks_df = pd.DataFrame(new_tasks_data)

                # 在主线程更新UI
                self.root.after(0, lambda: self.update_status("示例数据加载成功！", "success"))
            except Exception as e:
                self.root.after(0, lambda: self.update_status(f"加载示例数据失败: {str(e)}", "error"))

        # 在后台线程中加载数据
        thread = threading.Thread(target=load_data, daemon=True)
        thread.start()

    def load_history_data(self):
        """加载历史数据"""
        file_path = filedialog.askopenfilename(
            title="选择历史任务表",
            filetypes=[("Excel文件", "*.xlsx;*.xls"), ("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )

        if file_path:
            try:
                if file_path.endswith('.csv'):
                    df = pd.read_csv(file_path, encoding='gbk')
                else:
                    df = pd.read_excel(file_path)

                # 检查必要字段
                required_fields = ['nsrsbh', 'nsrmc', 'wchj_lz_rwpcmc',
                                   'ydczry_mc', 'ydcljg', 'xfsj', 'fksj', 'yjwcsj']

                missing_fields = [field for field in required_fields if field not in df.columns]
                if missing_fields:
                    messagebox.showerror("错误", f"缺少必要字段: {', '.join(missing_fields)}")
                    return

                self.history_df = df
                self.update_status(f"历史数据加载成功: {len(df)} 条记录", "success")
            except Exception as e:
                messagebox.showerror("错误", f"加载文件失败: {str(e)}")

    def load_new_tasks(self):
        """加载新任务数据"""
        file_path = filedialog.askopenfilename(
            title="选择新任务表",
            filetypes=[("Excel文件", "*.xlsx;*.xls"), ("CSV文件", "*.csv"), ("所有文件", "*.*")]
        )

        if file_path:
            try:
                if file_path.endswith('.csv'):
                    df = pd.read_csv(file_path, encoding='gbk')
                else:
                    df = pd.read_excel(file_path)

                required_fields = ['nsrsbh', 'nsrmc', 'wchj_lz_rwpcmc',
                                   'ydcljg', 'xfsj', 'yjwcsj']

                missing_fields = [field for field in required_fields if field not in df.columns]
                if missing_fields:
                    messagebox.showerror("错误", f"缺少必要字段: {', '.join(missing_fields)}")
                    return

                self.new_tasks_df = df
                self.update_status(f"新任务加载成功: {len(df)} 条记录", "success")
            except Exception as e:
                messagebox.showerror("错误", f"加载文件失败: {str(e)}")

    def preview_data(self):
        """预览数据"""
        if self.history_df is None or self.new_tasks_df is None:
            messagebox.showwarning("提示", "请先加载数据")
            return

        preview_window = tk.Toplevel(self.root)
        preview_window.title("数据预览")
        preview_window.geometry("1000x600")

        notebook = ttk.Notebook(preview_window)
        notebook.pack(fill=tk.BOTH, expand=True)

        # 历史数据预览
        history_frame = ttk.Frame(notebook)
        notebook.add(history_frame, text="历史数据")

        history_tree = self.create_preview_tree(history_frame, self.history_df)

        # 新任务预览
        new_frame = ttk.Frame(notebook)
        notebook.add(new_frame, text="新任务")

        new_tree = self.create_preview_tree(new_frame, self.new_tasks_df)

    def create_preview_tree(self, parent, df):
        """创建数据预览表格"""
        columns = list(df.columns)
        tree = ttk.Treeview(parent, columns=columns, show="headings", height=20)

        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=100)

        # 插入数据（限制行数，避免卡顿）
        for _, row in df.head(100).iterrows():
            tree.insert("", tk.END, values=list(row))

        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)

        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        return tree

    def run_allocation(self):
        """执行分配"""
        if self.history_df is None or self.new_tasks_df is None:
            messagebox.showwarning("提示", "请先加载历史数据和新任务数据")
            return

        # 在后台线程中执行分配
        def allocate():
            self.update_status("正在执行智能分配...", "info")
            try:
                # 显示处理进度
                self.info_text.delete(1.0, tk.END)
                self.info_text.insert(1.0, "开始数据预处理...\n")
                self.info_text.see(tk.END)
                self.root.update()

                # 创建分配器
                self.allocator = ModernTaskAllocator(self.history_df, self.new_tasks_df)

                self.info_text.insert(tk.END, "数据预处理完成\n")
                self.info_text.insert(tk.END, f"处理人员数量: {len(self.allocator.personnel_stats)}\n")
                self.info_text.insert(tk.END, f"机关数量: {len(self.allocator.dept_stats)}\n")
                self.info_text.insert(tk.END,
                                      f"系统负载率: {self.allocator.system_load['system_load_percentage']:.1f}%\n")

                # 执行分配
                self.info_text.insert(tk.END, "开始任务分配...\n")
                self.root.update()

                self.result = self.allocator.allocate_tasks()

                if self.result:
                    self.update_status("分配完成！", "success")
                    self.display_results()
                    self.update_dashboard()
                    self.update_analysis()

                    self.info_text.insert(tk.END, f"\n分配完成！成功分配 {len(self.result['allocations'])} 个任务\n")
                    if not self.result['unallocated_tasks'].empty:
                        self.info_text.insert(tk.END, f"未分配任务: {len(self.result['unallocated_tasks'])} 个\n")
                else:
                    self.update_status("系统负载过高，建议暂缓分配", "warning")
                    self.info_text.insert(tk.END, "分配失败：系统负载过高\n")

            except Exception as e:
                self.update_status(f"分配失败: {str(e)}", "error")
                error_msg = f"分配过程出现错误:\n{str(e)}\n\n详细错误信息:\n"
                error_msg += traceback.format_exc()
                self.info_text.delete(1.0, tk.END)
                self.info_text.insert(1.0, error_msg)

        # 启动线程
        thread = threading.Thread(target=allocate)
        thread.daemon = True
        thread.start()

    def display_results(self):
        """显示分配结果"""
        if not self.result:
            return

        # 清空表格
        for item in self.allocation_tree.get_children():
            self.allocation_tree.delete(item)

        allocations_df = self.result['allocations']

        if not allocations_df.empty:
            for _, row in allocations_df.iterrows():
                values = (
                    row['task_id'],
                    row['nsrmc'],
                    row['assigned_department'],
                    row['assigned_person'],
                    f"{row['urgency_score']:.1f}",
                    f"{row['person_suitability_score']:.1f}",
                    row['estimated_completion'].strftime("%Y-%m-%d %H:%M")
                )
                self.allocation_tree.insert("", tk.END, values=values)

        # 更新详情
        detail_text = ""
        if not self.result['unallocated_tasks'].empty:
            detail_text += "未分配任务:\n"
            for _, row in self.result['unallocated_tasks'].iterrows():
                detail_text += f"  - {row['task_id']}: {row['reason']}\n"
            detail_text += "\n"

        if not allocations_df.empty:
            detail_text += "分配统计:\n"
            dept_summary = allocations_df['assigned_department'].value_counts()
            for dept, count in dept_summary.items():
                detail_text += f"  {dept}: {count}个任务\n"

            detail_text += f"\n任务紧急度分布:\n"
            urgency_stats = allocations_df['urgency_score'].describe()
            detail_text += f"  最高: {urgency_stats['max']:.1f}\n"
            detail_text += f"  平均: {urgency_stats['mean']:.1f}\n"
            detail_text += f"  最低: {urgency_stats['min']:.1f}\n"

        self.detail_text.delete(1.0, tk.END)
        self.detail_text.insert(1.0, detail_text)

    def update_dashboard(self):
        """更新仪表板"""
        if not self.allocator or not self.result:
            return

        # 更新卡片
        load_pct = self.allocator.system_load['system_load_percentage']
        self.cards["system_load"].config(text=f"{load_pct:.1f}%")

        total_tasks = len(self.history_df) + len(self.new_tasks_df)
        self.cards["total_tasks"].config(text=str(total_tasks))

        allocated = len(self.result['allocations']) if not self.result['allocations'].empty else 0
        self.cards["allocated"].config(text=str(allocated))

        pending = len(self.result['unallocated_tasks']) if not self.result['unallocated_tasks'].empty else 0
        self.cards["pending"].config(text=str(pending))

        personnel = len(self.allocator.personnel_stats)
        self.cards["personnel"].config(text=str(personnel))

        departments = len(self.allocator.dept_stats)
        self.cards["departments"].config(text=str(departments))

        # 更新图表
        self.update_charts()

        # 更新信息文本
        info_text = f"""系统状态报告 ({datetime.now().strftime('%Y-%m-%d %H:%M')})

系统负载: {load_pct:.1f}%
总任务数: {total_tasks}
成功分配: {allocated}
待处理: {pending}

处理人员: {personnel}人
机关数量: {departments}个

分配完成时间: {datetime.now().strftime('%H:%M:%S')}
"""

        self.info_text.delete(1.0, tk.END)
        self.info_text.insert(1.0, info_text)

    def update_charts(self):
        """更新图表"""
        try:
            # 先清除旧的图表
            for widget in self.load_chart_canvas.winfo_children():
                widget.destroy()
            for widget in self.efficiency_chart_canvas.winfo_children():
                widget.destroy()

            # 负载分布图
            fig1, ax1 = plt.subplots(figsize=(5, 3))
            depts = list(self.allocator.system_load['dept_load'].keys())

            if depts:
                loads = [self.allocator.system_load['dept_load'][d]['load_percentage'] for d in depts]

                colors = []
                for load in loads:
                    if load < 60:
                        colors.append(self.success_color)
                    elif load < 85:
                        colors.append(self.warning_color)
                    else:
                        colors.append(self.danger_color)

                bars = ax1.bar(depts, loads, color=colors)
                ax1.set_title('各机关负载情况')
                ax1.set_ylabel('负载率 (%)')
                ax1.set_ylim(0, 100)
                ax1.axhline(y=90, color='red', linestyle='--', alpha=0.5, label='警戒线')
                ax1.legend()

                # 添加数值标签
                for bar in bars:
                    height = bar.get_height()
                    ax1.text(bar.get_x() + bar.get_width() / 2., height + 1,
                             f'{height:.1f}%', ha='center', va='bottom', fontsize=8)
            else:
                ax1.text(0.5, 0.5, '暂无负载数据', ha='center', va='center')

            plt.tight_layout()

            # 在Canvas上绘制
            canvas1 = FigureCanvasTkAgg(fig1, self.load_chart_canvas)
            canvas1.draw()
            canvas1.get_tk_widget().pack(fill=tk.BOTH, expand=True)

            # 效率分布图
            fig2, ax2 = plt.subplots(figsize=(5, 3))
            personnel_names = list(self.allocator.personnel_stats.keys())[:8]  # 只显示前8个

            if personnel_names:
                efficiency_scores = [self.allocator.personnel_stats[p]['completion_rate'] * 100
                                     for p in personnel_names]

                ax2.barh(personnel_names, efficiency_scores, color=self.accent_color)
                ax2.set_xlabel('完成率 (%)')
                ax2.set_title('人员效率排名')

                # 添加数值标签
                for i, v in enumerate(efficiency_scores):
                    ax2.text(v + 1, i, f'{v:.1f}%', va='center')

            plt.tight_layout()

            canvas2 = FigureCanvasTkAgg(fig2, self.efficiency_chart_canvas)
            canvas2.draw()
            canvas2.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        except Exception as e:
            print(f"更新图表失败: {e}")
            # 在画布上显示错误
            self.load_chart_canvas.delete("all")
            self.load_chart_canvas.create_text(150, 150, text="图表生成失败", fill="red")
            self.efficiency_chart_canvas.delete("all")
            self.efficiency_chart_canvas.create_text(150, 150, text="图表生成失败", fill="red")

    def update_analysis(self):
        """更新分析图表"""
        if not self.allocator:
            return

        analysis_type = self.analysis_var.get()

        try:
            # 清除旧的图表
            for widget in self.analysis_canvas.winfo_children():
                widget.destroy()

            fig, ax = plt.subplots(figsize=(8, 5))

            if analysis_type == "load_distribution":
                # 负载分布
                depts = list(self.allocator.system_load['dept_load'].keys())
                if depts:
                    loads = [self.allocator.system_load['dept_load'][d]['load_percentage'] for d in depts]

                    # 创建颜色映射
                    colors = plt.cm.RdYlGn_r(np.array(loads) / 100)
                    bars = ax.bar(depts, loads, color=colors)

                    ax.set_title('机关负载分布')
                    ax.set_ylabel('负载率 (%)')
                    ax.set_ylim(0, 100)
                    ax.axhline(y=90, color='red', linestyle='--', alpha=0.7, label='警戒线')
                    ax.legend()

                    # 添加数值标签
                    for bar in bars:
                        height = bar.get_height()
                        ax.text(bar.get_x() + bar.get_width() / 2., height + 1,
                                f'{height:.1f}%', ha='center', va='bottom')
                else:
                    ax.text(0.5, 0.5, '暂无机关负载数据', ha='center', va='center')

            elif analysis_type == "dept_efficiency":
                """机关处理效率分析"""
                depts = list(self.allocator.dept_stats.keys())
                if depts:
                    # 获取机关处理时间（转换为天）
                    avg_process_days = []
                    completion_rates = []

                    for dept in depts:
                        stats = self.allocator.dept_stats[dept]
                        # 将小时转换为天
                        avg_days = stats['avg_process_time'] / 24  # 小时转换为天
                        avg_process_days.append(avg_days)
                        completion_rates.append(stats['completion_rate'] * 100)

                    # 只显示前10个机关
                    display_count = min(10, len(depts))
                    depts_display = depts[:display_count]
                    avg_days_display = avg_process_days[:display_count]
                    completion_rates_display = completion_rates[:display_count]

                    x = np.arange(len(depts_display))
                    width = 0.35

                    # 创建双轴图表
                    ax1 = ax.twinx()

                    # 柱状图：平均处理时间（天）
                    bars1 = ax.bar(x - width / 2, avg_days_display, width,
                                   label='平均处理时间(天)', color='skyblue')

                    # 折线图：完成率
                    line1 = ax1.plot(x + width / 2, completion_rates_display,
                                     marker='o', color='orange',
                                     label='完成率(%)', linewidth=2)

                    ax.set_xlabel('机关名称')
                    ax.set_ylabel('平均处理时间 (天)', color='skyblue')
                    ax1.set_ylabel('完成率 (%)', color='orange')
                    ax.set_xticks(x)
                    ax.set_xticklabels(depts_display, rotation=45, ha='right')
                    ax.set_title('机关处理效率分析')

                    # 设置y轴范围
                    ax.set_ylim(0, max(avg_days_display) * 1.2 if avg_days_display else 10)
                    ax1.set_ylim(0, 110)

                    # 合并图例
                    bars1_line = [bars1]
                    lines1 = ax.get_legend_handles_labels()[0]
                    lines2 = ax1.get_legend_handles_labels()[0]
                    ax.legend(lines1 + lines2, ['平均处理时间(天)', '完成率(%)'], loc='upper right')

                    # 添加数值标签
                    for bar in bars1:
                        height = bar.get_height()
                        ax.text(bar.get_x() + bar.get_width() / 2., height + 0.1,
                                f'{height:.1f}天', ha='center', va='bottom', fontsize=9)

                    for i, v in enumerate(completion_rates_display):
                        ax1.text(i + width / 2, v + 2, f'{v:.1f}%', ha='center', va='bottom', fontsize=9)
                else:
                    ax.text(0.5, 0.5, '暂无机关效率数据', ha='center', va='center')

            elif analysis_type == "personnel_config":
                """各机关人员数量及当前负载分析"""
                depts = list(self.allocator.dept_stats.keys())
                if depts:
                    # 获取各机关的人员数量和平均负载
                    personnel_counts = []
                    avg_loads = []

                    for dept in depts:
                        stats = self.allocator.dept_stats[dept]
                        personnel_counts.append(stats['personnel_count'])

                        # 计算该机关的平均人员负载
                        if dept in self.allocator.system_load['dept_load']:
                            dept_load = self.allocator.system_load['dept_load'][dept]
                            # 当前负载率
                            avg_loads.append(dept_load['load_percentage'])
                        else:
                            avg_loads.append(0)

                    # 只显示前8个机关
                    display_count = min(8, len(depts))
                    depts_display = depts[:display_count]
                    personnel_display = personnel_counts[:display_count]
                    avg_loads_display = avg_loads[:display_count]

                    x = np.arange(len(depts_display))
                    width = 0.35

                    # 创建双轴图表
                    ax1 = ax.twinx()

                    # 柱状图：人员数量
                    bars1 = ax.bar(x - width / 2, personnel_display, width,
                                   label='人员数量', color='lightgreen')

                    # 折线图：平均负载率
                    line1 = ax1.plot(x + width / 2, avg_loads_display,
                                     marker='s', color='coral',
                                     label='平均负载率(%)', linewidth=2)

                    ax.set_xlabel('机关名称')
                    ax.set_ylabel('人员数量', color='lightgreen')
                    ax1.set_ylabel('平均负载率 (%)', color='coral')
                    ax.set_xticks(x)
                    ax.set_xticklabels(depts_display, rotation=45, ha='right')
                    ax.set_title('机关人员配置及负载分析')

                    # 设置y轴范围
                    if personnel_display:
                        ax.set_ylim(0, max(personnel_display) * 1.2)
                    ax1.set_ylim(0, 110)

                    # 合并图例
                    bars1_line = [bars1]
                    lines1 = ax.get_legend_handles_labels()[0]
                    lines2 = ax1.get_legend_handles_labels()[0]
                    ax.legend(lines1 + lines2, ['人员数量', '平均负载率(%)'], loc='upper right')

                    # 添加数值标签
                    for bar in bars1:
                        height = bar.get_height()
                        ax.text(bar.get_x() + bar.get_width() / 2., height + 0.1,
                                f'{int(height)}人', ha='center', va='bottom', fontsize=9)

                    for i, v in enumerate(avg_loads_display):
                        ax1.text(i + width / 2, v + 2, f'{v:.1f}%', ha='center', va='bottom', fontsize=9)
                else:
                    ax.text(0.5, 0.5, '暂无机关人员配置数据', ha='center', va='center')

            elif analysis_type == "time_trend":
                """时间趋势分析"""
                if self.history_df is not None and 'xfsj' in self.history_df.columns:
                    df_time = self.history_df.copy()

                    # 确保xfsj是datetime类型
                    if not pd.api.types.is_datetime64_any_dtype(df_time['xfsj']):
                        df_time['xfsj'] = pd.to_datetime(df_time['xfsj'], errors='coerce')

                    # 移除无效的时间数据
                    df_time = df_time.dropna(subset=['xfsj'])

                    if not df_time.empty:
                        # 按月统计任务数量
                        df_time['month'] = df_time['xfsj'].dt.to_period('M')
                        monthly_counts = df_time.groupby('month').size()

                        if not monthly_counts.empty:
                            # 转换Period为字符串用于显示
                            month_labels = [str(m) for m in monthly_counts.index]

                            ax.plot(range(len(monthly_counts)), monthly_counts.values,
                                    marker='o', color='steelblue', linewidth=2)
                            ax.set_xlabel('月份')
                            ax.set_ylabel('任务数量')
                            ax.set_title('月度任务趋势')
                            ax.set_xticks(range(len(monthly_counts)))
                            ax.set_xticklabels(month_labels, rotation=45)

                            # 添加数值标签
                            for i, v in enumerate(monthly_counts.values):
                                ax.text(i, v + 0.5, str(v), ha='center', va='bottom')

                            # 添加趋势线
                            if len(monthly_counts) > 1:
                                z = np.polyfit(range(len(monthly_counts)), monthly_counts.values, 1)
                                p = np.poly1d(z)
                                ax.plot(range(len(monthly_counts)), p(range(len(monthly_counts))),
                                        "r--", alpha=0.5, label='趋势线')
                                ax.legend()
                        else:
                            ax.text(0.5, 0.5, '暂无时间趋势数据', ha='center', va='center')
                    else:
                        ax.text(0.5, 0.5, '时间数据格式错误或缺失', ha='center', va='center')
                else:
                    ax.text(0.5, 0.5, '暂无时间趋势数据', ha='center', va='center')

            plt.tight_layout()

            # 更新Canvas
            canvas = FigureCanvasTkAgg(fig, self.analysis_canvas)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        except Exception as e:
            print(f"更新分析图表失败: {e}")
            import traceback
            traceback.print_exc()  # 打印详细错误信息

            # 在画布上显示错误信息
            self.analysis_canvas.delete("all")
            self.analysis_canvas.create_text(200, 150,
                                             text=f"图表生成失败:\n{str(e)[:50]}...",
                                             fill="red", font=("Arial", 12))

    def generate_report(self):
        """生成报告"""
        if not self.allocator or not self.result:
            messagebox.showwarning("提示", "请先执行分配")
            return

        report_window = tk.Toplevel(self.root)
        report_window.title("分配报告")
        report_window.geometry("800x600")

        report_text = scrolledtext.ScrolledText(report_window, font=("微软雅黑", 11))
        report_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 生成报告内容
        report_content = self._generate_report_content()
        report_text.insert(1.0, report_content)

    def _generate_report_content(self):
        """生成报告内容"""
        content = f"""智能任务分配系统报告
{'=' * 50}

报告时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
系统版本: 2.0

一、系统状态概览
系统负载率: {self.allocator.system_load['system_load_percentage']:.1f}%
当前总任务数: {self.allocator.system_load['total_current_tasks']}
系统总容量: {self.allocator.system_load['total_capacity']}

二、各机关负载情况
"""

        for dept, load_info in self.allocator.system_load['dept_load'].items():
            status = "正常" if load_info['load_percentage'] < 60 else \
                "繁忙" if load_info['load_percentage'] < 85 else "过载"
            content += f"  {dept}: {load_info['load_percentage']:.1f}% ({status}) - {load_info['current_tasks']}/{load_info['capacity']}任务\n"

        if not self.result['allocations'].empty:
            content += f"\n三、分配结果统计\n成功分配任务数: {len(self.result['allocations'])}\n"

            allocations_df = self.result['allocations']
            dept_summary = allocations_df.groupby('assigned_department').agg({
                'urgency_score': 'mean',
                'person_suitability_score': 'mean',
                'task_id': 'count'
            }).round(2)

            for dept, row in dept_summary.iterrows():
                content += f"  {dept}: {row['task_id']}个任务 (平均紧急度:{row['urgency_score']:.1f}, 平均适合度:{row['person_suitability_score']:.1f})\n"

        if not self.result['unallocated_tasks'].empty:
            content += f"\n四、未分配任务\n未分配任务数: {len(self.result['unallocated_tasks'])}\n"
            for _, row in self.result['unallocated_tasks'].iterrows():
                content += f"  - {row['task_id']}: {row['reason']}\n"

        content += f"\n五、系统建议\n"
        if self.allocator.system_load['system_load_percentage'] > 90:
            content += "  系统负载过高，建议暂停分配新任务\n"
        elif self.allocator.system_load['system_load_percentage'] > 70:
            content += "  系统负载较高，建议谨慎分配新任务\n"
        else:
            content += "  系统负载正常，可继续分配新任务\n"

        return content

    def export_results(self):
        """导出结果"""
        if not hasattr(self, 'result') or not self.result:
            messagebox.showwarning("提示", "请先执行分配")
            return

        file_path = filedialog.asksaveasfilename(
            title="保存分配结果",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )

        if file_path:
            try:
                with pd.ExcelWriter(file_path) as writer:
                    if not self.result['allocations'].empty:
                        self.result['allocations'].to_excel(writer, sheet_name='分配结果', index=False)

                    if not self.result['unallocated_tasks'].empty:
                        self.result['unallocated_tasks'].to_excel(writer, sheet_name='未分配任务', index=False)

                    # 系统状态
                    system_data = []
                    for dept, load_info in self.allocator.system_load['dept_load'].items():
                        system_data.append({
                            '机关名称': dept,
                            '负载率%': load_info['load_percentage'],
                            '当前任务': load_info['current_tasks'],
                            '总容量': load_info['capacity']
                        })

                    if system_data:
                        pd.DataFrame(system_data).to_excel(writer, sheet_name='系统状态', index=False)

                self.update_status(f"结果已导出到: {file_path}", "success")
                messagebox.showinfo("成功", "导出完成！")

            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {str(e)}")

    def save_settings(self):
        """保存设置"""
        try:
            settings = {
                'load_threshold': int(self.load_threshold_var.get()),
                'weights': self.weights_var.get(),
                'urgent_threshold': int(self.urgent_threshold_var.get())
            }

            self.update_status("设置已保存", "success")
            messagebox.showinfo("成功", "设置保存成功！")

        except ValueError:
            messagebox.showerror("错误", "请输入有效的数值！")

    def clear_all_data(self):
        """清空所有数据"""
        if messagebox.askyesno("确认", "确定要清空所有数据吗？"):
            self.history_df = None
            self.new_tasks_df = None
            self.allocator = None
            self.result = None

            # 清空界面
            for item in self.allocation_tree.get_children():
                self.allocation_tree.delete(item)

            self.detail_text.delete(1.0, tk.END)
            self.info_text.delete(1.0, tk.END)

            # 重置卡片
            for card in self.cards.values():
                card.config(text="0")

            self.update_status("数据已清空", "info")

    def clear_results(self):
        """清空分配结果"""
        for item in self.allocation_tree.get_children():
            self.allocation_tree.delete(item)
        self.detail_text.delete(1.0, tk.END)
        self.update_status("分配结果已清空", "info")

    def reset_system(self):
        """重置系统"""
        if messagebox.askyesno("确认", "确定要重置系统吗？"):
            self.clear_all_data()
            self.load_sample_data_async()
            self.update_status("系统已重置", "success")

    def show_help(self):
        """显示帮助"""
        help_text = """智能任务分配系统 - 使用说明

一、紧急度计算方式：
   紧急度 = 时间紧迫性(40%) + 机关负载(30%) + 历史延期风险(30%)

   1. 时间紧迫性评分：
      - 已逾期：100分
      - 24小时内：90分
      - 3天内：70分
      - 7天内：40分
      - 7天以上：20分

   2. 机关负载评分：
      - >85%负载：100分
      - 70-85%负载：80分
      - 50-70%负载：60分
      - 30-50%负载：40分
      - <30%负载：20分

   3. 历史延期风险评分：
      - 根据机关历史处理时间和剩余时间比例评估
      - 处理时间/剩余时间 > 2：100分（高风险）
      - 1.5-2：80分
      - 1-1.5：60分
      - 0.5-1：40分
      - <0.5：20分

二、复杂度计算方式（人员适合度）：
   适合度 = 当前负载(40%) + 处理效率(30%) + 任务匹配(30%)

   1. 当前负载评分：
      - 当前任务数/最大能力 * 100
      - 负载越低分数越高

   2. 处理效率评分：
      - 基于历史平均处理时间
      - 处理时间越短效率越高

   3. 任务匹配评分：
      - 高紧急任务：侧重效率(70%) + 负载(30%)
      - 中高紧急任务：效率(50%) + 负载(50%)
      - 低紧急任务：侧重负载均衡(70%) + 效率(30%)
"""

        messagebox.showinfo("使用帮助", help_text)

    def update_status(self, message, status_type="info"):
        """更新状态栏"""
        colors = {
            "info": "black",
            "success": self.success_color,
            "warning": self.warning_color,
            "error": self.danger_color
        }

        self.status_label.config(text=message, foreground=colors.get(status_type, "black"))


def main():
    try:
        root = tk.Tk()
        app = TaskAllocatorUI(root)
        root.mainloop()
    except Exception as e:
        print(f"程序启动失败: {e}")
        input("按回车键退出...")  # 在命令行中显示错误信息


if __name__ == "__main__":
    main()