import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt
from tkinter import *
from tkinter import ttk, messagebox
from tkinter.filedialog import askopenfilename
import os

def load_data(file_path):
    # 读取CSV格式的txt数据
    df = pd.read_csv(file_path, sep='|', skiprows=[0,1], header=0)
    # 清理空列
    df = df.dropna(axis=1, how='all')
    # 清理字符串首尾空格
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].str.strip()
    return df

def extract_association_points(df, dimensions):
    # 提取关联点，关联点是指在某个维度上关联了至少两个企业的值
    association_points = {}
    for dim in dimensions:
        grouped = df.groupby(dim)['企业名称'].unique().reset_index()
        for _, row in grouped.iterrows():
            point = row[dim]
            enterprises = row['企业名称']
            if len(enterprises) >= 2:
                association_points[point] = {
                    '关联维度': dim,
                    '关联企业': list(enterprises)
                }
    # 转换为DataFrame
    ap_df = pd.DataFrame.from_dict(association_points, orient='index').reset_index()
    ap_df = ap_df.rename(columns={'index': '关联点'})
    return ap_df

def calculate_association_strength(df, dimensions, weights=None):
    # 默认权重为1
    if weights is None:
        weights = {dim: 1 for dim in dimensions}
    
    # 获取所有唯一企业
    enterprises = df['企业名称'].unique()
    # 初始化关联强度矩阵
    strength_matrix = pd.DataFrame(0, index=enterprises, columns=enterprises)
    
    # 计算每个维度的关联强度
    for dim in dimensions:
        weight = weights[dim]
        grouped = df.groupby(dim)['企业名称'].unique().reset_index()
        for _, row in grouped.iterrows():
            ents = row['企业名称']
            if len(ents) >= 2:
                # 为每对企业增加关联强度
                for i in range(len(ents)):
                    for j in range(i+1, len(ents)):
                        ent1 = ents[i]
                        ent2 = ents[j]
                        strength_matrix.loc[ent1, ent2] += weight
                        strength_matrix.loc[ent2, ent1] += weight
    
    # 生成关联方表格
    association_partners = []
    for ent in enterprises:
        # 获取该企业的所有关联方
        partners = strength_matrix.loc[ent][strength_matrix.loc[ent] > 0].sort_values(ascending=False)
        for partner, strength in partners.items():
            if partner != ent:
                association_partners.append({
                    '关联主体': ent,
                    '关联方': partner,
                    '关联强度': strength
                })
    # 去重，避免重复记录
    ap_df = pd.DataFrame(association_partners).drop_duplicates(subset=['关联主体', '关联方'])
    return ap_df, strength_matrix

def visualize_association(strength_matrix, show_points=False, association_points=None):
    # 创建无向图
    G = nx.Graph()
    # 添加企业节点
    enterprises = strength_matrix.index
    G.add_nodes_from(enterprises)
    
    # 添加关联边，边的权重为关联强度
    edges = []
    for i in range(len(enterprises)):
        for j in range(i+1, len(enterprises)):
            ent1 = enterprises[i]
            ent2 = enterprises[j]
            strength = strength_matrix.loc[ent1, ent2]
            if strength > 0:
                edges.append((ent1, ent2, {'weight': strength}))
    G.add_edges_from(edges)
    
    # 设置图的布局
    pos = nx.spring_layout(G, k=0.5, iterations=20)
    
    # 绘制节点
    nx.draw_networkx_nodes(G, pos, node_size=1200, node_color='lightblue')
    # 绘制边，边的粗细根据关联强度调整
    edge_weights = [G[u][v]['weight'] for u, v in G.edges()]
    nx.draw_networkx_edges(G, pos, width=[w*0.6 for w in edge_weights], edge_color='gray')
    # 绘制企业名称标签
    nx.draw_networkx_labels(G, pos, font_size=10, font_weight='bold')
    
    # 如果需要显示关联点，添加边标签
    if show_points and association_points is not None:
        # 构建边与关联点的映射
        edge_point_map = {}
        for _, row in association_points.iterrows():
            point = row['关联点']
            ents = row['关联企业']
            for i in range(len(ents)):
                for j in range(i+1, len(ents)):
                    ent_pair = (ents[i], ents[j])
                    if ent_pair not in edge_point_map:
                        edge_point_map[ent_pair] = []
                    edge_point_map[ent_pair].append(point)
                    # 反向也添加
                    edge_point_map[(ents[j], ents[i])] = edge_point_map[ent_pair]
        # 生成边标签
        edge_labels = {}
        for (u, v), points in edge_point_map.items():
            if (u, v) in G.edges():
                edge_labels[(u, v)] = '\n'.join(points)
        nx.draw_networkx_edge_labels(G, pos, edge_labels=edge_labels, font_size=8, label_pos=0.3)
    
    # 设置图表标题
    plt.title('企业关联关系可视化图', fontsize=14, fontweight='bold')
    plt.axis('off')
    plt.tight_layout()
    plt.show()

class AssociationAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title('企业关联关系分析工具')
        self.root.geometry('1200x700')
        
        # 初始化数据变量
        self.df = None
        self.association_points_df = None
        self.association_partners_df = None
        self.strength_matrix = None
        
        # 1. 数据加载区域
        self.load_frame = ttk.Frame(root, padding=10)
        self.load_frame.pack(fill=X, anchor=N)
        
        ttk.Label(self.load_frame, text='数据文件路径:').pack(side=LEFT, padx=5)
        self.file_path_var = StringVar()
        self.file_entry = ttk.Entry(self.load_frame, textvariable=self.file_path_var, width=60)
        self.file_entry.pack(side=LEFT, padx=5)
        self.browse_btn = ttk.Button(self.load_frame, text='浏览', command=self.browse_file)
        self.browse_btn.pack(side=LEFT, padx=5)
        self.load_btn = ttk.Button(self.load_frame, text='加载数据', command=self.load_data)
        self.load_btn.pack(side=LEFT, padx=5)
        
        # 2. 分析设置区域
        self.settings_frame = ttk.Frame(root, padding=10)
        self.settings_frame.pack(fill=X, anchor=N)
        
        ttk.Label(self.settings_frame, text='选择关联维度:').pack(side=LEFT, padx=5)
        self.dimensions = ['企业经营地址', '人员名称', '联系电话']
        self.dim_vars = {dim: BooleanVar(value=True) for dim in self.dimensions}
        for dim in self.dimensions:
            ttk.Checkbutton(self.settings_frame, text=dim, variable=self.dim_vars[dim]).pack(side=LEFT, padx=8)
        
        ttk.Label(self.settings_frame, text='设置维度权重:').pack(side=LEFT, padx=20)
        self.weight_entries = {}
        for dim in self.dimensions:
            ttk.Label(self.settings_frame, text=f'{dim}:').pack(side=LEFT, padx=2)
            entry = ttk.Entry(self.settings_frame, width=6)
            entry.insert(0, '1')
            entry.pack(side=LEFT, padx=2)
            self.weight_entries[dim] = entry
        
        # 3. 操作按钮区域
        self.ops_frame = ttk.Frame(root, padding=10)
        self.ops_frame.pack(fill=X, anchor=N)
        
        self.extract_btn = ttk.Button(self.ops_frame, text='提取关联点', command=self.extract_association_points)
        self.extract_btn.pack(side=LEFT, padx=8)
        self.calculate_btn = ttk.Button(self.ops_frame, text='计算关联强度', command=self.calculate_association_strength)
        self.calculate_btn.pack(side=LEFT, padx=8)
        self.visualize_btn = ttk.Button(self.ops_frame, text='生成关联图', command=self.visualize_association)
        self.visualize_btn.pack(side=LEFT, padx=8)
        self.save_btn = ttk.Button(self.ops_frame, text='保存分析结果', command=self.save_results)
        self.save_btn.pack(side=LEFT, padx=8)
        
        # 4. 结果显示区域
        self.result_frame = ttk.Frame(root, padding=10)
        self.result_frame.pack(fill=BOTH, expand=True)
        
        # 关联点分析结果表格
        self.ap_label = ttk.Label(self.result_frame, text='关联点分析结果', font=('Arial', 12, 'bold'))
        self.ap_label.pack(anchor=W, padx=5)
        self.ap_tree = ttk.Treeview(self.result_frame, columns=('关联点', '关联维度', '关联企业'), show='headings')
        self.ap_tree.heading('关联点', text='关联点')
        self.ap_tree.heading('关联维度', text='关联维度')
        self.ap_tree.heading('关联企业', text='关联企业')
        self.ap_tree.column('关联点', width=250)
        self.ap_tree.column('关联维度', width=150)
        self.ap_tree.column('关联企业', width=350)
        self.ap_tree.pack(side=LEFT, fill=BOTH, expand=True, padx=5, pady=5)
        
        # 关联方分析结果表格
        self.partner_label = ttk.Label(self.result_frame, text='关联方与关联强度分析结果', font=('Arial', 12, 'bold'))
        self.partner_label.pack(anchor=E, padx=5)
        self.partner_tree = ttk.Treeview(self.result_frame, columns=('关联主体', '关联方', '关联强度'), show='headings')
        self.partner_tree.heading('关联主体', text='关联主体')
        self.partner_tree.heading('关联方', text='关联方')
        self.partner_tree.heading('关联强度', text='关联强度')
        self.partner_tree.column('关联主体', width=200)
        self.partner_tree.column('关联方', width=200)
        self.partner_tree.column('关联强度', width=100)
        self.partner_tree.pack(side=RIGHT, fill=BOTH, expand=True, padx=5, pady=5)
    
    def browse_file(self):
        # 打开文件选择对话框
        file_path = askopenfilename(filetypes=[('文本文件', '*.txt'), ('CSV文件', '*.csv')])
        if file_path:
            self.file_path_var.set(file_path)
    
    def load_data(self):
        # 加载数据文件
        file_path = self.file_path_var.get()
        if not file_path or not os.path.exists(file_path):
            messagebox.showerror('错误', '请选择有效的数据文件！')
            return
        try:
            self.df = load_data(file_path)
            messagebox.showinfo('成功', '数据加载完成！')
        except Exception as e:
            messagebox.showerror('加载失败', f'数据加载出错: {str(e)}')
    
    def extract_association_points(self):
        # 提取关联点
        if self.df is None:
            messagebox.showerror('错误', '请先加载数据！')
            return
        # 获取用户选择的关联维度
        selected_dims = [dim for dim in self.dimensions if self.dim_vars[dim].get()]
        if not selected_dims:
            messagebox.showerror('错误', '请至少选择一个关联维度！')
            return
        # 提取关联点
        self.association_points_df = extract_association_points(self.df, selected_dims)
        # 清空表格并插入数据
        for item in self.ap_tree.get_children():
            self.ap_tree.delete(item)
        for _, row in self.association_points_df.iterrows():
            self.ap_tree.insert('', END, values=(
                row['关联点'],
                row['关联维度'],
                ', '.join(row['关联企业'])
            ))
    
    def calculate_association_strength(self):
        # 计算关联强度
        if self.df is None:
            messagebox.showerror('错误', '请先加载数据！')
            return
        selected_dims = [dim for dim in self.dimensions if self.dim_vars[dim].get()]
        if not selected_dims:
            messagebox.showerror('错误', '请至少选择一个关联维度！')
            return
        # 获取权重设置
        weights = {}
        for dim in selected_dims:
            try:
                weight = float(self.weight_entries[dim].get())
                weights[dim] = weight
            except ValueError:
                messagebox.showerror('错误', f'{dim}的权重必须是数字！')
                return
        # 计算关联强度
        self.association_partners_df, self.strength_matrix = calculate_association_strength(self.df, selected_dims, weights)
        # 清空表格并插入数据
        for item in self.partner_tree.get_children():
            self.partner_tree.delete(item)
        for _, row in self.association_partners_df.iterrows():
            self.partner_tree.insert('', END, values=(
                row['关联主体'],
                row['关联方'],
                row['关联强度']
            ))
    
    def visualize_association(self):
        # 生成关联关系可视化图
        if self.strength_matrix is None:
            messagebox.showerror('错误', '请先计算关联强度！')
            return
        # 询问是否显示关联点
        show_points = messagebox.askyesno('显示关联点', '是否在关联图中显示具体关联点？')
        visualize_association(self.strength_matrix, show_points, self.association_points_df)
    
    def save_results(self):
        # 保存分析结果
        if self.association_points_df is None or self.association_partners_df is None:
            messagebox.showerror('错误', '请先完成关联点提取和关联强度计算！')
            return
        try:
            # 保存关联点结果
            self.association_points_df.to_csv('关联点分析结果.csv', index=False, encoding='utf-8-sig')
            # 保存关联方结果
            self.association_partners_df.to_csv('关联方分析结果.csv', index=False, encoding='utf-8-sig')
            messagebox.showinfo('成功', '分析结果已保存为CSV文件！')
        except Exception as e:
            messagebox.showerror('保存失败', f'保存出错: {str(e)}')

if __name__ == '__main__':
    # 启动GUI程序
    root = Tk()
    app = AssociationAnalyzerApp(root)
    root.mainloop()