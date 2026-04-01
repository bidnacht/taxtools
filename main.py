import pandas as pd
import os
from itertools import combinations
import networkx as nx
from community import community_louvain
from pyvis.network import Network
import argparse
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QPushButton, QLabel, QComboBox, QCheckBox, QLineEdit, QVBoxLayout, QHBoxLayout, QWidget, QScrollArea, QMessageBox
import sys
import difflib

class AssociationAnalyzer:
    def __init__(self, file_path=None, main_entity=None, dimensions=None, weights=None, similarity_dimensions=None):
        self.file_path = file_path
        self.main_entity = main_entity
        self.dimensions = dimensions or []
        self.weights = weights or []
        self.similarity_dimensions = similarity_dimensions or []
        self.df = None
        self.graph = nx.Graph()
        self.communities = {}
    
    def calculate_similarity(self, str1, str2):
        """计算两个字符串的相似度"""
        if pd.isna(str1) or pd.isna(str2):
            return 0.0
        return difflib.SequenceMatcher(None, str(str1), str(str2)).ratio()
    
    def load_data(self):
        try:
            if self.file_path.endswith('.csv') or self.file_path.endswith('.txt'):
                self.df = pd.read_csv(self.file_path)
            else:
                self.df = pd.read_excel(self.file_path)
            print(f"成功加载文件: {os.path.basename(self.file_path)}")
            print(f"数据列: {list(self.df.columns)}")
            print(f"数据行数: {len(self.df)}")
        except Exception as e:
            print(f"读取文件失败: {e}")
            return False
        return True
    
    def build_graph(self):
        """构建关联关系图"""
        edge_list = []
        bridge_list = []
        
        for i, dim_col in enumerate(self.dimensions):
            weight = self.weights[i] if i < len(self.weights) else 1.0
            
            if dim_col in self.similarity_dimensions:
                # 基于相似度的关联
                entities_with_values = self.df[self.df[dim_col].notna()][[self.main_entity, dim_col]].values.tolist()
                
                # 计算所有实体对之间的相似度
                for j, (ent1, val1) in enumerate(entities_with_values):
                    for k, (ent2, val2) in enumerate(entities_with_values[j+1:]):
                        if ent1 != ent2:
                            similarity = self.calculate_similarity(val1, val2)
                            if similarity >= 0.6:  # 设置相似度阈值
                                # 计算调整后的权重
                                adjusted_weight = weight * similarity
                                
                                # 添加关联点
                                bridge_list.append({
                                    "关联点内容": f"{val1} (相似度: {similarity:.2f})",
                                    "关联点维度": dim_col,
                                    "涉及主体": ent1
                                })
                                bridge_list.append({
                                    "关联点内容": f"{val2} (相似度: {similarity:.2f})",
                                    "关联点维度": dim_col,
                                    "涉及主体": ent2
                                })
                                
                                # 添加边
                                edge_list.append({
                                    "Source": ent1,
                                    "Target": ent2,
                                    "Bridge": f"{val1} ↔ {val2} (相似度: {similarity:.2f})",
                                    "Type": dim_col,
                                    "Weight": adjusted_weight
                                })
            else:
                # 精确匹配的关联
                grouped = self.df[self.df[dim_col].notna()].groupby(dim_col)[self.main_entity].unique()
                
                for bridge_value, entities in grouped.items():
                    if len(entities) > 1:
                        # 提取关联点
                        for ent in entities:
                            bridge_list.append({
                                "关联点内容": bridge_value,
                                "关联点维度": dim_col,
                                "涉及主体": ent
                            })
                        
                        # 计算关联对
                        for u, v in combinations(entities, 2):
                            edge_list.append({
                                "Source": u,
                                "Target": v,
                                "Bridge": bridge_value,
                                "Type": dim_col,
                                "Weight": weight
                            })
        
        # 构建图
        for edge in edge_list:
            u = edge["Source"]
            v = edge["Target"]
            weight = edge["Weight"]
            bridge = edge["Bridge"]
            edge_type = edge["Type"]
            
            # 添加节点
            self.graph.add_node(u, label=u, type="entity")
            self.graph.add_node(v, label=v, type="entity")
            
            # 添加边
            if self.graph.has_edge(u, v):
                # 如果边已存在，累加权重
                existing_weight = self.graph[u][v].get("weight", 0)
                existing_bridges = self.graph[u][v].get("bridges", [])
                existing_types = self.graph[u][v].get("types", [])
                
                self.graph[u][v]["weight"] = existing_weight + weight
                self.graph[u][v]["bridges"].append(bridge)
                self.graph[u][v]["types"].append(edge_type)
            else:
                # 新边
                self.graph.add_edge(u, v, weight=weight, bridges=[bridge], types=[edge_type])
        
        return bridge_list, edge_list
    
    def detect_communities(self):
        """使用Louvain算法检测社区"""
        # 提取边权重
        weight_dict = {(u, v): d['weight'] for u, v, d in self.graph.edges(data=True)}
        
        # 应用Louvain算法
        partition = community_louvain.best_partition(self.graph, weight='weight')
        self.communities = partition
        
        # 为节点添加社区属性
        for node, community_id in partition.items():
            self.graph.nodes[node]['community'] = community_id
    
    def calculate_centrality(self):
        """计算节点中心性"""
        # 度中心性
        degree_centrality = nx.degree_centrality(self.graph)
        # 中介中心性
        betweenness_centrality = nx.betweenness_centrality(self.graph, weight='weight')
        
        for node in self.graph.nodes():
            self.graph.nodes[node]['degree'] = degree_centrality[node]
            self.graph.nodes[node]['betweenness'] = betweenness_centrality[node]
    
    def generate_visualization(self, output_dir):
        """生成可视化图谱"""
        net = Network(height="800px", width="100%", bgcolor="#ffffff", font_color="black", notebook=False)
        
        # 添加节点
        for node, attrs in self.graph.nodes(data=True):
            community_id = attrs.get('community', 0)
            degree = attrs.get('degree', 0)
            size = 15 + degree * 50  # 根据中心性调整节点大小
            
            net.add_node(
                node, 
                label=node, 
                size=size,
                group=community_id,
                title=f"企业: {node}\n社区: {community_id}\n关联度: {degree:.2f}"
            )
        
        # 添加边
        for u, v, attrs in self.graph.edges(data=True):
            weight = attrs.get('weight', 1)
            bridges = attrs.get('bridges', [])
            types = attrs.get('types', [])
            
            # 根据权重调整边的粗细
            width = 1 + weight * 0.5
            
            # 边的标题，显示关联点
            bridge_str = "\n".join([f"{t}: {b}" for t, b in zip(types, bridges)])
            title = f"关联强度: {weight}\n关联点:\n{bridge_str}"
            
            net.add_edge(u, v, width=width, title=title, color="#888888")
        
        # 配置
        net.set_options("""
        var options = {
          "nodes": {
            "shape": "circle",
            "shadow": true,
            "scaling": {
              "min": 15,
              "max": 50
            }
          },
          "edges": {
            "smooth": {
              "type": "continuous"
            }
          },
          "physics": {
            "forceAtlas2Based": {
              "gravitationalConstant": -100,
              "centralGravity": 0.01,
              "springLength": 100,
              "springConstant": 0.08
            },
            "maxVelocity": 50,
            "solver": "forceAtlas2Based",
            "stabilization": {
              "iterations": 100
            }
          }
        }
        """)
        
        # 保存为HTML文件
        net.save_graph(f"{output_dir}/association_network.html")
    
    def generate_reports(self, output_dir, bridge_list):
        """生成分析报告"""
        # 1. 关联企业明细表
        edge_data = []
        for u, v, attrs in self.graph.edges(data=True):
            edge_data.append({
                "企业A": u,
                "企业B": v,
                "关联强度": attrs.get('weight', 1),
                "关联类型": ",".join(set(attrs.get('types', []))),
                "关联点": "|".join(set(map(str, attrs.get('bridges', []))))
            })
        pd.DataFrame(edge_data).to_csv(f"{output_dir}/关联企业明细表.csv", index=False, encoding='utf_8_sig')
        
        # 2. 核心关联点摘要
        bridge_df = pd.DataFrame(bridge_list)
        bridge_summary = bridge_df.groupby(['关联点内容', '关联点维度']).agg({
            '涉及主体': lambda x: ",".join(set(x)),
            '关联点内容': 'count'
        }).rename(columns={'关联点内容': '出现次数'}).sort_values('出现次数', ascending=False)
        bridge_summary.to_csv(f"{output_dir}/核心关联点摘要.csv", encoding='utf_8_sig')
        
        # 3. 企业关联分析报告
        node_data = []
        for node, attrs in self.graph.nodes(data=True):
            node_data.append({
                "企业名称": node,
                "社区ID": attrs.get('community', 0),
                "关联度": attrs.get('degree', 0),
                "中介中心性": attrs.get('betweenness', 0),
                "关联企业数": len(list(self.graph.neighbors(node)))
            })
        pd.DataFrame(node_data).to_csv(f"{output_dir}/企业关联分析报告.csv", index=False, encoding='utf_8_sig')
    
    def run_analysis(self, output_dir):
        """运行数据分析流程（生成表格报告）"""
        if self.df is None:
            print("请先加载数据")
            return False
        
        # 1. 构建关联图
        print("构建关联关系图...")
        bridge_list, edge_list = self.build_graph()
        
        if not edge_list:
            print("未发现任何关联关系")
            return True
        
        # 2. 检测社区
        print("检测关联团簇...")
        self.detect_communities()
        
        # 3. 计算中心性
        print("计算节点中心性...")
        self.calculate_centrality()
        
        # 4. 生成报告
        print("生成分析报告...")
        self.generate_reports(output_dir, bridge_list)
        
        print(f"数据分析完成！表格结果已保存至：{output_dir}")
        return True
    
    def generate_visualization_with_options(self, output_dir):
        """生成可视化图谱（带配置选项）"""
        net = Network(height="800px", width="100%", bgcolor="#ffffff", font_color="black", notebook=False)
        
        # 用于存储已经添加的关联点
        added_bridges = set()
        
        # 添加节点
        for node, attrs in self.graph.nodes(data=True):
            community_id = attrs.get('community', 0)
            degree = attrs.get('degree', 0)
            size = 15 + degree * 50  # 根据中心性调整节点大小
            
            net.add_node(
                node, 
                label=node, 
                size=size,
                group=community_id,
                title=f"企业: {node}\n社区: {community_id}\n关联度: {degree:.2f}",
                shape="circle",
                color="#4285F4",  # 主体节点使用蓝色
                is_bridge=False  # 标识为主体节点
            )
        
        # 添加边和关联点
        for u, v, attrs in self.graph.edges(data=True):
            weight = attrs.get('weight', 1)
            bridges = attrs.get('bridges', [])
            types = attrs.get('types', [])
            
            # 根据权重调整边的粗细
            width = 1 + weight * 0.5
            
            # 直接连接两个主体节点（用于隐藏关联点时）
            net.add_edge(u, v, width=width, title=f"关联强度: {weight}", color="#888888", hidden=True, is_direct_edge=True)
            
            # 为每个关联点创建节点
            if bridges:
                for i, (bridge, bridge_type) in enumerate(zip(bridges, types)):
                    bridge_id = f"bridge_{u}_{v}_{i}"
                    if bridge_id not in added_bridges:
                        # 添加关联点节点
                        net.add_node(
                            bridge_id, 
                            label=f"{bridge_type}: {bridge}",
                            size=10,  # 关联点节点较小
                            group=999,  # 使用单独的组
                            title=f"关联点: {bridge}\n类型: {bridge_type}",
                            shape="square",
                            color="#34A853",  # 关联点使用绿色
                            is_bridge=True  # 标识为关联点
                        )
                        added_bridges.add(bridge_id)
                        
                        # 连接主体节点和关联点
                        net.add_edge(u, bridge_id, width=width/2, color="#888888", dashed=True, is_bridge_edge=True)
                        net.add_edge(bridge_id, v, width=width/2, color="#888888", dashed=True, is_bridge_edge=True)
        
        # 配置
        net.set_options("""
        var options = {
          "nodes": {
            "shape": "circle",
            "shadow": true,
            "scaling": {
              "min": 15,
              "max": 50
            }
          },
          "edges": {
            "smooth": {
              "type": "continuous"
            }
          },
          "physics": {
            "forceAtlas2Based": {
              "gravitationalConstant": -100,
              "centralGravity": 0.01,
              "springLength": 100,
              "springConstant": 0.08
            },
            "maxVelocity": 50,
            "solver": "forceAtlas2Based",
            "stabilization": {
              "iterations": 100
            }
          }
        }
        """)
        
        # 保存为HTML文件
        output_file = f"{output_dir}/association_network.html"
        net.save_graph(output_file)
        
        # 读取生成的HTML文件并添加自定义代码
        with open(output_file, 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        # 先修改drawGraph函数，确保network对象是全局的
        html_content = html_content.replace('network = new vis.Network(container, data, options);', 'window.network = new vis.Network(container, data, options); network = window.network;')
        
        # 添加自定义HTML和JavaScript
        custom_html = """
        <div style="position: absolute; top: 10px; left: 10px; z-index: 1000; background: white; padding: 10px; border-radius: 5px; box-shadow: 0 2px 5px rgba(0,0,0,0.2);">
          <button id="toggleBridges" style="padding: 8px 12px; background: #4285F4; color: white; border: none; border-radius: 4px; cursor: pointer;">隐藏关联点</button>
        </div>
        
        <script>
          // 等待drawGraph函数执行完成
          setTimeout(function() {
            // 使用全局network对象
            var network = window.network;
            
            if (network) {
              console.log('Network object found:', network);
              
              // 初始化状态：显示关联点
              var showBridges = true;
              
              // 更新显示状态的函数
              function updateDisplay() {
                console.log('Updating display, showBridges:', showBridges);
                
                // 遍历所有节点
                var nodes = network.body.data.nodes.get();
                console.log('Nodes:', nodes);
                
                // 遍历所有边
                var edges = network.body.data.edges.get();
                console.log('Edges:', edges);
                
                // 更新节点
                for (var i = 0; i < nodes.length; i++) {
                  var node = nodes[i];
                  if (node.is_bridge) {
                    network.body.data.nodes.update({id: node.id, hidden: !showBridges});
                  }
                }
                
                // 更新边
                for (var j = 0; j < edges.length; j++) {
                  var edge = edges[j];
                  if (edge.is_bridge_edge) {
                    network.body.data.edges.update({id: edge.id, hidden: !showBridges});
                  } else if (edge.is_direct_edge) {
                    network.body.data.edges.update({id: edge.id, hidden: showBridges});
                  }
                }
                
                network.redraw();
                console.log('Display updated');
              }
              
              // 绑定按钮点击事件
              var button = document.getElementById('toggleBridges');
              console.log('Button element:', button);
              
              button.addEventListener('click', function() {
                console.log('Button clicked, current showBridges:', showBridges);
                showBridges = !showBridges;
                console.log('New showBridges:', showBridges);
                
                if (showBridges) {
                  button.textContent = '隐藏关联点';
                  button.style.background = '#EA4335';
                } else {
                  button.textContent = '显示关联点';
                  button.style.background = '#4285F4';
                }
                
                updateDisplay();
              });
              
              // 初始化显示
              updateDisplay();
            } else {
              console.error('Network object not found');
            }
          }, 100);
        </script>
        """
        
        # 在</body>标签前添加自定义HTML
        modified_content = html_content.replace('</body>', custom_html + '</body>')
        
        # 保存修改后的HTML文件
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(modified_content)
        
        return output_file

class AssociationApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("企业关联关系分析系统")
        self.setGeometry(100, 100, 800, 600)
        
        self.df = None
        self.analyzer = None
        self.dim_vars = {}
        self.weight_vars = {}
        self.similarity_vars = {}
        self.output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
        self.load_config()
        
        self.setup_ui()
    
    def setup_ui(self):
        # 创建主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 标题和状态区域
        title_layout = QVBoxLayout()
        title_label = QLabel("企业关联关系分析系统")
        title_label.setStyleSheet("font-size: 16pt; font-weight: bold;")
        title_layout.addWidget(title_label)
        
        # 状态显示区域
        self.status_layout = QHBoxLayout()
        self.status_layout.addWidget(QLabel("状态:"))
        
        # 数据加载状态
        self.data_status_label = QLabel("未加载")
        self.data_status_label.setStyleSheet("color: gray; font-weight: bold;")
        self.status_layout.addWidget(QLabel("数据:"))
        self.status_layout.addWidget(self.data_status_label)
        self.status_layout.addWidget(QLabel(" | "))
        
        # 分析状态
        self.analysis_status_label = QLabel("未分析")
        self.analysis_status_label.setStyleSheet("color: gray; font-weight: bold;")
        self.status_layout.addWidget(QLabel("分析:"))
        self.status_layout.addWidget(self.analysis_status_label)
        self.status_layout.addStretch()
        
        title_layout.addLayout(self.status_layout)
        main_layout.addLayout(title_layout)
        
        # 文件选择
        file_layout = QHBoxLayout()
        file_layout.addWidget(QLabel("数据文件:"))
        self.file_edit = QLineEdit()
        file_layout.addWidget(self.file_edit)
        browse_btn = QPushButton("浏览")
        browse_btn.clicked.connect(self.browse_file)
        file_layout.addWidget(browse_btn)
        # 添加文件格式说明按钮
        format_btn = QPushButton("格式说明")
        format_btn.clicked.connect(self.show_format_info)
        file_layout.addWidget(format_btn)
        main_layout.addLayout(file_layout)
        
        # 输出路径配置
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("输出路径:"))
        self.output_edit = QLineEdit(self.output_dir)
        output_layout.addWidget(self.output_edit)
        output_browse_btn = QPushButton("浏览")
        output_browse_btn.clicked.connect(self.browse_output_dir)
        output_layout.addWidget(output_browse_btn)
        main_layout.addLayout(output_layout)
        
        # 主体列选择
        main_entity_layout = QHBoxLayout()
        main_entity_layout.addWidget(QLabel("关联主体列:"))
        self.main_entity_cb = QComboBox()
        main_entity_layout.addWidget(self.main_entity_cb)
        main_layout.addLayout(main_entity_layout)
        
        # 维度选择
        dim_label = QLabel("关联维度:")
        dim_label.setStyleSheet("font-weight: bold;")
        main_layout.addWidget(dim_label)
        
        self.scroll_area = QScrollArea()
        self.dim_widget = QWidget()
        self.dim_layout = QVBoxLayout(self.dim_widget)
        self.scroll_area.setWidget(self.dim_widget)
        self.scroll_area.setWidgetResizable(True)
        main_layout.addWidget(self.scroll_area)
        
        # 图像生成配置
        viz_layout = QHBoxLayout()
        viz_layout.addWidget(QLabel("图像配置:"))
        viz_layout.addStretch()
        main_layout.addLayout(viz_layout)
        
        # 按钮
        btn_layout = QHBoxLayout()
        self.load_btn = QPushButton("加载数据")
        self.load_btn.clicked.connect(self.load_data)
        self.analyze_btn = QPushButton("开始分析")
        self.analyze_btn.clicked.connect(self.run_analysis)
        self.viz_btn = QPushButton("生成图像")
        self.viz_btn.clicked.connect(self.generate_visualization)
        self.clear_btn = QPushButton("清空")
        self.clear_btn.clicked.connect(self.clear_data)
        exit_btn = QPushButton("退出")
        exit_btn.clicked.connect(self.close)
        
        btn_layout.addWidget(self.load_btn)
        btn_layout.addWidget(self.analyze_btn)
        btn_layout.addWidget(self.viz_btn)
        btn_layout.addWidget(self.clear_btn)
        btn_layout.addWidget(exit_btn)
        main_layout.addLayout(btn_layout)
        
        # 初始化按钮状态
        self.update_button_states()
    
    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择数据文件", "", "数据文件 (*.xlsx *.xls *.csv *.txt)")
        if file_path:
            self.file_edit.setText(file_path)
            # 更新按钮状态
            self.update_button_states()
    
    def load_data(self):
        file_path = self.file_edit.text()
        if not file_path:
            QMessageBox.warning(self, "提示", "请选择数据文件")
            return
        
        try:
            if file_path.endswith('.csv') or file_path.endswith('.txt'):
                self.df = pd.read_csv(file_path)
            else:
                self.df = pd.read_excel(file_path)
            
            # 更新主体列下拉框
            cols = self.df.columns.tolist()
            self.main_entity_cb.clear()
            self.main_entity_cb.addItems(cols)
            
            # 清空维度选择
            for i in reversed(range(self.dim_layout.count())):
                widget = self.dim_layout.itemAt(i).widget()
                if widget:
                    widget.deleteLater()
            
            # 创建维度选择复选框
            self.dim_vars = {}
            self.weight_vars = {}
            self.similarity_vars = {}
            
            for col in cols:
                dim_row = QWidget()
                row_layout = QHBoxLayout(dim_row)
                
                chk = QCheckBox(col)
                row_layout.addWidget(chk)
                self.dim_vars[col] = chk
                
                row_layout.addWidget(QLabel("权重:"))
                weight_edit = QLineEdit("1")
                weight_edit.setFixedWidth(50)
                row_layout.addWidget(weight_edit)
                self.weight_vars[col] = weight_edit
                
                # 添加相似度匹配复选框
                sim_chk = QCheckBox("相似度匹配")
                row_layout.addWidget(sim_chk)
                self.similarity_vars[col] = sim_chk
                
                self.dim_layout.addWidget(dim_row)
            
            # 更新状态
            self.data_status_label.setText("已加载")
            self.data_status_label.setStyleSheet("color: green; font-weight: bold;")
            self.analysis_status_label.setText("未分析")
            self.analysis_status_label.setStyleSheet("color: gray; font-weight: bold;")
            
            # 更新按钮状态
            self.update_button_states()
            
            QMessageBox.information(self, "成功", "数据加载成功！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取文件失败: {e}")
    
    def load_config(self):
        """加载配置"""
        config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
        if os.path.exists(config_file):
            try:
                import json
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    if 'output_dir' in config:
                        self.output_dir = config['output_dir']
            except Exception as e:
                print(f"加载配置失败: {e}")
    
    def save_config(self):
        """保存配置"""
        config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
        try:
            import json
            config = {'output_dir': self.output_dir}
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存配置失败: {e}")
    
    def browse_output_dir(self):
        """浏览输出目录"""
        directory = QFileDialog.getExistingDirectory(self, "选择输出目录")
        if directory:
            self.output_dir = directory
            self.output_edit.setText(directory)
            self.save_config()
    
    def show_format_info(self):
        """显示文件格式说明"""
        format_info = "支持的文件格式：\n\n"
        format_info += "1. Excel文件：.xlsx, .xls\n"
        format_info += "2. CSV文件：.csv\n"
        format_info += "3. 文本文件：.txt\n\n"
        format_info += "文件要求：\n"
        format_info += "- 第一行必须为列名\n"
        format_info += "- 包含至少一个关联主体列（如企业名称）\n"
        format_info += "- 包含至少一个关联维度列（如地址、人员等）\n"
        format_info += "- 数据编码建议使用UTF-8"
        QMessageBox.information(self, "文件格式说明", format_info)
    
    def run_analysis(self):
        if self.df is None:
            QMessageBox.warning(self, "提示", "请先加载数据")
            return
        
        main_entity = self.main_entity_cb.currentText()
        if not main_entity:
            QMessageBox.warning(self, "提示", "请选择关联主体列")
            return
        
        # 收集选中的维度、权重和相似度匹配设置
        selected_dims = []
        weights = []
        similarity_dims = []
        
        for col, chk in self.dim_vars.items():
            if chk.isChecked():
                selected_dims.append(col)
                try:
                    weight = float(self.weight_vars[col].text())
                    weights.append(weight)
                except:
                    weights.append(1.0)
                
                # 检查是否选择了相似度匹配
                if self.similarity_vars[col].isChecked():
                    similarity_dims.append(col)
        
        if not selected_dims:
            QMessageBox.warning(self, "提示", "请至少选择一个关联维度")
            return
        
        # 使用默认输出路径
        output_dir = self.output_edit.text()
        if not output_dir:
            QMessageBox.warning(self, "提示", "请设置输出路径")
            return
        
        # 确保输出目录存在
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                QMessageBox.critical(self, "错误", f"创建输出目录失败: {e}")
                return
        
        # 创建分析器
        self.analyzer = AssociationAnalyzer(
            file_path=self.file_edit.text(),
            main_entity=main_entity,
            dimensions=selected_dims,
            weights=weights,
            similarity_dimensions=similarity_dims
        )
        self.analyzer.df = self.df
        
        # 运行分析
        try:
            success = self.analyzer.run_analysis(output_dir)
            if success:
                # 更新状态
                self.analysis_status_label.setText("已分析")
                self.analysis_status_label.setStyleSheet("color: blue; font-weight: bold;")
                
                # 更新按钮状态
                self.update_button_states()
                
                QMessageBox.information(self, "成功", f"分析完成！结果已保存至：\n{output_dir}")
            else:
                QMessageBox.information(self, "提示", "未发现任何关联关系")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"分析过程中出错: {e}")
    
    def generate_visualization(self):
        """生成可视化图像"""
        if self.analyzer is None or self.analyzer.graph is None:
            QMessageBox.warning(self, "提示", "请先运行分析")
            return
        
        # 使用默认输出路径
        output_dir = self.output_edit.text()
        if not output_dir:
            QMessageBox.warning(self, "提示", "请设置输出路径")
            return
        
        # 确保输出目录存在
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                QMessageBox.critical(self, "错误", f"创建输出目录失败: {e}")
                return
        
        try:
            # 生成图像
            output_file = self.analyzer.generate_visualization_with_options(output_dir)
            
            # 自动打开HTML文件
            import webbrowser
            webbrowser.open(f"file://{output_file}")
            
            QMessageBox.information(self, "成功", f"图像生成完成并已打开！\n保存位置：{output_file}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"生成图像时出错: {e}")
    
    def update_button_states(self):
        """更新按钮状态"""
        # 重置所有按钮样式
        default_style = ""
        highlight_style = "background-color: #4285F4; color: white; font-weight: bold;"
        disabled_style = "color: #CCCCCC;"
        
        # 检查是否有文件路径
        has_file_path = bool(self.file_edit.text().strip())
        
        # 根据状态更新按钮
        if self.df is None:
            # 未加载数据
            if has_file_path:
                # 有文件路径，加载数据按钮可用
                self.load_btn.setEnabled(True)
                self.load_btn.setStyleSheet(highlight_style)
            else:
                # 无文件路径，加载数据按钮不可用
                self.load_btn.setEnabled(False)
                self.load_btn.setStyleSheet(disabled_style)
            # 分析和图像按钮始终禁用
            self.analyze_btn.setEnabled(False)
            self.analyze_btn.setStyleSheet(disabled_style)
            self.viz_btn.setEnabled(False)
            self.viz_btn.setStyleSheet(disabled_style)
        elif self.analyzer is None or not hasattr(self.analyzer, 'graph') or self.analyzer.graph is None:
            # 已加载数据，未分析，加载数据和开始分析按钮可用
            self.load_btn.setEnabled(True)
            self.load_btn.setStyleSheet(default_style)
            self.analyze_btn.setEnabled(True)
            self.analyze_btn.setStyleSheet(highlight_style)
            self.viz_btn.setEnabled(False)
            self.viz_btn.setStyleSheet(disabled_style)
        else:
            # 已分析，所有按钮都可用
            self.load_btn.setEnabled(True)
            self.load_btn.setStyleSheet(default_style)
            self.analyze_btn.setEnabled(True)
            self.analyze_btn.setStyleSheet(default_style)
            self.viz_btn.setEnabled(True)
            self.viz_btn.setStyleSheet(highlight_style)
    
    def clear_data(self):
        """清空数据"""
        # 显示确认对话框
        reply = QMessageBox.question(self, "确认清空", "是否同时删除输出文件？",
                                    QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel)
        
        if reply == QMessageBox.Cancel:
            return
        
        # 删除输出文件
        if reply == QMessageBox.Yes:
            output_dir = self.output_edit.text()
            if os.path.exists(output_dir):
                try:
                    import shutil
                    shutil.rmtree(output_dir)
                    QMessageBox.information(self, "成功", "输出文件已删除")
                except Exception as e:
                    QMessageBox.critical(self, "错误", f"删除输出文件失败: {e}")
        
        # 清空数据
        self.df = None
        self.analyzer = None
        
        # 清空维度选择
        for i in reversed(range(self.dim_layout.count())):
            widget = self.dim_layout.itemAt(i).widget()
            if widget:
                widget.deleteLater()
        
        # 清空主体列下拉框
        self.main_entity_cb.clear()
        
        # 重置状态
        self.data_status_label.setText("未加载")
        self.data_status_label.setStyleSheet("color: gray; font-weight: bold;")
        self.analysis_status_label.setText("未分析")
        self.analysis_status_label.setStyleSheet("color: gray; font-weight: bold;")
        
        # 更新按钮状态
        self.update_button_states()
        
        QMessageBox.information(self, "成功", "数据已清空")

def main():
    import sys
    if len(sys.argv) > 1:
        # 命令行模式
        parser = argparse.ArgumentParser(description="企业关联关系分析工具")
        parser.add_argument('--file', required=True, help='数据文件路径（支持Excel/CSV/TXT）')
        parser.add_argument('--main-entity', required=True, help='关联主体列名')
        parser.add_argument('--dimensions', nargs='+', required=True, help='关联维度列名')
        parser.add_argument('--weights', nargs='+', type=float, default=[], help='关联维度权重')
        parser.add_argument('--similarity-dimensions', nargs='+', default=[], help='使用相似度匹配的维度列名')
        parser.add_argument('--output', default='output_plan2', help='结果保存目录')
        
        args = parser.parse_args()
        
        analyzer = AssociationAnalyzer(
            file_path=args.file,
            main_entity=args.main_entity,
            dimensions=args.dimensions,
            weights=args.weights,
            similarity_dimensions=args.similarity_dimensions
        )
        
        if analyzer.load_data():
            if not os.path.exists(args.output):
                os.makedirs(args.output)
            analyzer.run_analysis(args.output)
            # 生成可视化文件
            analyzer.generate_visualization_with_options(args.output)
    else:
        # GUI模式
        try:
            app = QApplication(sys.argv)
            window = AssociationApp()
            window.show()
            sys.exit(app.exec_())
        except Exception as e:
            print(f"GUI运行失败: {e}")
            print("请使用命令行模式运行")
            print("示例: python plan2_association_analysis.py --file ceshi.txt --main-entity 企业名称 --dimensions 企业经营地址 人员名称 联系电话")

if __name__ == "__main__":
    main()