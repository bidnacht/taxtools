import tkinter as tk
from tkinter import ttk, messagebox

class NewToolUI:
    def __init__(self, root):
        self.root = root
        self.root.title("新小工具")
        self.root.geometry("600x400")
        
        # 设置样式
        self.setup_styles()
        
        # 创建UI
        self.create_widgets()
    
    def setup_styles(self):
        """设置样式"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # 自定义颜色
        self.bg_color = "#f5f7fa"
        self.accent_color = "#3498db"
        
        # 配置样式
        style.configure("TFrame", background=self.bg_color)
        style.configure("TLabel", background=self.bg_color, font=("微软雅黑", 10))
        style.configure("Title.TLabel", font=("微软雅黑", 16, "bold"))
        style.configure("TButton", padding=6)
        
        self.root.configure(bg=self.bg_color)
    
    def create_widgets(self):
        """创建所有UI部件"""
        # 主容器
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 标题
        title_label = ttk.Label(main_container, text="新小工具", style="Title.TLabel")
        title_label.pack(pady=20)
        
        # 功能按钮
        btn_frame = ttk.Frame(main_container)
        btn_frame.pack(pady=20)
        
        ttk.Button(btn_frame, text="功能1", command=self.function1).pack(pady=5, fill=tk.X)
        ttk.Button(btn_frame, text="功能2", command=self.function2).pack(pady=5, fill=tk.X)
        ttk.Button(btn_frame, text="关于", command=self.show_about).pack(pady=5, fill=tk.X)
    
    def function1(self):
        """功能1"""
        messagebox.showinfo("功能1", "这是功能1")
    
    def function2(self):
        """功能2"""
        messagebox.showinfo("功能2", "这是功能2")
    
    def show_about(self):
        """显示关于信息"""
        messagebox.showinfo("关于", "新小工具 v1.0\n\n用于测试打包环境")

if __name__ == "__main__":
    root = tk.Tk()
    app = NewToolUI(root)
    root.mainloop()