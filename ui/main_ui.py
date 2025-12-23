import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk # 用于处理图像
from ui.csv_processor_ui import CsvProcessorUI # 导入 CSV 处理功能UI
from ui.folder_manager_ui import FolderManagerUI # 导入文件夹管理功能UI
from ui.reconciliation_ui import ReconciliationUI # 导入科目勾稽功能UI

class AuditToolApp:
    """
    审计工具箱主应用程序类。
    负责主界面的布局、样式设置、功能导航以及通用UI元素的管理。
    """
    def __init__(self, root):
        self.root = root
        self.root.title("审计工具箱")
        self.root.geometry("1200x800") # 调整窗口大小以适应更多功能
        
        # 设置UI主题和样式
        self._setup_styles()
        
        # 创建主界面框架
        self._create_main_frames()
        
        # 加载并显示图标
        self._load_icon()
        
        # 创建导航菜单
        self._create_navigation()
        
        # 创建内容区域
        self._create_content_area()
        
        # 初始显示某个功能页面
        self._show_welcome_page()

    def _setup_styles(self):
        """
        设置应用程序的 ttk 样式。
        定义了蓝白色调和字体。
        """
        style = ttk.Style()
        
        # 尝试使用不同的主题，确保在不同系统上显示良好
        try:
            style.theme_use('clam') # 'clam', 'alt', 'default', 'vista', 'xpnative'
        except tk.TclError:
            pass # 如果主题不存在，则使用默认主题

        # 定义颜色方案
        self.colors = {
            "primary": "#4A90E2",  # 蓝色
            "secondary": "#FFFFFF", # 白色
            "background": "#F0F8FF", # 浅蓝色背景
            "text": "#333333",     # 深灰色文本
            "accent": "#5DADE2",   # 较亮的蓝色
            "border": "#CCCCCC"    # 边框颜色
        }

        # 配置通用样式
        self.root.configure(bg=self.colors["background"])
        style.configure('.', font=('Microsoft YaHei UI', 10), background=self.colors["background"], foreground=self.colors["text"])
        style.configure('TFrame', background=self.colors["background"])
        style.configure('TLabel', background=self.colors["background"], foreground=self.colors["text"])
        style.configure('TButton', font=('Microsoft YaHei UI', 10, 'bold'), background=self.colors["primary"], foreground=self.colors["secondary"], borderwidth=0, focusthickness=3, focuscolor=self.colors["accent"])
        style.map('TButton', background=[('active', self.colors["accent"])])
        
        style.configure('TNotebook', background=self.colors["background"], borderwidth=0)
        style.configure('TNotebook.Tab', background=self.colors["primary"], foreground=self.colors["secondary"], padding=[10, 5])
        style.map('TNotebook.Tab', background=[('selected', self.colors["accent"])], foreground=[('selected', self.colors["secondary"])])

        style.configure('TLabelframe', background=self.colors["background"], foreground=self.colors["primary"], font=('Microsoft YaHei UI', 11, 'bold'))
        style.configure('TLabelframe.Label', background=self.colors["background"], foreground=self.colors["primary"], font=('Microsoft YaHei UI', 11, 'bold'))
        
        style.configure('Treeview', background=self.colors["secondary"], foreground=self.colors["text"], fieldbackground=self.colors["secondary"], borderwidth=1, relief="solid")
        style.configure('Treeview.Heading', font=('Microsoft YaHei UI', 10, 'bold'), background=self.colors["primary"], foreground=self.colors["secondary"])
        style.map('Treeview.Heading', background=[('active', self.colors["accent"])])
        
        style.configure('TEntry', fieldbackground=self.colors["secondary"], foreground=self.colors["text"], bordercolor=self.colors["border"])
        style.configure('TCombobox', fieldbackground=self.colors["secondary"], foreground=self.colors["text"], selectbackground=self.colors["accent"], selectforeground=self.colors["secondary"])

    def _create_main_frames(self):
        """
        创建主界面的左右分栏框架：侧边栏用于导航，主区域用于显示内容。
        """
        # 侧边栏框架 (左侧)
        self.sidebar_frame = ttk.Frame(self.root, width=200, style='TFrame', relief=tk.RAISED, borderwidth=1)
        self.sidebar_frame.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)
        self.sidebar_frame.pack_propagate(False) # 防止子组件改变框架大小

        # 内容区域框架 (右侧)
        self.content_frame = ttk.Frame(self.root, style='TFrame', relief=tk.FLAT, borderwidth=0)
        self.content_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)

        # 状态栏
        self.status_bar_var = tk.StringVar(value="就绪")
        self.status_bar = ttk.Label(self.root, textvariable=self.status_bar_var, relief=tk.SUNKEN, anchor=tk.W, font=('Microsoft YaHei UI', 9), background=self.colors["background"], foreground=self.colors["text"])
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def _load_icon(self):
        """
        加载并显示应用程序图标。
        图标放置在侧边栏的左上方。
        """
        icon_path = "assets/icon.png" # 假设图标文件在 assets 文件夹中
        try:
            # 尝试加载图标，并将其调整为合适的大小
            original_image = Image.open(icon_path)
            resized_image = original_image.resize((64, 64), Image.Resampling.LANCZOS)
            self.icon_photo = ImageTk.PhotoImage(resized_image)
            
            icon_label = ttk.Label(self.sidebar_frame, image=self.icon_photo, background=self.colors["background"])
            icon_label.pack(pady=10)
        except FileNotFoundError:
            messagebox.showwarning("图标缺失", f"未找到图标文件: {icon_path}")
            # 如果找不到图标，显示一个占位符文本
            placeholder_label = ttk.Label(self.sidebar_frame, text="[图标]", font=('Microsoft YaHei UI', 12, 'bold'), foreground=self.colors["primary"], background=self.colors["background"])
            placeholder_label.pack(pady=10)
        except Exception as e:
            messagebox.showerror("图标加载错误", f"加载图标时发生错误: {e}")
            placeholder_label = ttk.Label(self.sidebar_frame, text="[图标]", font=('Microsoft YaHei UI', 12, 'bold'), foreground=self.colors["primary"], background=self.colors["background"])
            placeholder_label.pack(pady=10)

    def _create_navigation(self):
        """
        创建侧边栏导航按钮，用于切换不同的功能模块。
        """
        # 标题
        ttk.Label(self.sidebar_frame, text="功能菜单", font=('Microsoft YaHei UI', 12, 'bold'), foreground=self.colors["primary"], background=self.colors["background"]).pack(pady=(10, 20))

        # 导航按钮
        self.nav_buttons = {}
        features = {
            "CSV 文件处理": self._show_csv_processor_page,
            "文件夹状态管理": self._show_folder_manager_page,
            "科目勾稽": self._show_reconciliation_page,
            "使用说明": self._show_help_page,
            "关于": self._show_about_page,
            "退出": self.root.quit
        }

        for text, command in features.items():
            btn = ttk.Button(self.sidebar_frame, text=text, command=command, style='TButton')
            btn.pack(fill=tk.X, pady=5, padx=10)
            self.nav_buttons[text] = btn

    def _create_content_area(self):
        """
        创建内容区域，用于动态加载不同功能模块的界面。
        """
        self.current_page = None # 用于跟踪当前显示的功能页面

        # 头部标题区域
        self.header_frame = ttk.Frame(self.content_frame, style='TFrame', relief=tk.FLAT)
        self.header_frame.pack(fill=tk.X, pady=(0, 10))
        self.page_title_label = ttk.Label(self.header_frame, text="欢迎使用审计工具箱", font=('Microsoft YaHei UI', 16, 'bold'), foreground=self.colors["primary"], background=self.colors["background"])
        self.page_title_label.pack(side=tk.LEFT, padx=10)

        # 页面容器，用于动态切换不同功能的界面
        self.page_container = ttk.Frame(self.content_frame, style='TFrame')
        self.page_container.pack(fill=tk.BOTH, expand=True)

        # 进度条区域，初始隐藏
        self.progress_frame = ttk.Frame(self.content_frame, style='TFrame')
        self.progress_label = ttk.Label(self.progress_frame, text="", font=("Microsoft YaHei UI", 9), background=self.colors["background"], foreground=self.colors["text"])
        self.progress_label.pack(anchor=tk.CENTER, pady=(2, 0))
        self.progress_bar = ttk.Progressbar(self.progress_frame, mode='determinate', length=400)
        self.progress_bar.pack(fill=tk.X, pady=(0, 5))
        self.progress_frame.pack_forget() # 默认隐藏

        # 实例化各个功能UI，但不立即显示
        self.csv_processor_ui = CsvProcessorUI(self.page_container, self)
        self.folder_manager_ui = FolderManagerUI(self.page_container, self) # 实例化 FolderManagerUI
        self.reconciliation_ui = ReconciliationUI(self.page_container, self) # 实例化 ReconciliationUI

    def _clear_content_area(self):
        """
        清除内容区域的所有子组件，准备加载新的功能页面。
        """
        for widget in self.page_container.winfo_children():
            widget.destroy()

    def _set_page_title(self, title):
        """
        设置内容区域的页面标题。
        """
        self.page_title_label.config(text=title)

    def update_status(self, message, level="info"):
        """
        更新状态栏信息。
        level: 'info', 'success', 'warning', 'error'
        """
        colors_map = {
            "info": self.colors["text"],
            "success": "#28A745", # 绿色
            "warning": "#FFC107", # 黄色
            "error": "#DC3545"    # 红色
        }
        self.status_bar_var.set(message)
        self.status_bar.config(foreground=colors_map.get(level, self.colors["text"]))

    def show_progress(self, text="处理中..."):
        """显示进度条"""
        self.progress_frame.pack(fill=tk.X, pady=(0, 5))
        self.progress_label.config(text=text)
        self.progress_bar['value'] = 0
        self.root.update_idletasks() # 强制更新UI

    def hide_progress(self):
        """隐藏进度条"""
        self.progress_frame.pack_forget()

    def update_progress(self, value, text=None):
        """更新进度条（线程安全）"""
        def _update():
            self.progress_bar['value'] = value
            if text:
                self.progress_label.config(text=text)
            self.root.update_idletasks()
        self.root.after(0, _update)

    def _show_welcome_page(self):
        """
        显示欢迎页面。
        """
        self._clear_content_area()
        self._set_page_title("欢迎使用审计工具箱")
        
        welcome_label = ttk.Label(self.page_container, text="请从左侧菜单选择一个功能开始使用。", font=('Microsoft YaHei UI', 14), background=self.colors["background"], foreground=self.colors["text"])
        welcome_label.pack(expand=True)
        
        # 添加一个简要的工具箱介绍
        intro_text = """
        本工具箱旨在提高财务审计效率，集成以下核心功能：
        1. 超大 CSV 文件处理：高效合并、汇总和筛选 CSV 数据。
        2. 文件夹状态管理：根据文件内容自动更新文件夹名称。
        3. 科目勾稽：自动化财务科目数据核对与分析。

        我们致力于提供一个直观、高效、可扩展的审计辅助工具。
        """
        intro_label = ttk.Label(self.page_container, text=intro_text, font=('Microsoft YaHei UI', 10), background=self.colors["background"], foreground=self.colors["text"], wraplength=700, justify=tk.LEFT)
        intro_label.pack(pady=20, padx=20)

    # 以下是各个功能页面的占位符方法，后续会填充具体实现
    def _show_csv_processor_page(self):
        """
        显示 CSV 文件处理功能页面。
        """
        self._clear_content_area()
        self._set_page_title("CSV 文件处理")
        self.csv_processor_ui._setup_ui() # 调用其setup_ui方法来显示界面

    def _show_folder_manager_page(self):
        """
        显示文件夹状态管理功能页面。
        """
        self._clear_content_area()
        self._set_page_title("文件夹状态管理")
        self.folder_manager_ui._setup_ui() # 调用其setup_ui方法来显示界面

    def _show_reconciliation_page(self):
        """
        显示科目勾稽功能页面。
        """
        self._clear_content_area()
        self._set_page_title("科目勾稽")
        self.reconciliation_ui._setup_ui() # 调用其setup_ui方法来显示界面

    def _show_help_page(self):
        """
        显示使用说明页面。
        """
        self._clear_content_area()
        self._set_page_title("使用说明")
        ttk.Label(self.page_container, text="这里将显示工具箱的使用说明。", font=('Microsoft YaHei UI', 12), background=self.colors["background"], foreground=self.colors["text"]).pack(pady=50)

    def _show_about_page(self):
        """
        显示关于页面。
        """
        self._clear_content_area()
        self._set_page_title("关于审计工具箱")
        ttk.Label(self.page_container, text="审计工具箱 v1.0\n开发者：Allen\n\n致力于提升审计工作效率。", font=('Microsoft YaHei UI', 12), background=self.colors["background"], foreground=self.colors["text"]).pack(pady=50)
