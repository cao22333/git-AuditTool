import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
from datetime import datetime

from core.reconciliation import ReconciliationManager

class ReconciliationUI:
    """
    科目勾稽功能的用户界面。
    集成了薪酬勾稽和资产折旧摊销勾稽两个子功能。
    """
    def __init__(self, parent_frame, app_instance):
        self.parent_frame = parent_frame
        self.app_instance = app_instance # 主应用程序实例，用于调用update_progress等方法
        
        # 实例化核心逻辑模块
        self.reconciliation_manager = ReconciliationManager(logger=self._log_message) # 注入UI的日志方法
        
        # 状态变量
        self.processing = False
        
        # 初始化UI变量
        self._init_variables()
        
        # 搭建界面
        self._setup_ui()

    def _init_variables(self):
        """初始化所有用户输入变量"""
        # 薪酬勾稽相关
        self.payroll_zt_path_var = tk.StringVar()
        self.payroll_mapping_file_var = tk.StringVar()
        self.payroll_output_dir_var = tk.StringVar()

        # 资产折旧摊销勾稽相关
        self.asset_yb_path_var = tk.StringVar()
        self.asset_template_file_var = tk.StringVar()
        self.asset_output_dir_var = tk.StringVar()
        
    def _setup_ui(self):
        """搭建主界面"""
        # 创建一个 Notebook (标签页) 来组织两个勾稽功能
        self.notebook = ttk.Notebook(self.parent_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建并添加各个功能标签页
        self._create_payroll_reconciliation_tab()
        self._create_asset_reconciliation_tab()

    def _create_payroll_reconciliation_tab(self):
        """创建薪酬勾稽标签页"""
        payroll_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(payroll_frame, text="💰 薪酬勾稽")
        
        # 账套文件路径
        zt_path_frame = ttk.LabelFrame(payroll_frame, text="账套文件目录", padding="10")
        zt_path_frame.pack(fill=tk.X, pady=5)
        ttk.Entry(zt_path_frame, textvariable=self.payroll_zt_path_var, width=80).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        ttk.Button(zt_path_frame, text="浏览", command=lambda: self.browse_directory(self.payroll_zt_path_var)).pack(side=tk.RIGHT)
        
        # 勾稽映射表文件
        mapping_file_frame = ttk.LabelFrame(payroll_frame, text="勾稽映射表文件 (Excel)", padding="10")
        mapping_file_frame.pack(fill=tk.X, pady=5)
        ttk.Entry(mapping_file_frame, textvariable=self.payroll_mapping_file_var, width=80).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        ttk.Button(mapping_file_frame, text="浏览", command=lambda: self.browse_file(self.payroll_mapping_file_var, [("Excel文件", "*.xlsx")])).pack(side=tk.RIGHT)
        
        # 输出目录
        output_dir_frame = ttk.LabelFrame(payroll_frame, text="输出底稿文件目录", padding="10")
        output_dir_frame.pack(fill=tk.X, pady=5)
        ttk.Entry(output_dir_frame, textvariable=self.payroll_output_dir_var, width=80).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        ttk.Button(output_dir_frame, text="浏览", command=lambda: self.browse_directory(self.payroll_output_dir_var)).pack(side=tk.RIGHT)
        
        # 执行按钮
        ttk.Button(payroll_frame, text="开始薪酬勾稽", command=self.process_payroll_reconciliation, style='TButton').pack(pady=10)

    def _create_asset_reconciliation_tab(self):
        """创建资产折旧摊销勾稽标签页"""
        asset_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(asset_frame, text="🏛️ 资产折旧摊销勾稽")

        # 原始报表目录
        yb_path_frame = ttk.LabelFrame(asset_frame, text="原始报表目录", padding="10")
        yb_path_frame.pack(fill=tk.X, pady=5)
        ttk.Entry(yb_path_frame, textvariable=self.asset_yb_path_var, width=80).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        ttk.Button(yb_path_frame, text="浏览", command=lambda: self.browse_directory(self.asset_yb_path_var)).pack(side=tk.RIGHT)

        # 折旧分配表模板
        template_file_frame = ttk.LabelFrame(asset_frame, text="折旧分配表模板 (Excel)", padding="10")
        template_file_frame.pack(fill=tk.X, pady=5)
        ttk.Entry(template_file_frame, textvariable=self.asset_template_file_var, width=80).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        ttk.Button(template_file_frame, text="浏览", command=lambda: self.browse_file(self.asset_template_file_var, [("Excel文件", "*.xlsx")])).pack(side=tk.RIGHT)

        # 输出目录
        output_dir_frame = ttk.LabelFrame(asset_frame, text="输出底稿文件目录", padding="10")
        output_dir_frame.pack(fill=tk.X, pady=5)
        ttk.Entry(output_dir_frame, textvariable=self.asset_output_dir_var, width=80).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        ttk.Button(output_dir_frame, text="浏览", command=lambda: self.browse_directory(self.asset_output_dir_var)).pack(side=tk.RIGHT)
        
        # 执行按钮
        ttk.Button(asset_frame, text="开始资产折旧摊销勾稽", command=self.process_asset_reconciliation, style='TButton').pack(pady=10)

    def _log_message(self, msg, level="INFO"):
        """
        用于接收 ReconciliationManager 模块的日志消息并在主应用的日志/状态栏显示。
        这是注入给 ReconciliationManager 的日志回调函数。
        """
        self.app_instance.root.after(0, lambda: self.app_instance.update_status(f"科目勾稽: {msg}", level.lower()))

    def browse_file(self, file_var, filetypes):
        """文件选择对话框"""
        filename = filedialog.askopenfilename(filetypes=filetypes)
        if filename:
            file_var.set(filename)
            return True
        return False

    def browse_directory(self, dir_var):
        """目录选择对话框"""
        directory = filedialog.askdirectory()
        if directory:
            dir_var.set(directory)
            return True
        return False

    def process_payroll_reconciliation(self):
        """启动薪酬勾稽线程"""
        if not self._validate_payroll_inputs():
            return
        
        if self.processing:
            messagebox.showwarning("警告", "已有任务在处理中，请稍候。")
            return
        
        self.processing = True
        self.app_instance.update_status("正在执行薪酬勾稽...", "info")
        self.app_instance.show_progress("正在执行薪酬勾稽...")
        threading.Thread(target=self._payroll_reconciliation_thread, daemon=True).start()

    def _payroll_reconciliation_thread(self):
        """薪酬勾稽后台线程"""
        try:
            zt_path = self.payroll_zt_path_var.get()
            mapping_file = self.payroll_mapping_file_var.get()
            output_base_dir = self.payroll_output_dir_var.get()

            success, msg = self.reconciliation_manager.payroll_reconciliation(
                zt_path, mapping_file, output_base_dir,
                progress_callback=self.app_instance.update_progress
            )
            
            if success:
                self.app_instance.root.after(0, lambda: messagebox.showinfo("成功", msg))
                self.app_instance.update_status("薪酬勾稽完成", "success")
            else:
                self.app_instance.root.after(0, lambda: messagebox.showerror("错误", msg))
                self.app_instance.update_status("薪酬勾稽失败", "error")
        except Exception as e:
            self.app_instance.root.after(0, lambda: messagebox.showerror("错误", f"薪酬勾稽失败: {str(e)}"))
            self.app_instance.update_status("薪酬勾稽失败", "error")
        finally:
            self.processing = False
            self.app_instance.root.after(0, self.app_instance.hide_progress)

    def _validate_payroll_inputs(self):
        """验证薪酬勾稽输入"""
        if not self.payroll_zt_path_var.get():
            messagebox.showerror("错误", "请选择账套文件目录。")
            return False
        if not self.payroll_mapping_file_var.get():
            messagebox.showerror("错误", "请选择勾稽映射表文件。")
            return False
        if not self.payroll_output_dir_var.get():
            messagebox.showerror("错误", "请选择输出底稿文件目录。")
            return False
        return True

    def process_asset_reconciliation(self):
        """启动资产折旧摊销勾稽线程"""
        if not self._validate_asset_inputs():
            return
        
        if self.processing:
            messagebox.showwarning("警告", "已有任务在处理中，请稍候。")
            return
        
        self.processing = True
        self.app_instance.update_status("正在执行资产折旧摊销勾稽...", "info")
        self.app_instance.show_progress("正在执行资产折旧摊销勾稽...")
        threading.Thread(target=self._asset_reconciliation_thread, daemon=True).start()

    def _asset_reconciliation_thread(self):
        """资产折旧摊销勾稽后台线程"""
        try:
            yb_path = self.asset_yb_path_var.get()
            template_file = self.asset_template_file_var.get()
            output_base_dir = self.asset_output_dir_var.get()

            success, msg = self.reconciliation_manager.asset_depreciation_reconciliation(
                yb_path, template_file, output_base_dir,
                progress_callback=self.app_instance.update_progress
            )
            
            if success:
                self.app_instance.root.after(0, lambda: messagebox.showinfo("成功", msg))
                self.app_instance.update_status("资产折旧摊销勾稽完成", "success")
            else:
                self.app_instance.root.after(0, lambda: messagebox.showerror("错误", msg))
                self.app_instance.update_status("资产折旧摊销勾稽失败", "error")
        except Exception as e:
            self.app_instance.root.after(0, lambda: messagebox.showerror("错误", f"资产折旧摊销勾稽失败: {str(e)}"))
            self.app_instance.update_status("资产折旧摊销勾稽失败", "error")
        finally:
            self.processing = False
            self.app_instance.root.after(0, self.app_instance.hide_progress)

    def _validate_asset_inputs(self):
        """验证资产折旧摊销勾稽输入"""
        if not self.asset_yb_path_var.get():
            messagebox.showerror("错误", "请选择原始报表目录。")
            return False
        if not self.asset_template_file_var.get():
            messagebox.showerror("错误", "请选择折旧分配表模板文件。")
            return False
        if not self.asset_output_dir_var.get():
            messagebox.showerror("错误", "请选择输出底稿文件目录。")
            return False
        return True