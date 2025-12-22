import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
from datetime import datetime
from tkinterdnd2 import DND_FILES, TkinterDnD # 导入拖放功能

from core.folder_manager import FolderManager

class FolderManagerUI:
    """
    文件夹状态管理功能的用户界面。
    集成了文件夹选择、分析、预览、执行重命名和撤销等功能。
    """
    def __init__(self, parent_frame, app_instance):
        self.parent_frame = parent_frame
        self.app_instance = app_instance # 主应用程序实例，用于调用update_progress等方法

        # 实例化核心逻辑模块
        self.folder_manager = FolderManager(logger=self._log_message) # 注入UI的日志方法
        
        # 状态变量
        self.current_dir = ""
        self.processing = False
        self.preview_data = [] # 存储预览数据
        
        # 初始化UI变量
        self._init_variables()
        
        # 尝试启用拖放功能
        try:
            # 检查 parent_frame 是否是 TkinterDnD.Tk() 的实例
            # 如果不是，则假定拖放功能可能不可用或需要父窗口协助
            if isinstance(self.app_instance.root, TkinterDnD.Tk):
                self.dnd_supported = True
            else:
                self.dnd_supported = False
        except NameError: # TkinterDnD 可能未安装或未导入
            self.dnd_supported = False

        # 搭建界面
        self._setup_ui()
        


    def _init_variables(self):
        """初始化所有用户输入变量"""
        self.empty_suffix_var = tk.StringVar(value="-空")
        self.parent_partial_suffix_var = tk.StringVar(value="-缺")
        self.parent_all_empty_suffix_var = tk.StringVar(value="-全空")
        
        self.var_auto_analyze = tk.BooleanVar(value=True) # 自动分析
        self.var_remove_correct_suffix = tk.BooleanVar(value=True) # 移除正确状态后缀
        self.var_create_backup = tk.BooleanVar(value=False) # 创建备份

        self.filter_var = tk.StringVar(value="全部") # 预览筛选变量

    def _setup_ui(self):
        """搭建主界面"""
        # 主框架
        main_frame = ttk.Frame(self.parent_frame, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 左侧控制面板
        control_panel = ttk.LabelFrame(main_frame, text="控制面板", padding="10")
        control_panel.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        
        # 目录选择区域
        dir_section = ttk.LabelFrame(control_panel, text="目录操作", padding="10")
        dir_section.pack(fill=tk.X, pady=(0, 10))
        
        self.dir_label = ttk.Label(dir_section, text="未选择目录", wraplength=200, justify=tk.LEFT, anchor=tk.W)
        self.dir_label.pack(fill=tk.X, pady=(0, 5))
        
        btn_frame_dir = ttk.Frame(dir_section)
        btn_frame_dir.pack(fill=tk.X)
        
        ttk.Button(btn_frame_dir, text="选择目录", command=self.select_directory).pack(side=tk.LEFT, expand=True, padx=(0,5))
        ttk.Button(btn_frame_dir, text="分析", command=self.analyze).pack(side=tk.LEFT, expand=True)
        
        # 拖放区域
        drop_frame = ttk.LabelFrame(control_panel, text="拖放目录", padding="10")
        drop_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.drop_label = ttk.Label(drop_frame, text="📁 将文件夹拖放到此处", font=('Microsoft YaHei UI', 9, 'italic'), relief=tk.GROOVE, padding=10, anchor=tk.CENTER)
        self.drop_label.pack(fill=tk.X)
        
        if self.dnd_supported:
            self.drop_label.drop_target_register(DND_FILES)
            self.drop_label.dnd_bind('<<Drop>>', self.handle_drop)
        else:
            self.drop_label.config(text="⚠ 拖放功能不可用 (需要安装 tkinterdnd2)", background='#FFEBEE')

        # 后缀配置区域
        suffix_section = ttk.LabelFrame(control_panel, text="后缀配置", padding="10")
        suffix_section.pack(fill=tk.X, pady=(0, 10))
        
        suffix_configs = [
            ("空子文件夹后缀:", self.empty_suffix_var),
            ("部分空父文件夹后缀:", self.parent_partial_suffix_var),
            ("全部空父文件夹后缀:", self.parent_all_empty_suffix_var)
        ]
        
        for i, (label_text, var) in enumerate(suffix_configs):
            row_frame = ttk.Frame(suffix_section)
            row_frame.pack(fill=tk.X, pady=2)
            ttk.Label(row_frame, text=label_text).pack(side=tk.LEFT)
            ttk.Entry(row_frame, textvariable=var, width=15).pack(side=tk.RIGHT, fill=tk.X, expand=True)
        
        # 选项区域
        options_section = ttk.LabelFrame(control_panel, text="处理选项", padding="10")
        options_section.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Checkbutton(options_section, text="自动分析选定目录", variable=self.var_auto_analyze).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(options_section, text="移除正确状态的后缀", variable=self.var_remove_correct_suffix).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(options_section, text="执行前创建备份", variable=self.var_create_backup).pack(anchor=tk.W, pady=2)

        # 右侧内容区域 (Notebook)
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # 预览标签页
        self.preview_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.preview_frame, text="📋 预览")
        self._setup_preview_tab(self.preview_frame)

        # 统计标签页
        self.stats_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.stats_frame, text="📊 统计")
        self._setup_stats_tab(self.stats_frame)

        # 日志标签页
        self.log_frame_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.log_frame_tab, text="📝 日志")
        self._setup_log_tab(self.log_frame_tab)
        
        # 右下角操作按钮区域
        action_panel = ttk.Frame(main_frame)
        action_panel.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))

        ttk.Button(action_panel, text="导出报告", command=self.export_report).pack(side=tk.RIGHT, padx=(5, 0))
        self.btn_undo = ttk.Button(action_panel, text="撤销操作", command=self.undo, state='disabled')
        self.btn_undo.pack(side=tk.RIGHT, padx=(5, 0))
        self.btn_process = ttk.Button(action_panel, text="执行处理", command=self.process, state='disabled')
        self.btn_process.pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(action_panel, text="生成预览", command=self.generate_preview).pack(side=tk.RIGHT)

        # 绑定选项变化事件
        self.empty_suffix_var.trace_add("write", self._on_suffix_changed)
        self.parent_partial_suffix_var.trace_add("write", self._on_suffix_changed)
        self.parent_all_empty_suffix_var.trace_add("write", self._on_suffix_changed)
        self.var_remove_correct_suffix.trace_add("write", self._on_options_changed)
        self.var_auto_analyze.trace_add("write", self._on_options_changed)

    def _setup_preview_tab(self, parent):
        """设置预览标签页的组件"""
        filter_frame = ttk.Frame(parent)
        filter_frame.pack(fill=tk.X, pady=(0, 5))
        
        ttk.Label(filter_frame, text="筛选类型:").pack(side=tk.LEFT, padx=(0, 5))
        filter_combo = ttk.Combobox(filter_frame, textvariable=self.filter_var,
                                   values=["全部", "需要添加后缀", "需要移除后缀", "需要更正后缀", "无需更改"],
                                   state="readonly", width=20)
        filter_combo.pack(side=tk.LEFT)
        filter_combo.bind("<<ComboboxSelected>>", self.filter_preview_display)
        
        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        scroll_y = ttk.Scrollbar(tree_frame)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.preview_tree = ttk.Treeview(tree_frame, 
                                         columns=('类型', '当前名称', '新名称', '当前状态', '操作'),
                                         yscrollcommand=scroll_y.set,
                                         xscrollcommand=scroll_x.set,
                                         selectmode='extended')
        
        scroll_y.config(command=self.preview_tree.yview)
        scroll_x.config(command=self.preview_tree.xview)
        
        self.preview_tree.heading('#0', text='完整路径')
        self.preview_tree.heading('类型', text='类型')
        self.preview_tree.heading('当前名称', text='当前名称')
        self.preview_tree.heading('新名称', text='新名称')
        self.preview_tree.heading('当前状态', text='当前状态')
        self.preview_tree.heading('操作', text='操作')
        
        self.preview_tree.column('#0', width=250, minwidth=200, stretch=tk.NO)
        self.preview_tree.column('类型', width=80, anchor=tk.CENTER, stretch=tk.NO)
        self.preview_tree.column('当前名称', width=150, stretch=tk.NO)
        self.preview_tree.column('新名称', width=150, stretch=tk.NO)
        self.preview_tree.column('当前状态', width=100, anchor=tk.CENTER, stretch=tk.NO)
        self.preview_tree.column('操作', width=120, anchor=tk.CENTER, stretch=tk.NO)
        
        self.preview_tree.pack(fill=tk.BOTH, expand=True)

    def _setup_stats_tab(self, parent):
        """设置统计标签页的组件"""
        self.stats_text = scrolledtext.ScrolledText(parent, wrap=tk.WORD, font=('Consolas', 10), padx=5, pady=5)
        self.stats_text.pack(fill=tk.BOTH, expand=True)

    def _setup_log_tab(self, parent):
        """设置日志标签页的组件"""
        log_toolbar = ttk.Frame(parent)
        log_toolbar.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(log_toolbar, text="清空日志", command=self.clear_log).pack(side=tk.LEFT, padx=2)
        
        self.log_text = scrolledtext.ScrolledText(parent, wrap=tk.WORD, font=('Consolas', 9), padx=5, pady=5)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def _log_message(self, msg, level="INFO"):
        """
        用于接收 FolderManager 模块的日志消息并在UI的日志区域显示。
        这是注入给 FolderManager 的日志回调函数。
        """
        levels = {
            "INFO": ("ℹ️", "#2196F3"),
            "SUCCESS": ("✅", "#4CAF50"),
            "WARNING": ("⚠️", "#FF9800"),
            "ERROR": ("❌", "#F44336")
        }
        
        icon, color = levels.get(level, ("ℹ️", "#2196F3"))
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        log_msg = f"[{timestamp}] {icon} {msg}\n"
        
        # 在主线程中更新UI
        self.app_instance.root.after(0, lambda: self._append_log_text(log_msg, color))
        self.app_instance.root.after(0, lambda: self.app_instance.update_status(f"文件夹管理: {msg}", level.lower()))


    def _append_log_text(self, log_msg, color):
        """向日志文本框追加内容并设置颜色"""
        self.log_text.insert(tk.END, log_msg)
        # 可以根据需要为不同级别的日志添加颜色标签
        # self.log_text.tag_config(color_tag, foreground=color)
        # self.log_text.tag_add(color_tag, "end-1c linestart", "end-1c lineend")
        self.log_text.see(tk.END) # 滚动到最新日志

    def select_directory(self):
        """选择目录"""
        directory = filedialog.askdirectory(title="选择要处理的目录")
        if directory:
            self.set_current_directory(directory)
    
    def set_current_directory(self, directory):
        """设置当前目录并触发分析"""
        self.current_dir = directory
        display_dir = directory
        if len(directory) > 50: # 限制显示长度
            display_dir = "..." + directory[-47:]
        
        self.dir_label.config(text=display_dir)
        self._log_message(f"已选择目录: {directory}", "INFO")
        
        # 清空预览和统计
        self.clear_preview_display()
        self.stats_text.delete(1.0, tk.END)
        
        # 自动分析
        if self.var_auto_analyze.get():
            self.analyze()

    def handle_drop(self, event):
        """处理拖放事件"""
        if self.dnd_supported:
            path = event.data.strip('{}')
            if os.path.isdir(path):
                self.set_current_directory(path)
            else:
                messagebox.showwarning("提示", "请拖放一个有效的文件夹。")
                self._log_message("拖放失败: 请拖放一个有效的文件夹。", "WARNING")
        else:
            messagebox.showwarning("提示", "拖放功能不可用。")

    def _on_suffix_changed(self, *args):
        """后缀配置变化时更新FolderManager并尝试重新分析"""
        self.folder_manager.set_suffixes(
            self.empty_suffix_var.get(),
            self.parent_partial_suffix_var.get(),
            self.parent_all_empty_suffix_var.get()
        )
        if self.var_auto_analyze.get() and self.current_dir:
            self.analyze()
        self._log_message("后缀配置已更新。", "INFO")

    def _on_options_changed(self, *args):
        """选项变化时尝试重新分析"""
        if self.var_auto_analyze.get() and self.current_dir:
            self.analyze()
        self._log_message("处理选项已更新。", "INFO")

    def analyze(self):
        """分析文件夹结构并在统计标签页显示结果"""
        if not self.current_dir:
            messagebox.showwarning("提示", "请先选择目录。")
            return
        
        if self.processing:
            messagebox.showwarning("警告", "已有任务在处理中，请稍候。")
            return

        self.processing = True
        self.app_instance.update_status("正在分析文件夹结构...", "info")
        self.app_instance.show_progress("正在分析文件夹结构...")
        threading.Thread(target=self._analyze_thread, daemon=True).start()

    def _analyze_thread(self):
        """后台分析线程"""
        try:
            stats = self.folder_manager.analyze_folder_structure(
                self.current_dir, 
                self.var_remove_correct_suffix.get()
            )
            if stats:
                self.app_instance.root.after(0, lambda: self.display_stats(stats))
                self.app_instance.root.after(0, lambda: self.notebook.select(self.stats_frame)) # 切换到统计标签页
                self._log_message(f"分析完成，共 {stats['total_parents']} 个父文件夹，{stats['total_subfolders']} 个子文件夹。", "SUCCESS")
            else:
                self._log_message("文件夹分析失败。", "ERROR")
        except Exception as e:
            self._log_message(f"分析过程中发生错误: {str(e)}", "ERROR")
        finally:
            self.processing = False
            self.app_instance.root.after(0, self.app_instance.hide_progress)

    def display_stats(self, stats):
        """在统计文本框中显示分析结果"""
        self.stats_text.delete(1.0, tk.END)
        if stats:
            report_content = self.folder_manager._format_analysis_report(stats)
            self.stats_text.insert(tk.END, report_content)

    def generate_preview(self):
        """生成重命名操作的预览"""
        if not self.current_dir:
            messagebox.showwarning("提示", "请先选择目录。")
            return
        
        if self.processing:
            messagebox.showwarning("警告", "已有任务在处理中，请稍候。")
            return

        self.processing = True
        self.app_instance.update_status("正在生成重命名预览...", "info")
        self.app_instance.show_progress("正在生成重命名预览...")
        threading.Thread(target=self._generate_preview_thread, daemon=True).start()

    def _generate_preview_thread(self):
        """后台生成预览线程"""
        try:
            self.preview_data = self.folder_manager.get_rename_preview(
                self.current_dir,
                self.var_remove_correct_suffix.get()
            )
            self.app_instance.root.after(0, self.display_preview_data)
            self.app_instance.root.after(0, lambda: self.notebook.select(self.preview_frame)) # 切换到预览标签页
            if self.preview_data:
                self.btn_process.config(state='normal') # 有预览数据才启用执行按钮
                self._log_message(f"预览生成完成，共发现 {len(self.preview_data)} 个需要处理的文件夹。", "SUCCESS")
            else:
                self.btn_process.config(state='disabled')
                self._log_message("没有发现需要重命名的文件夹。", "INFO")
        except Exception as e:
            self._log_message(f"生成预览过程中发生错误: {str(e)}", "ERROR")
        finally:
            self.processing = False
            self.app_instance.root.after(0, self.app_instance.hide_progress)

    def display_preview_data(self):
        """在预览树形视图中显示预览数据"""
        self.clear_preview_display()
        
        for item in self.preview_data:
            # 根据操作类型设置标签，用于Treeview的样式
            tag = item['operation'].replace(' ', '_') # 移除空格，作为tag
            self.preview_tree.insert('', 'end', 
                text=item['old_path'],
                values=(item['type'], item['old_name'], item['new_name'], 
                       item['current_status'], item['operation']),
                tags=(tag,))
        
        # 配置标签样式
        self.preview_tree.tag_configure('需要添加后缀', background='#E8F5E9') # 浅绿色
        self.preview_tree.tag_configure('需要移除后缀', background='#FFF3E0') # 浅黄色
        self.preview_tree.tag_configure('需要更正后缀', background='#E3F2FD') # 浅蓝色
        self.preview_tree.tag_configure('无需更改', background='#F5F5F5') # 浅灰色

        self.filter_preview_display() # 应用当前筛选

    def filter_preview_display(self, event=None):
        """根据筛选条件显示预览列表"""
        filter_type = self.filter_var.get()
        for item_id in self.preview_tree.get_children():
            item_values = self.preview_tree.item(item_id, 'values')
            operation = item_values[4] # '操作' 列
            
            if filter_type == "全部":
                self.preview_tree.item(item_id, open=True) # 重新显示所有项
            else:
                if filter_type in operation: # 简单匹配，如果操作包含筛选文本
                    self.preview_tree.item(item_id, open=True)
                else:
                    self.preview_tree.item(item_id, open=False) # 隐藏不匹配的项

    def clear_preview_display(self):
        """清空预览显示区域"""
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        self.filter_var.set("全部") # 重置筛选

    def process(self):
        """执行重命名操作"""
        if not self.preview_data:
            messagebox.showwarning("提示", "请先生成预览。")
            return
        
        if self.processing:
            messagebox.showwarning("警告", "已有任务在处理中，请稍候。")
            return
        
        if not messagebox.askyesno("确认执行", f"即将对 {len(self.preview_data)} 个文件夹执行重命名操作，是否继续？"):
            return

        self.processing = True
        self.btn_process.config(state='disabled')
        self.btn_undo.config(state='disabled') # 执行时禁用撤销
        self.app_instance.update_status("正在执行重命名操作...", "info")
        self.app_instance.show_progress("正在执行重命名操作...")
        threading.Thread(target=self._process_thread, daemon=True).start()

    def _process_thread(self):
        """后台执行重命名线程"""
        try:
            success_count, failed_count = self.folder_manager.execute_renames(
                self.preview_data, 
                self.current_dir, 
                self.var_create_backup.get(),
                progress_callback=self.app_instance.update_progress
            )
            
            msg = f"重命名操作完成！成功: {success_count}, 失败: {failed_count}"
            self.app_instance.root.after(0, lambda: messagebox.showinfo("完成", msg))
            self._log_message(msg, "SUCCESS" if success_count > 0 else "WARNING")

            # 重新分析和生成预览
            self.app_instance.root.after(0, self.analyze)
            self.app_instance.root.after(0, self.generate_preview) # 重新生成预览以反映最新状态
            
            if self.folder_manager.rename_history:
                self.app_instance.root.after(0, lambda: self.btn_undo.config(state='normal'))

        except Exception as e:
            self.app_instance.root.after(0, lambda: messagebox.showerror("错误", f"执行重命名失败: {str(e)}"))
            self._log_message(f"执行重命名失败: {str(e)}", "ERROR")
        finally:
            self.processing = False
            self.app_instance.root.after(0, lambda: self.btn_process.config(state='normal'))
            self.app_instance.root.after(0, self.app_instance.hide_progress)

    def undo(self):
        """撤销操作"""
        if not self.folder_manager.rename_history:
            messagebox.showwarning("提示", "没有可撤销的操作。")
            return
        
        if self.processing:
            messagebox.showwarning("警告", "已有任务在处理中，请稍候。")
            return
        
        if not messagebox.askyesno("确认撤销", f"确定要撤销最近的 {len(self.folder_manager.rename_history)} 个重命名操作吗？"):
            return
        
        self.processing = True
        self.btn_undo.config(state='disabled')
        self.btn_process.config(state='disabled') # 撤销时禁用执行
        self.app_instance.update_status("正在撤销重命名操作...", "info")
        self.app_instance.show_progress("正在撤销重命名操作...")
        threading.Thread(target=self._undo_thread, daemon=True).start()

    def _undo_thread(self):
        """后台撤销线程"""
        try:
            success_count, failed_count = self.folder_manager.undo_last_operations(
                progress_callback=self.app_instance.update_progress
            )
            msg = f"撤销操作完成！成功撤销: {success_count}, 失败: {failed_count}"
            self.app_instance.root.after(0, lambda: messagebox.showinfo("完成", msg))
            self._log_message(msg, "SUCCESS" if success_count > 0 else "WARNING")

            # 重新分析和生成预览
            self.app_instance.root.after(0, self.analyze)
            self.app_instance.root.after(0, self.generate_preview) # 重新生成预览以反映最新状态

            if self.folder_manager.rename_history:
                self.app_instance.root.after(0, lambda: self.btn_undo.config(state='normal'))
            
        except Exception as e:
            self.app_instance.root.after(0, lambda: messagebox.showerror("错误", f"撤销操作失败: {str(e)}"))
            self._log_message(f"撤销操作失败: {str(e)}", "ERROR")
        finally:
            self.processing = False
            self.app_instance.root.after(0, lambda: self.btn_process.config(state='normal'))
            self.app_instance.root.after(0, self.app_instance.hide_progress)

    def clear_log(self):
        """清空日志文本框"""
        self.log_text.delete(1.0, tk.END)
        self._log_message("日志已清空。", "INFO")

    def export_report(self):
        """导出分析报告"""
        if not self.current_dir:
            messagebox.showwarning("提示", "请先选择目录并分析。")
            return
        
        report_content = self.stats_text.get(1.0, tk.END)
        if not report_content.strip():
            messagebox.showwarning("提示", "没有可导出的报告内容，请先进行分析。")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
            initialfile=f"folder_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(report_content)
                messagebox.showinfo("成功", f"报告已导出到: {file_path}")
                self._log_message(f"报告已导出到: {file_path}", "SUCCESS")
            except Exception as e:
                messagebox.showerror("错误", f"导出报告失败: {str(e)}")
                self._log_message(f"导出报告失败: {str(e)}", "ERROR")