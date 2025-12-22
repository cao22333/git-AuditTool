import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os

from core.csv_processor import FileProcessor, DataMerger, DataSummarizer, DataFilter

class CsvProcessorUI:
    """
    CSV 文件处理功能的用户界面。
    集成了文件合并、数据汇总和数据筛选三个子功能。
    """
    def __init__(self, parent_frame, app_instance):
        self.parent_frame = parent_frame
        self.app_instance = app_instance # 主应用程序实例，用于调用update_progress等方法
        
        # 初始化功能模块
        self.file_processor = FileProcessor()
        self.merger = DataMerger(self.file_processor)
        self.summarizer = DataSummarizer(self.file_processor)
        self.filter = DataFilter(self.file_processor)
        
        # 状态变量
        self.processing = False
        
        # 初始化输入变量
        self._init_variables()
        
        # 搭建界面
        self._setup_ui()

    def _init_variables(self):
        """初始化所有用户输入变量"""
        # 多文件合并相关
        self.merge_files_var = tk.StringVar()
        self.merge_encoding_var = tk.StringVar(value="auto")
        self.merge_delimiter_var = tk.StringVar(value="auto")
        self.merge_output_var = tk.StringVar()
        self.merge_chunk_processing_var = tk.BooleanVar(value=False)
        self.merge_chunk_size_var = tk.StringVar(value="50000")
        
        # 数据汇总相关
        self.summary_file_var = tk.StringVar()
        self.summary_encoding_var = tk.StringVar(value="auto")
        self.summary_delimiter_var = tk.StringVar(value="auto")
        self.summary_chunk_processing_var = tk.BooleanVar(value=False)
        self.summary_chunk_size_var = tk.StringVar(value="10000")
        self.group_var = tk.StringVar()
        self.descending_var = tk.BooleanVar(value=False)  # 降序排列
        
        # 数据筛选相关
        self.filter_data_file_var = tk.StringVar()
        self.filter_condition_file_var = tk.StringVar()
        self.filter_encoding_var = tk.StringVar(value="auto")
        self.filter_delimiter_var = tk.StringVar(value="auto")
        self.filter_column_var = tk.StringVar()
        self.filter_chunk_processing_var = tk.BooleanVar(value=False)
        
    def _setup_ui(self):
        """搭建主界面"""
        # 创建一个 Notebook (标签页) 来组织三个功能
        self.notebook = ttk.Notebook(self.parent_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建并添加各个功能标签页
        self._create_merge_tab()
        self._create_summary_tab()
        self._create_filter_tab()

    def _create_merge_tab(self):
        """创建多文件合并标签页"""
        merge_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(merge_frame, text="📁 多文件合并")
        
        # 文件选择区域
        file_frame = ttk.LabelFrame(merge_frame, text="选择CSV文件 (可多选)", padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        
        file_select_frame = ttk.Frame(file_frame)
        file_select_frame.pack(fill=tk.X, expand=True, pady=5)
        ttk.Entry(file_select_frame, textvariable=self.merge_files_var, width=80).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        ttk.Button(file_select_frame, text="浏览", command=self.browse_merge_files).pack(side=tk.RIGHT)
        
        # 高级设置
        advanced_frame = ttk.LabelFrame(merge_frame, text="高级文件设置", padding="10")
        advanced_frame.pack(fill=tk.X, pady=5)
        
        # 编码和分隔符
        enc_delim_frame = ttk.Frame(advanced_frame)
        enc_delim_frame.pack(fill=tk.X, pady=5)
        ttk.Label(enc_delim_frame, text="文件编码:").pack(side=tk.LEFT, padx=(0,5))
        ttk.Combobox(enc_delim_frame, textvariable=self.merge_encoding_var,
                    values=["auto", "utf-8", "gbk", "gb2312", "gb18030", "utf-8-sig"],
                    state="readonly", width=15).pack(side=tk.LEFT, padx=(0,15))
        ttk.Label(enc_delim_frame, text="分隔符:").pack(side=tk.LEFT, padx=(0,5))
        ttk.Combobox(enc_delim_frame, textvariable=self.merge_delimiter_var,
                    values=["auto", ",", ";", "\t", "|", " "], # 增加空格作为分隔符选项
                    state="readonly", width=10).pack(side=tk.LEFT, padx=(0,5))
        
        # 分块设置
        chunk_frame = ttk.Frame(advanced_frame)
        chunk_frame.pack(fill=tk.X, pady=5)
        ttk.Checkbutton(chunk_frame, text="启用大文件分块处理", variable=self.merge_chunk_processing_var).pack(side=tk.LEFT)
        ttk.Label(chunk_frame, text="分块大小:").pack(side=tk.LEFT, padx=(20,5))
        ttk.Entry(chunk_frame, textvariable=self.merge_chunk_size_var, width=10).pack(side=tk.LEFT)
        ttk.Label(chunk_frame, text="行").pack(side=tk.LEFT, padx=2)
        
        # 输出设置
        output_frame = ttk.LabelFrame(merge_frame, text="输出设置", padding="10")
        output_frame.pack(fill=tk.X, pady=5)
        
        output_select_frame = ttk.Frame(output_frame)
        output_select_frame.pack(fill=tk.X, expand=True, pady=5)
        ttk.Entry(output_select_frame, textvariable=self.merge_output_var, width=80).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        ttk.Button(output_select_frame, text="浏览", command=self.browse_merge_output).pack(side=tk.RIGHT)
        
        # 处理按钮
        ttk.Button(merge_frame, text="开始合并", command=self.process_merge, style='TButton').pack(pady=10)

    def _create_summary_tab(self):
        """创建数据汇总标签页"""
        summary_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(summary_frame, text="📊 数据汇总")
        
        # 步骤1：文件选择
        self.summary_step1_frame = ttk.LabelFrame(summary_frame, text="步骤1: 选择CSV文件", padding="10")
        self.summary_step1_frame.pack(fill=tk.X, pady=5)
        
        file_select_frame = ttk.Frame(self.summary_step1_frame)
        file_select_frame.pack(fill=tk.X, expand=True, pady=5)
        ttk.Entry(file_select_frame, textvariable=self.summary_file_var, width=80).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        ttk.Button(file_select_frame, text="浏览", command=lambda: self.browse_file(self.summary_file_var, [("CSV文件", "*.csv")])).pack(side=tk.RIGHT)
        
        # 高级设置
        advanced_frame = ttk.LabelFrame(self.summary_step1_frame, text="高级文件设置", padding="10")
        advanced_frame.pack(fill=tk.X, pady=5)
        
        # 编码和分隔符
        enc_delim_frame = ttk.Frame(advanced_frame)
        enc_delim_frame.pack(fill=tk.X, pady=5)
        ttk.Label(enc_delim_frame, text="文件编码:").pack(side=tk.LEFT, padx=(0,5))
        ttk.Combobox(enc_delim_frame, textvariable=self.summary_encoding_var,
                    values=["auto", "utf-8", "gbk", "gb2312", "gb18030", "utf-8-sig"],
                    state="readonly", width=15).pack(side=tk.LEFT, padx=(0,15))
        ttk.Label(enc_delim_frame, text="分隔符:").pack(side=tk.LEFT, padx=(0,5))
        ttk.Combobox(enc_delim_frame, textvariable=self.summary_delimiter_var,
                    values=["auto", ",", ";", "\t", "|", " "],
                    state="readonly", width=10).pack(side=tk.LEFT, padx=(0,5))
        
        # 分块设置
        chunk_frame = ttk.Frame(advanced_frame)
        chunk_frame.pack(fill=tk.X, pady=5)
        ttk.Checkbutton(chunk_frame, text="启用大文件分块处理", variable=self.summary_chunk_processing_var).pack(side=tk.LEFT)
        ttk.Label(chunk_frame, text="分块大小:").pack(side=tk.LEFT, padx=(20,5))
        ttk.Entry(chunk_frame, textvariable=self.summary_chunk_size_var, width=10).pack(side=tk.LEFT)
        ttk.Label(chunk_frame, text="行").pack(side=tk.LEFT, padx=2)
        
        # 下一步按钮
        ttk.Button(self.summary_step1_frame, text="下一步 →", command=self.load_summary_columns).pack(pady=5)
        
        # 步骤2：列选择（初始隐藏）
        self.summary_step2_frame = ttk.LabelFrame(summary_frame, text="步骤2: 选择分组列和求和列", padding="10")
        
        # 分组列
        group_frame = ttk.Frame(self.summary_step2_frame)
        group_frame.pack(fill=tk.X, pady=5)
        ttk.Label(group_frame, text="分组列:").pack(side=tk.LEFT, padx=(0,5))
        self.group_combo = ttk.Combobox(group_frame, textvariable=self.group_var, state="readonly", width=40)
        self.group_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        ttk.Checkbutton(group_frame, text="降序排列结果", variable=self.descending_var).pack(side=tk.RIGHT)
        
        # 求和列（多选）
        sum_frame = ttk.Frame(self.summary_step2_frame)
        sum_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        ttk.Label(sum_frame, text="求和列 (可多选):").pack(anchor=tk.W, pady=(0,5))
        
        listbox_frame = ttk.Frame(sum_frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True)
        self.sum_listbox = tk.Listbox(listbox_frame, selectmode=tk.MULTIPLE, height=8, exportselection=False) # exportselection=False 避免Listbox失去焦点时清除选择
        self.sum_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(listbox_frame, command=self.sum_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.sum_listbox.config(yscrollcommand=scrollbar.set)
        
        # 按钮
        btn_frame = ttk.Frame(self.summary_step2_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        ttk.Button(btn_frame, text="← 上一步", command=self.back_to_summary_step1).pack(side=tk.LEFT, padx=(0,5))
        ttk.Button(btn_frame, text="开始汇总", command=self.process_summary, style='TButton').pack(side=tk.LEFT)

    def _create_filter_tab(self):
        """创建数据筛选标签页"""
        filter_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(filter_frame, text="🔍 数据筛选")
        
        # 步骤1：文件选择
        self.filter_step1_frame = ttk.LabelFrame(filter_frame, text="步骤1: 选择文件", padding="10")
        self.filter_step1_frame.pack(fill=tk.X, pady=5)
        
        # 数据文件
        data_file_frame = ttk.Frame(self.filter_step1_frame)
        data_file_frame.pack(fill=tk.X, pady=5)
        ttk.Label(data_file_frame, text="数据CSV文件:").pack(anchor=tk.W, padx=(0,5))
        data_entry_frame = ttk.Frame(data_file_frame)
        data_entry_frame.pack(fill=tk.X, expand=True, pady=5)
        ttk.Entry(data_entry_frame, textvariable=self.filter_data_file_var, width=80).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        ttk.Button(data_entry_frame, text="浏览", command=lambda: self.browse_file(self.filter_data_file_var, [("CSV文件", "*.csv")])).pack(side=tk.RIGHT)
        
        # 筛选条件文件
        filter_file_frame = ttk.Frame(self.filter_step1_frame)
        filter_file_frame.pack(fill=tk.X, pady=5)
        ttk.Label(filter_file_frame, text="筛选条件Excel文件 (第一列为筛选值):").pack(anchor=tk.W, padx=(0,5))
        filter_entry_frame = ttk.Frame(filter_file_frame)
        filter_entry_frame.pack(fill=tk.X, expand=True, pady=5)
        ttk.Entry(filter_entry_frame, textvariable=self.filter_condition_file_var, width=80).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        ttk.Button(filter_entry_frame, text="浏览", command=lambda: self.browse_file(self.filter_condition_file_var, [("Excel文件", "*.xlsx")])).pack(side=tk.RIGHT)
        
        # 高级设置
        advanced_frame = ttk.LabelFrame(self.filter_step1_frame, text="高级文件设置", padding="10")
        advanced_frame.pack(fill=tk.X, pady=5)
        
        # 编码和分隔符
        enc_delim_frame = ttk.Frame(advanced_frame)
        enc_delim_frame.pack(fill=tk.X, pady=5)
        ttk.Label(enc_delim_frame, text="文件编码:").pack(side=tk.LEFT, padx=(0,5))
        ttk.Combobox(enc_delim_frame, textvariable=self.filter_encoding_var,
                    values=["auto", "utf-8", "gbk", "gb2312", "gb18030", "utf-8-sig"],
                    state="readonly", width=15).pack(side=tk.LEFT, padx=(0,15))
        ttk.Label(enc_delim_frame, text="分隔符:").pack(side=tk.LEFT, padx=(0,5))
        ttk.Combobox(enc_delim_frame, textvariable=self.filter_delimiter_var,
                    values=["auto", ",", ";", "\t", "|", " "],
                    state="readonly", width=10).pack(side=tk.LEFT, padx=(0,5))
        
        # 分块设置
        ttk.Checkbutton(advanced_frame, text="启用大文件分块处理", variable=self.filter_chunk_processing_var).pack(anchor=tk.W, pady=5)
        
        # 下一步按钮
        ttk.Button(self.filter_step1_frame, text="下一步 →", command=self.load_filter_columns).pack(pady=10)
        
        # 步骤2：筛选列选择（初始隐藏）
        self.filter_step2_frame = ttk.LabelFrame(filter_frame, text="步骤2: 选择筛选列", padding="10")
        
        # 筛选列
        col_frame = ttk.Frame(self.filter_step2_frame)
        col_frame.pack(fill=tk.X, pady=5)
        ttk.Label(col_frame, text="筛选列:").pack(side=tk.LEFT, padx=(0,5))
        self.filter_column_combo = ttk.Combobox(col_frame, textvariable=self.filter_column_var, state="readonly", width=40)
        self.filter_column_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0,5))
        
        # 按钮
        btn_frame = ttk.Frame(self.filter_step2_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        ttk.Button(btn_frame, text="← 上一步", command=self.back_to_filter_step1).pack(side=tk.LEFT, padx=(0,5))
        ttk.Button(btn_frame, text="开始筛选", command=self.process_filter, style='TButton').pack(side=tk.LEFT)

    # ------------------------------
    # 通用UI组件和方法
    # ------------------------------
    def browse_file(self, file_var, filetypes):
        """文件选择对话框"""
        filename = filedialog.askopenfilename(filetypes=filetypes)
        if filename:
            file_var.set(filename)
            return True
        return False
    
    # ------------------------------
    # 多文件合并相关方法
    # ------------------------------
    def browse_merge_files(self):
        """选择多个合并文件"""
        files = filedialog.askopenfilenames(filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")])
        if files:
            self.merge_files_var.set(";".join(files))
            
    def browse_merge_output(self):
        """选择合并输出路径"""
        file = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV文件", "*.csv"), ("所有文件", "*.*")])
        if file:
            self.merge_output_var.set(file)

    def process_merge(self):
        """启动合并线程"""
        if not self._validate_merge_inputs():
            return
        
        if self.processing:
            messagebox.showwarning("警告", "已有任务在处理中，请稍候。")
            return
        
        self.processing = True
        self.app_instance.update_status("正在合并文件...", "info")
        self.app_instance.show_progress("正在合并文件...")
        threading.Thread(target=self._merge_thread, daemon=True).start()

    def _merge_thread(self):
        """合并后台线程"""
        try:
            file_paths = self.merge_files_var.get().split(";")
            output_path = self.merge_output_var.get()
            use_chunking = self.merge_chunk_processing_var.get()
            chunk_size = int(self.merge_chunk_size_var.get()) if use_chunking else 50000
            
            # 调用合并模块
            success, msg = self.merger.merge_files(
                file_paths=file_paths,
                output_path=output_path,
                encoding=self.merge_encoding_var.get(),
                delimiter=self.merge_delimiter_var.get(),
                use_chunking=use_chunking,
                chunk_size=chunk_size,
                progress_callback=self.app_instance.update_progress
            )
            
            if success:
                self.app_instance.root.after(0, lambda: messagebox.showinfo("成功", msg))
                self.app_instance.update_status("合并完成", "success")
            else:
                self.app_instance.root.after(0, lambda: messagebox.showerror("错误", msg))
                self.app_instance.update_status("合并失败", "error")
        except Exception as e:
            self.app_instance.root.after(0, lambda: messagebox.showerror("错误", f"合并失败: {str(e)}"))
            self.app_instance.update_status("合并失败", "error")
        finally:
            self.processing = False
            self.app_instance.root.after(0, self.app_instance.hide_progress)

    def _validate_merge_inputs(self):
        """验证合并输入"""
        if not self.merge_files_var.get():
            messagebox.showerror("错误", "请选择要合并的CSV文件。")
            return False
        if not self.merge_output_var.get():
            messagebox.showerror("错误", "请选择合并结果的输出路径。")
            return False
        if self.merge_chunk_processing_var.get():
            try:
                if int(self.merge_chunk_size_var.get()) <= 0:
                    messagebox.showerror("错误", "分块大小必须为正整数。")
                    return False
            except ValueError:
                messagebox.showerror("错误", "分块大小必须为数字。")
                return False
        return True

    # ------------------------------
    # 数据汇总相关方法
    # ------------------------------
    def load_summary_columns(self):
        """加载文件列信息"""
        file_path = self.summary_file_var.get()
        if not file_path:
            messagebox.showerror("错误", "请选择CSV文件。")
            return
        
        if self.processing:
            messagebox.showwarning("警告", "已有任务在处理中，请稍候。")
            return

        self.processing = True
        self.app_instance.update_status("正在读取文件列信息...", "info")
        self.app_instance.show_progress("正在读取文件列信息...")
        threading.Thread(target=self._load_summary_columns_thread, args=(file_path,), daemon=True).start()

    def _load_summary_columns_thread(self, file_path):
        """加载列信息的后台线程"""
        try:
            # 读取文件获取列信息
            df, _, _ = self.file_processor.read_csv_robust(
                file_path, self.summary_encoding_var.get(), self.summary_delimiter_var.get()
            )
            
            if df is None:
                self.app_instance.root.after(0, lambda: messagebox.showerror("错误", "无法读取文件或文件内容为空，请检查文件格式、编码和分隔符。"))
                self.app_instance.update_status("加载列信息失败", "error")
                return
            
            # 更新UI
            columns = df.columns.tolist()
            if not columns:
                self.app_instance.root.after(0, lambda: messagebox.showerror("错误", "文件中未检测到任何列，请检查文件格式。"))
                self.app_instance.update_status("加载列信息失败", "error")
                return

            self.app_instance.root.after(0, lambda: self._show_summary_step2(columns))
            self.app_instance.update_status(f"文件加载完成，共 {len(columns)} 列", "success")
        except Exception as e:
            self.app_instance.root.after(0, lambda: messagebox.showerror("错误", f"加载列信息失败: {str(e)}"))
            self.app_instance.update_status("加载列信息失败", "error")
        finally:
            self.processing = False
            self.app_instance.root.after(0, self.app_instance.hide_progress)

    def _show_summary_step2(self, columns):
        """显示汇总步骤2"""
        self.group_combo['values'] = columns
        self.sum_listbox.delete(0, tk.END)
        for col in columns:
            self.sum_listbox.insert(tk.END, col)
        
        if columns:
            self.group_combo.set(columns[0])  # 默认选择第一列
        
        self.summary_step1_frame.pack_forget()
        self.summary_step2_frame.pack(fill=tk.BOTH, expand=True, pady=5)
    
    def back_to_summary_step1(self):
        """返回汇总步骤1"""
        self.summary_step2_frame.pack_forget()
        self.summary_step1_frame.pack(fill=tk.X, pady=5)
    
    def process_summary(self):
        """启动汇总线程"""
        if not self._validate_summary_inputs():
            return
        
        if self.processing:
            messagebox.showwarning("警告", "已有任务在处理中，请稍候。")
            return
        
        self.processing = True
        self.app_instance.update_status("正在汇总数据...", "info")
        self.app_instance.show_progress("正在汇总数据...")
        threading.Thread(target=self._summary_thread, daemon=True).start()

    def _summary_thread(self):
        """汇总后台线程"""
        try:
            # 获取用户输入
            file_path = self.summary_file_var.get()
            group_col = self.group_var.get()
            sum_cols = [self.sum_listbox.get(i) for i in self.sum_listbox.curselection()]
            
            if not sum_cols:
                self.app_instance.root.after(0, lambda: messagebox.showerror("错误", "请选择至少一个求和列。"))
                self.app_instance.update_status("汇总失败", "error")
                return

            output_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
            )
            
            if not output_path:
                self.app_instance.update_status("用户取消汇总", "warning")
                return
            
            # 调用汇总模块
            success, msg = self.summarizer.summarize(
                file_path=file_path,
                group_col=group_col,
                sum_cols=sum_cols,
                output_path=output_path,
                encoding=self.summary_encoding_var.get(),
                delimiter=self.summary_delimiter_var.get(),
                use_chunking=self.summary_chunk_processing_var.get(),
                chunk_size=int(self.summary_chunk_size_var.get()),
                descending=self.descending_var.get(),
                progress_callback=self.app_instance.update_progress
            )
            
            if success:
                self.app_instance.root.after(0, lambda: messagebox.showinfo("成功", msg))
                self.app_instance.update_status("汇总完成", "success")
            else:
                self.app_instance.root.after(0, lambda: messagebox.showerror("错误", msg))
                self.app_instance.update_status("汇总失败", "error")
        except Exception as e:
            self.app_instance.root.after(0, lambda: messagebox.showerror("错误", f"汇总失败: {str(e)}"))
            self.app_instance.update_status("汇总失败", "error")
        finally:
            self.processing = False
            self.app_instance.root.after(0, self.app_instance.hide_progress)

    def _validate_summary_inputs(self):
        """验证汇总输入"""
        if not self.summary_file_var.get():
            messagebox.showerror("错误", "请选择CSV文件。")
            return False
        if not self.group_var.get():
            messagebox.showerror("错误", "请选择分组列。")
            return False
        if not self.sum_listbox.curselection():
            messagebox.showerror("错误", "请选择至少一个求和列。")
            return False
        if self.summary_chunk_processing_var.get():
            try:
                if int(self.summary_chunk_size_var.get()) <= 0:
                    messagebox.showerror("错误", "分块大小必须为正整数。")
                    return False
            except ValueError:
                messagebox.showerror("错误", "分块大小必须为数字。")
                return False
        return True

    # ------------------------------
    # 数据筛选相关方法
    # ------------------------------
    def load_filter_columns(self):
        """加载筛选列信息"""
        data_file = self.filter_data_file_var.get()
        if not data_file:
            messagebox.showerror("错误", "请选择数据CSV文件。")
            return
        
        if not self.filter_condition_file_var.get():
            messagebox.showerror("错误", "请选择筛选条件Excel文件。")
            return

        if self.processing:
            messagebox.showwarning("警告", "已有任务在处理中，请稍候。")
            return

        self.processing = True
        self.app_instance.update_status("正在读取数据文件列信息...", "info")
        self.app_instance.show_progress("正在读取数据文件列信息...")
        threading.Thread(target=self._load_filter_columns_thread, args=(data_file,), daemon=True).start()

    def _load_filter_columns_thread(self, data_file):
        """加载筛选列的后台线程"""
        try:
            # 读取文件获取列信息
            df, _, _ = self.file_processor.read_csv_robust(
                data_file, self.filter_encoding_var.get(), self.filter_delimiter_var.get()
            )
            
            if df is None:
                self.app_instance.root.after(0, lambda: messagebox.showerror("错误", "无法读取数据文件或文件内容为空，请检查文件格式、编码和分隔符。"))
                self.app_instance.update_status("加载筛选列失败", "error")
                return
            
            # 更新UI
            columns = df.columns.tolist()
            if not columns:
                self.app_instance.root.after(0, lambda: messagebox.showerror("错误", "数据文件中未检测到任何列，请检查文件格式。"))
                self.app_instance.update_status("加载筛选列失败", "error")
                return

            self.app_instance.root.after(0, lambda: self._show_filter_step2(columns))
            self.app_instance.update_status(f"数据文件加载完成，共 {len(columns)} 列", "success")
        except Exception as e:
            self.app_instance.root.after(0, lambda: messagebox.showerror("错误", f"加载筛选列信息失败: {str(e)}"))
            self.app_instance.update_status("加载筛选列失败", "error")
        finally:
            self.processing = False
            self.app_instance.root.after(0, self.app_instance.hide_progress)

    def _show_filter_step2(self, columns):
        """显示筛选步骤2"""
        self.filter_column_combo['values'] = columns
        if columns:
            self.filter_column_combo.set(columns[0])  # 默认选择第一列
        
        self.filter_step1_frame.pack_forget()
        self.filter_step2_frame.pack(fill=tk.BOTH, expand=True, pady=5)
    
    def back_to_filter_step1(self):
        """返回筛选步骤1"""
        self.filter_step2_frame.pack_forget()
        self.filter_step1_frame.pack(fill=tk.X, pady=5)
    
    def process_filter(self):
        """启动筛选线程"""
        if not self._validate_filter_inputs():
            return
        
        if self.processing:
            messagebox.showwarning("警告", "已有任务在处理中，请稍候。")
            return
        
        self.processing = True
        self.app_instance.update_status("正在筛选数据...", "info")
        self.app_instance.show_progress("正在筛选数据...")
        threading.Thread(target=self._filter_thread, daemon=True).start()

    def _filter_thread(self):
        """筛选后台线程"""
        try:
            # 获取用户输入
            data_file = self.filter_data_file_var.get()
            filter_file = self.filter_condition_file_var.get()
            filter_col = self.filter_column_var.get()
            
            output_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
            )
            
            if not output_path:
                self.app_instance.update_status("用户取消筛选", "warning")
                return
            
            # 调用筛选模块
            success, msg = self.filter.filter_data(
                data_file=data_file,
                filter_file=filter_file,
                filter_col=filter_col,
                output_path=output_path,
                encoding=self.filter_encoding_var.get(),
                delimiter=self.filter_delimiter_var.get(),
                use_chunking=self.filter_chunk_processing_var.get(),
                progress_callback=self.app_instance.update_progress
            )
            
            if success:
                self.app_instance.root.after(0, lambda: messagebox.showinfo("成功", msg))
                self.app_instance.update_status("筛选完成", "success")
            else:
                self.app_instance.root.after(0, lambda: messagebox.showerror("错误", msg))
                self.app_instance.update_status("筛选失败", "error")
        except Exception as e:
            self.app_instance.root.after(0, lambda: messagebox.showerror("错误", f"筛选失败: {str(e)}"))
            self.app_instance.update_status("筛选失败", "error")
        finally:
            self.processing = False
            self.app_instance.root.after(0, self.app_instance.hide_progress)
    
    def _validate_filter_inputs(self):
        """验证筛选输入"""
        if not self.filter_data_file_var.get():
            messagebox.showerror("错误", "请选择数据CSV文件。")
            return False
        if not self.filter_condition_file_var.get():
            messagebox.showerror("错误", "请选择筛选条件Excel文件。")
            return False
        if not self.filter_column_var.get():
            messagebox.showerror("错误", "请选择筛选列。")
            return False
        return True