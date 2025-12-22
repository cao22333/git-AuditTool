import pandas as pd
import chardet
import csv
import os
import threading
from queue import Queue

# ==============================================================================
# 通用文件处理工具
# ==============================================================================
class FileProcessor:
    """
    通用文件处理工具类。
    负责文件的读取、编码检测、分隔符检测、文件大小获取等底层操作。
    """
    def detect_encoding(self, file_path):
        """
        自动检测文件编码。
        尝试读取文件的前50KB数据进行编码检测，如果置信度低则默认使用 'gbk'。
        """
        try:
            with open(file_path, 'rb') as f:
                raw_data = f.read(50000)  # 读取前50KB用于检测
                result = chardet.detect(raw_data)
                encoding = result.get('encoding', 'utf-8')
                confidence = result.get('confidence', 0)
                
                # 置信度低时使用常见编码
                if confidence < 0.7:
                    return 'gbk'
                return encoding
        except Exception as e:
            print(f"编码检测失败: {str(e)}")
            return 'gbk'  # 默认编码
    
    def detect_delimiter(self, file_path, encoding='utf-8'):
        """
        自动检测CSV文件分隔符。
        通过读取文件的前两行，统计常见分隔符（逗号、分号、制表符、竖线）的出现次数，
        并返回出现次数最多的分隔符。
        """
        try:
            with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
                first_line = f.readline().strip()
                second_line = f.readline().strip()
            
            # 统计常见分隔符出现次数
            delimiters = [',', ';', '\t', '|']
            delimiter_counts = {}
            for delim in delimiters:
                count1 = first_line.count(delim)
                count2 = second_line.count(delim) if second_line else 0
                delimiter_counts[delim] = (count1 + count2) / 2  # 取平均值
            
            # 返回出现次数最多的分隔符
            best_delimiter = max(delimiter_counts, key=delimiter_counts.get)
            if delimiter_counts[best_delimiter] > 0:
                return best_delimiter
            return ','  # 默认分隔符
        except Exception as e:
            print(f"分隔符检测失败: {str(e)}")
            return ','  # 默认分隔符
    
    def read_csv_robust(self, file_path, encoding=None, delimiter=None, chunksize=None):
        """
        健壮的CSV读取函数，支持自动检测编码/分隔符、分块读取。
        会尝试多种编码和分隔符组合，直到成功读取文件。
        """
        # 准备要尝试的编码和分隔符
        encodings = [encoding] if encoding and encoding != "auto" else \
                    ['utf-8-sig', 'gbk', 'gb2312', 'utf-8', 'latin1']
        delimiters = [delimiter] if delimiter and delimiter != "auto" else \
                    [',', ';', '\t', '|', ' ']
        
        # 尝试所有可能的组合
        for enc in encodings:
            for delim in delimiters:
                try:
                    if chunksize:
                        # 分块读取（大文件）
                        chunk_reader = pd.read_csv(
                            file_path,
                            encoding=enc,
                            sep=delim,
                            chunksize=chunksize,
                            engine='python',
                            on_bad_lines='skip' # 忽略错误行
                        )
                        # 验证是否能正常读取（尝试读取第一块）
                        next(chunk_reader)
                        # 重新创建读取器，因为上面的 next() 已经移动了指针
                        chunk_reader = pd.read_csv(
                            file_path,
                            encoding=enc,
                            sep=delim,
                            chunksize=chunksize,
                            engine='python',
                            on_bad_lines='skip'
                        )
                        return chunk_reader, enc, delim
                    else:
                        # 整体读取（小文件）
                        df = pd.read_csv(
                            file_path,
                            encoding=enc,
                            sep=delim,
                            engine='python',
                            on_bad_lines='skip'
                        )
                        return df, enc, delim
                except Exception:
                    continue # 出现异常则尝试下一个组合
        
        # 最终尝试宽松模式（utf-8编码，逗号分隔符）
        try:
            if chunksize:
                return pd.read_csv(file_path, encoding='utf-8', sep=',', chunksize=chunksize, engine='python', on_bad_lines='skip'), 'utf-8', ','
            else:
                return pd.read_csv(file_path, encoding='utf-8', sep=',', engine='python', on_bad_lines='skip'), 'utf-8', ','
        except Exception as e:
            print(f"最终读取CSV失败: {e}")
            return None, None, None
    
    def get_file_size(self, file_path):
        """
        获取文件大小（MB）。
        """
        try:
            size_bytes = os.path.getsize(file_path)
            return size_bytes / (1024 * 1024)  # 转换为MB
        except:
            return 0

# ==============================================================================
# 多文件合并工具
# ==============================================================================
class DataMerger:
    """
    多文件合并工具类。
    负责将多个CSV文件合并为一个，支持分块处理和编码/分隔符自动检测。
    """
    def __init__(self, file_processor):
        self.file_processor = file_processor  # 依赖注入：复用通用文件处理
    
    def merge_files(self, file_paths, output_path, encoding='auto', delimiter='auto', 
                   use_chunking=False, chunk_size=50000, progress_callback=None):
        """
        合并多个CSV文件。
        progress_callback: 进度更新回调函数（接收进度值和文本）
        """
        try:
            total_files = len(file_paths)
            if progress_callback:
                progress_callback(5, f"开始合并 {total_files} 个文件...")
            
            # 检测第一个文件的编码和分隔符作为标准
            first_file = file_paths[0]
            used_encoding, used_delimiter = self._detect_file_params(
                first_file, encoding, delimiter
            )
            
            if progress_callback:
                progress_callback(10, f"使用编码: {used_encoding}, 分隔符: '{used_delimiter}'")
            
            # 收集所有文件的列（确保合并后列完整）
            all_columns = self._collect_all_columns(
                file_paths, used_encoding, used_delimiter, progress_callback
            )
            if not all_columns:
                return False, "无法解析文件列信息"
            
            if progress_callback:
                progress_callback(30, f"确定统一结构: {len(all_columns)} 列")
            
            # 执行合并
            self._do_merge(
                file_paths, output_path, all_columns, used_encoding, used_delimiter,
                use_chunking, chunk_size, progress_callback
            )
            
            # 验证结果
            return True, f"合并完成！共 {total_files} 个文件，保存至: {output_path}"
            
        except Exception as e:
            return False, f"合并失败: {str(e)}"
    
    def _detect_file_params(self, file_path, encoding, delimiter):
        """
        检测文件编码和分隔符。
        如果编码或分隔符设置为 'auto'，则使用 FileProcessor 进行自动检测。
        """
        if encoding == "auto":
            encoding = self.file_processor.detect_encoding(file_path)
        if delimiter == "auto":
            delimiter = self.file_processor.detect_delimiter(file_path, encoding)
        return encoding, delimiter
    
    def _collect_all_columns(self, file_paths, encoding, delimiter, progress_callback):
        """
        收集所有文件的列名，以确保合并后的 DataFrame 包含所有可能的列。
        """
        all_columns = set()
        total = len(file_paths)
        
        for i, file_path in enumerate(file_paths):
            # 更新进度
            if progress_callback:
                progress = 10 + int(i / total * 20)  # 10% ~ 30%
                progress_callback(progress, f"分析文件结构 {i+1}/{total}")
            
            # 只读取前10行获取列信息
            df_sample, _, _ = self.file_processor.read_csv_robust(
                file_path, encoding, delimiter
            )
            if df_sample is not None:
                all_columns.update(df_sample.columns.tolist())
        
        return sorted(list(all_columns))
    
    def _do_merge(self, file_paths, output_path, all_columns, encoding, delimiter,
                 use_chunking, chunk_size, progress_callback):
        """
        执行实际的合并操作，将所有文件的数据写入到输出文件。
        支持分块写入以处理大文件。
        """
        total_files = len(file_paths)
        first_file_flag = True  # 标记是否需要写入表头
        
        for file_idx, file_path in enumerate(file_paths):
            # 更新进度（30% ~ 90%）
            base_progress = 30 + int(file_idx / total_files * 60)
            if progress_callback:
                progress_callback(base_progress, 
                                 f"处理文件 {file_idx+1}/{total_files}: {os.path.basename(file_path)}")
            
            if use_chunking and chunk_size:
                # 分块处理大文件
                chunks, _, _ = self.file_processor.read_csv_robust(
                    file_path, encoding, delimiter, chunksize=chunk_size
                )
                if chunks is None:
                    continue # 跳过无法读取的文件

                for chunk_idx, chunk in enumerate(chunks):
                    self._process_chunk(chunk, all_columns, output_path, first_file_flag)
                    first_file_flag = False  # 只有第一个块需要表头
                    
                    # 更新块进度
                    if progress_callback:
                        chunk_progress = min(90, base_progress + int((chunk_idx + 1) * 5))
                        progress_callback(chunk_progress, 
                                         f"处理块 {chunk_idx+1} (文件 {file_idx+1}/{total_files})")
            else:
                # 整体处理小文件
                df, _, _ = self.file_processor.read_csv_robust(
                    file_path, encoding, delimiter
                )
                if df is not None:
                    self._process_chunk(df, all_columns, output_path, first_file_flag)
                    first_file_flag = False
        
        if progress_callback:
            progress_callback(95, "合并完成，验证结果中...")
    
    def _process_chunk(self, chunk, all_columns, output_path, write_header):
        """
        处理单个数据块（补充缺失列并写入文件）。
        """
        # 补充缺失的列
        for col in all_columns:
            if col not in chunk.columns:
                chunk[col] = ""
        # 按统一列顺序排序
        chunk = chunk.reindex(columns=all_columns)
        # 写入文件
        chunk.to_csv(
            output_path,
            mode='w' if write_header else 'a', # 第一次写入带表头，后续追加不带表头
            header=write_header,
            index=False,
            encoding='utf-8-sig' # 使用带BOM的UTF-8编码，确保Excel正确识别中文
        )

# ==============================================================================
# 数据汇总工具
# ==============================================================================
class DataSummarizer:
    """
    数据汇总工具类。
    负责按指定列分组求和，支持分块处理和结果排序。
    """
    def __init__(self, file_processor):
        self.file_processor = file_processor  # 复用通用文件处理
    
    def summarize(self, file_path, group_col, sum_cols, output_path,
                 encoding='auto', delimiter='auto', use_chunking=False, 
                 chunk_size=10000, descending=False, progress_callback=None):
        """
        按指定列分组求和。
        group_col: 分组列名
        sum_cols: 需要求和的列名列表
        descending: 是否降序排列结果
        progress_callback: 进度更新回调
        """
        try:
            if not sum_cols:
                return False, "请选择至少一个求和列"
            
            file_size = self.file_processor.get_file_size(file_path)
            if progress_callback:
                progress_callback(5, f"开始汇总 ({file_size:.2f} MB)...")
            
            if use_chunking and chunk_size:
                # 分块处理大文件
                result = self._process_in_chunks(
                    file_path, group_col, sum_cols, encoding, delimiter,
                    chunk_size, progress_callback
                )
            else:
                # 整体处理小文件
                result = self._process_entire_file(
                    file_path, group_col, sum_cols, encoding, delimiter,
                    progress_callback
                )
            
            if result is None or result.empty:
                return False, "未找到有效数据或汇总结果为空"
            
            # 排序（升序/降序）
            if descending:
                # 如果有求和列，按第一个求和列排序；否则按分组列排序
                sort_col = sum_cols[0] if sum_cols else group_col
                if sort_col in result.columns:
                    result = result.sort_values(by=sort_col, ascending=False)
                    if progress_callback:
                        progress_callback(96, "按降序排列结果...")
            
            # 保存结果到Excel
            result.to_excel(output_path, index=False)
            return True, f"汇总完成！共 {len(result)} 个分组，保存至: {output_path}"
            
        except Exception as e:
            return False, f"汇总失败: {str(e)}"
    
    def _process_entire_file(self, file_path, group_col, sum_cols, encoding, delimiter, progress_callback):
        """
        整体处理小文件，读取整个文件后进行分组求和。
        """
        if progress_callback:
            progress_callback(30, "正在读取文件...")
        
        # 读取完整文件
        df, _, _ = self.file_processor.read_csv_robust(
            file_path, encoding, delimiter
        )
        if df is None:
            return None
        
        if progress_callback:
            progress_callback(60, "正在转换数据类型...")
        
        # 处理分组列（转为字符串确保一致性）
        if group_col not in df.columns:
            raise ValueError(f"分组列 '{group_col}' 不存在于文件中。")
        df[group_col] = df[group_col].astype(str)
        
        # 处理求和列（转换为数值，处理千分位）
        for col in sum_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(
                    df[col].astype(str).str.replace(',', ''),  # 移除千分位逗号
                    errors='coerce' # 无法转换的设置为 NaN
                ).fillna(0) # 将 NaN 填充为 0
            else:
                print(f"警告：求和列 '{col}' 不存在于文件中，将被忽略。")
                df[col] = 0 # 如果列不存在，添加并填充0
        
        if progress_callback:
            progress_callback(80, "正在执行分组汇总...")
        
        # 分组求和
        return df.groupby(group_col, as_index=False)[sum_cols].sum()
    
    def _process_in_chunks(self, file_path, group_col, sum_cols, encoding, delimiter,
                          chunk_size, progress_callback):
        """
        分块处理大文件，逐步读取和汇总数据，减少内存占用。
        """
        if progress_callback:
            progress_callback(10, "正在分块读取文件...")
        
        # 获取分块读取器
        chunks, _, _ = self.file_processor.read_csv_robust(
            file_path, encoding, delimiter, chunksize=chunk_size
        )
        if chunks is None:
            return None
        
        result = None
        processed_rows = 0
        chunk_count = 0
        
        for chunk in chunks:
            # 检查分组列是否存在
            if group_col not in chunk.columns:
                raise ValueError(f"分组列 '{group_col}' 不存在于当前数据块中。")
            
            # 处理分组列
            chunk[group_col] = chunk[group_col].astype(str)
            
            # 处理求和列
            for col in sum_cols:
                if col in chunk.columns:
                    chunk[col] = pd.to_numeric(
                        chunk[col].astype(str).str.replace(',', ''),
                        errors='coerce'
                    ).fillna(0)
                else:
                    print(f"警告：求和列 '{col}' 不存在于当前数据块中，将被忽略。")
                    chunk[col] = 0 # 如果列不存在，添加并填充0
            
            # 分组求和
            chunk_result = chunk.groupby(group_col, as_index=False)[sum_cols].sum()
            
            # 合并分块结果
            if result is None:
                result = chunk_result
            else:
                # 合并相同分组，对数值列进行累加
                result = pd.merge(result, chunk_result, on=group_col, how='outer', suffixes=('', '_new'))
                for col in sum_cols:
                    # 将两个DataFrame中相同组的求和列进行累加
                    result[col] = result[col].fillna(0) + result.get(f"{col}_new", pd.Series(0)).fillna(0)
                    if f"{col}_new" in result.columns:
                        result = result.drop(columns=[f"{col}_new"])
            
            # 更新进度
            processed_rows += len(chunk)
            chunk_count += 1
            progress = min(90, 10 + (chunk_count * 5))
            if progress_callback:
                progress_callback(progress, f"已处理 {processed_rows} 行数据...")
        
        return result

# ==============================================================================
# 数据筛选工具
# ==============================================================================
class DataFilter:
    """
    数据筛选工具类。
    根据筛选条件文件筛选数据，支持分块处理。
    """
    def __init__(self, file_processor):
        self.file_processor = file_processor  # 复用通用文件处理
    
    def filter_data(self, data_file, filter_file, filter_col, output_path,
                   encoding='auto', delimiter='auto', use_chunking=False,
                   progress_callback=None):
        """
        根据筛选条件文件筛选数据。
        data_file: 源数据CSV文件
        filter_file: 筛选条件Excel文件（第一列是筛选值）
        filter_col: 要筛选的列名
        output_path: 筛选结果输出路径
        """
        try:
            file_size = self.file_processor.get_file_size(data_file)
            if progress_callback:
                progress_callback(5, f"开始筛选 ({file_size:.2f} MB)...")
            
            # 读取筛选条件（Excel第一列）
            if progress_callback:
                progress_callback(20, "正在读取筛选条件...")
            
            filter_df = pd.read_excel(filter_file)
            filter_values = filter_df.iloc[:, 0].dropna().astype(str).tolist() # 读取第一列并转换为字符串列表
            
            if not filter_values:
                return False, "未找到有效筛选条件"
            
            if progress_callback:
                progress_callback(30, f"找到 {len(filter_values)} 个筛选条件")
            
            # 执行筛选
            if use_chunking:
                # 分块处理大文件
                filtered_data = self._filter_in_chunks(
                    data_file, filter_col, filter_values, encoding, delimiter,
                    progress_callback
                )
            else:
                # 整体处理小文件
                filtered_data = self._filter_entire_file(
                    data_file, filter_col, filter_values, encoding, delimiter,
                    progress_callback
                )
            
            if filtered_data is None or filtered_data.empty:
                return False, "未找到匹配的数据"
            
            # 保存结果到Excel
            filtered_data.to_excel(output_path, index=False)
            return True, f"筛选完成！找到 {len(filtered_data)} 条匹配记录，保存至: {output_path}"
            
        except Exception as e:
            return False, f"筛选失败: {str(e)}"
    
    def _filter_entire_file(self, data_file, filter_col, filter_values, encoding, delimiter, progress_callback):
        """
        整体处理小文件筛选，一次性读取整个数据文件进行筛选。
        """
        if progress_callback:
            progress_callback(40, "正在读取数据文件...")
        
        # 读取完整文件
        df, _, _ = self.file_processor.read_csv_robust(
            data_file, encoding, delimiter
        )
        if df is None:
            return None
        
        if filter_col not in df.columns:
            raise ValueError(f"筛选列 '{filter_col}' 不存在于数据文件中。")

        if progress_callback:
            progress_callback(70, "正在执行筛选...")
        
        # 转换为字符串后筛选
        df[filter_col] = df[filter_col].astype(str)
        return df[df[filter_col].isin(filter_values)]
    
    def _filter_in_chunks(self, data_file, filter_col, filter_values, encoding, delimiter, progress_callback):
        """
        分块处理大文件筛选，逐步读取数据块并进行筛选，减少内存占用。
        """
        chunk_size = 10000 # 默认分块大小
        chunks, _, _ = self.file_processor.read_csv_robust(
            data_file, encoding, delimiter, chunksize=chunk_size
        )
        if chunks is None:
            return None
        
        filtered_chunks = []
        processed_rows = 0
        matched_rows = 0
        chunk_count = 0
        
        for chunk in chunks:
            # 检查筛选列是否存在
            if filter_col not in chunk.columns:
                raise ValueError(f"筛选列 '{filter_col}' 不存在于当前数据块中。")

            # 转换为字符串后筛选
            chunk[filter_col] = chunk[filter_col].astype(str)
            chunk_filtered = chunk[chunk[filter_col].isin(filter_values)]
            
            if not chunk_filtered.empty:
                filtered_chunks.append(chunk_filtered)
                matched_rows += len(chunk_filtered)
            
            # 更新进度
            processed_rows += len(chunk)
            chunk_count += 1
            progress = min(90, 30 + (chunk_count * 3)) # 进度计算
            if progress_callback:
                progress_callback(progress, f"已处理 {processed_rows} 行，找到 {matched_rows} 条匹配记录...")
        
        # 合并筛选结果
        if progress_callback:
            progress_callback(95, "正在合并筛选结果...")
        
        return pd.concat(filtered_chunks, ignore_index=True) if filtered_chunks else pd.DataFrame() # 如果没有筛选结果，返回空DataFrame