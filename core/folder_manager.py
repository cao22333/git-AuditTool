import os
import shutil
from datetime import datetime
import json

class FolderManager:
    """
    智能文件夹状态管理的核心逻辑类。
    负责分析文件夹结构、检测空子文件夹、生成重命名预览以及执行重命名操作。
    不包含任何UI逻辑，方便集成和测试。
    """
    def __init__(self, logger=None):
        self.logger = logger if logger else self._default_logger # 注入日志器
        self.rename_history = [] # 用于存储重命名操作，以便撤销
        
        # 默认后缀配置
        self.empty_suffix = "-空"
        self.parent_partial_suffix = "-缺"
        self.parent_all_empty_suffix = "-全空"

    def _default_logger(self, message, level="INFO"):
        """默认日志器，如果未提供外部日志器则打印到控制台"""
        print(f"[{level}] {message}")

    def set_suffixes(self, empty_suffix, parent_partial_suffix, parent_all_empty_suffix):
        """
        设置文件夹后缀。
        """
        self.empty_suffix = empty_suffix
        self.parent_partial_suffix = parent_partial_suffix
        self.parent_all_empty_suffix = parent_all_empty_suffix

    def has_suffix(self, folder_name):
        """
        检查文件夹名是否包含当前配置的任何后缀。
        返回 (匹配到的后缀, 不带后缀的基础名称) 或 (None, 原始名称)。
        """
        suffixes = [
            self.empty_suffix,
            self.parent_partial_suffix,
            self.parent_all_empty_suffix
        ]
        
        for suffix in suffixes:
            if folder_name.endswith(suffix):
                return suffix, folder_name[:-len(suffix)]
        
        return None, folder_name

    def analyze_folder_structure(self, base_dir, remove_correct_suffix=True):
        """
        分析指定目录下的文件夹结构，生成状态报告和重命名建议。
        base_dir: 要分析的根目录。
        remove_correct_suffix: 是否移除正确状态下的后缀（如非空文件夹的"-空"后缀）。
        """
        if not os.path.isdir(base_dir):
            self.logger(f"错误: 指定的目录不存在或不是一个文件夹: {base_dir}", "ERROR")
            return None
        
        stats = {
            'total_parents': 0,
            'total_subfolders': 0,
            'empty_subfolders': 0,
            'status_counts': {
                '需要添加后缀': 0,
                '需要移除后缀': 0,
                '需要更正后缀': 0, # 新增更正后缀类型
                '无需更改': 0
            },
            'parent_status': {
                '全部为空': 0,
                '部分为空': 0,
                '全部非空': 0,
                '无子文件夹': 0
            },
            'folder_details': [] # 存储所有文件夹的详细信息，包括重命名建议
        }
        
        self.logger(f"开始分析目录: {base_dir}", "INFO")

        try:
            # 获取所有直接子目录作为父文件夹
            parent_items = [item for item in os.listdir(base_dir) if os.path.isdir(os.path.join(base_dir, item))]
            stats['total_parents'] = len(parent_items)
            
            for parent_name in parent_items:
                parent_path = os.path.join(base_dir, parent_name)
                
                # 检查父文件夹是否有后缀
                parent_current_suffix, parent_base_name = self.has_suffix(parent_name)
                
                # 获取父文件夹下的所有直接子目录
                subfolders = [item for item in os.listdir(parent_path) if os.path.isdir(os.path.join(parent_path, item))]
                stats['total_subfolders'] += len(subfolders)
                
                # 分析子文件夹状态
                subfolder_details = []
                empty_count = 0
                
                for sub_name in subfolders:
                    sub_path = os.path.join(parent_path, sub_name)
                    is_empty = not bool(os.listdir(sub_path)) # 判断子文件夹是否为空
                    
                    # 检查子文件夹后缀
                    sub_current_suffix, sub_base_name = self.has_suffix(sub_name)
                    
                    new_sub_name = sub_name
                    sub_operation = "无需更改"
                    sub_need_rename = False

                    if is_empty:
                        stats['empty_subfolders'] += 1
                        empty_count += 1
                        if sub_current_suffix != self.empty_suffix:
                            # 空文件夹但没有正确后缀或后缀不匹配
                            new_sub_name = f"{sub_base_name}{self.empty_suffix}"
                            sub_operation = "需要添加后缀" if not sub_current_suffix else "需要更正后缀"
                            sub_need_rename = True
                            stats['status_counts']['需要添加后缀'] += 1 if not sub_current_suffix else 0
                            stats['status_counts']['需要更正后缀'] += 1 if sub_current_suffix else 0
                        else:
                            sub_operation = "后缀正确" # 空文件夹且后缀正确
                    else:
                        # 非空文件夹
                        if sub_current_suffix == self.empty_suffix:
                            # 非空文件夹但有"空"后缀，需要移除
                            new_sub_name = sub_base_name
                            sub_operation = "需要移除后缀"
                            sub_need_rename = True
                            stats['status_counts']['需要移除后缀'] += 1
                        else:
                            sub_operation = "无需更改" # 非空文件夹且没有"空"后缀

                    subfolder_details.append({
                        'name': sub_name,
                        'base_name': sub_base_name,
                        'path': sub_path,
                        'current_suffix': sub_current_suffix,
                        'is_empty': is_empty,
                        'new_name': new_sub_name,
                        'operation': sub_operation,
                        'need_rename': sub_need_rename
                    })
                
                # 判断父文件夹状态
                parent_status_text = ""
                parent_expected_suffix = ""
                if len(subfolders) == 0:
                    parent_status_text = "无子文件夹"
                    stats['parent_status']['无子文件夹'] += 1
                elif empty_count == len(subfolders):
                    parent_status_text = "全部为空"
                    stats['parent_status']['全部为空'] += 1
                    parent_expected_suffix = self.parent_all_empty_suffix
                elif empty_count > 0:
                    parent_status_text = "部分为空"
                    stats['parent_status']['部分为空'] += 1
                    parent_expected_suffix = self.parent_partial_suffix
                else:
                    parent_status_text = "全部非空"
                    stats['parent_status']['全部非空'] += 1
                    parent_expected_suffix = "" # 全部非空不应该有后缀
                
                new_parent_name = parent_name
                parent_operation = "无需更改"
                parent_need_rename = False

                # 根据父文件夹的实际状态和期望后缀，判断是否需要重命名
                if parent_expected_suffix: # 期望有后缀 (全部为空或部分为空)
                    if parent_current_suffix != parent_expected_suffix:
                        new_parent_name = f"{parent_base_name}{parent_expected_suffix}"
                        parent_operation = "需要添加后缀" if not parent_current_suffix else "需要更正后缀"
                        parent_need_rename = True
                        stats['status_counts']['需要添加后缀'] += 1 if not parent_current_suffix else 0
                        stats['status_counts']['需要更正后缀'] += 1 if parent_current_suffix else 0
                else: # 期望没有后缀 (全部非空或无子文件夹)
                    if parent_current_suffix: # 实际有后缀
                        if remove_correct_suffix: # 配置为移除正确状态的后缀
                            new_parent_name = parent_base_name
                            parent_operation = "需要移除后缀"
                            parent_need_rename = True
                            stats['status_counts']['需要移除后缀'] += 1
                        else:
                            parent_operation = "有额外后缀 (未移除)" # 不移除，则标记为有额外后缀
                    else:
                        parent_operation = "无需更改" # 实际没有后缀，符合期望
                
                stats['folder_details'].append({
                    'type': 'parent',
                    'name': parent_name,
                    'base_name': parent_base_name,
                    'path': parent_path,
                    'current_suffix': parent_current_suffix,
                    'subfolders_count': len(subfolders),
                    'empty_subfolders_count': empty_count,
                    'status_text': parent_status_text,
                    'expected_suffix': parent_expected_suffix,
                    'new_name': new_parent_name,
                    'operation': parent_operation,
                    'need_rename': parent_need_rename,
                    'subfolder_details': subfolder_details
                })
                
        except Exception as e:
            self.logger(f"分析文件夹结构时发生错误: {str(e)}", "ERROR")
            return None
        
        self.logger("文件夹结构分析完成。", "INFO")
        return stats

    def get_rename_preview(self, base_dir, remove_correct_suffix=True):
        """
        获取将要执行的重命名操作的预览列表。
        """
        stats = self.analyze_folder_structure(base_dir, remove_correct_suffix)
        if not stats:
            return []

        preview_list = []
        for folder_info in stats['folder_details']:
            # 处理子文件夹的重命名
            for sub_info in folder_info['subfolder_details']:
                if sub_info['need_rename']:
                    preview_list.append({
                        'type': '子文件夹',
                        'old_path': sub_info['path'],
                        'new_path': os.path.join(os.path.dirname(sub_info['path']), sub_info['new_name']),
                        'old_name': sub_info['name'],
                        'new_name': sub_info['new_name'],
                        'current_status': '空' if sub_info['is_empty'] else '非空',
                        'operation': sub_info['operation'],
                        'parent_path': folder_info['path'] # 方便UI显示父级
                    })
            
            # 处理父文件夹的重命名
            if folder_info['need_rename']:
                preview_list.append({
                    'type': '父文件夹',
                    'old_path': folder_info['path'],
                    'new_path': os.path.join(base_dir, folder_info['new_name']),
                    'old_name': folder_info['name'],
                    'new_name': folder_info['new_name'],
                    'current_status': folder_info['status_text'],
                    'operation': folder_info['operation'],
                    'parent_path': base_dir # 根目录
                })
        
        return preview_list

    def execute_renames(self, rename_operations, base_dir, create_backup=False, progress_callback=None):
        """
        执行重命名操作。
        rename_operations: 包含 {old_path, new_path, ...} 的字典列表。
        base_dir: 根目录，用于备份。
        create_backup: 是否在执行前创建备份。
        progress_callback: 进度更新回调函数。
        """
        if not rename_operations:
            self.logger("没有需要执行的重命名操作。", "INFO")
            return 0, 0

        success_count = 0
        failed_count = 0
        current_rename_history = [] # 本次操作的历史记录

        # 创建备份
        if create_backup and base_dir:
            backup_path = os.path.join(base_dir, f"backup_folder_manager_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
            try:
                # 忽略备份文件夹本身以及日志文件
                shutil.copytree(base_dir, backup_path, ignore=shutil.ignore_patterns('backup_folder_manager_*', '*.log'))
                self.logger(f"已创建备份到: {backup_path}", "SUCCESS")
            except Exception as e:
                self.logger(f"备份失败: {str(e)}", "ERROR")
        
        total_operations = len(rename_operations)

        # 按照子文件夹先于父文件夹的顺序进行处理
        # 否则如果父文件夹先改名，子文件夹的旧路径就会失效
        sorted_operations = sorted(rename_operations, key=lambda x: x['type'] == 'parent') # '子文件夹'在前，'父文件夹'在后

        for i, item in enumerate(sorted_operations):
            old_path = item['old_path']
            new_path = item['new_path']
            operation_type = item['operation']

            try:
                if os.path.exists(old_path):
                    os.rename(old_path, new_path)
                    self.logger(f"重命名成功: '{os.path.basename(old_path)}' -> '{os.path.basename(new_path)}' ({operation_type})", "SUCCESS")
                    current_rename_history.append((new_path, old_path)) # 存储新旧路径，方便撤销
                    success_count += 1
                else:
                    self.logger(f"警告: 原始文件夹不存在，跳过重命名: {old_path}", "WARNING")
                    failed_count += 1
            except Exception as e:
                self.logger(f"重命名失败: '{old_path}' -> '{new_path}'。错误: {str(e)}", "ERROR")
                failed_count += 1
            
            if progress_callback:
                progress_callback(int((i + 1) / total_operations * 100), f"正在执行重命名 ({i+1}/{total_operations})...")

        if current_rename_history:
            self.rename_history.extend(current_rename_history) # 将本次操作的历史记录添加到总历史记录中
        
        self.logger(f"重命名操作完成。成功: {success_count}, 失败: {failed_count}", "INFO")
        return success_count, failed_count

    def undo_last_operations(self, progress_callback=None):
        """
        撤销上一次执行的重命名操作。
        """
        if not self.rename_history:
            self.logger("没有可撤销的操作。", "INFO")
            return 0, 0
        
        self.logger("开始撤销最近的重命名操作...", "INFO")
        undo_count = 0
        # 撤销是 FILO (First-In, Last-Out) 的，所以从历史记录末尾开始
        # 但实际重命名时，我们是先处理子文件夹，再处理父文件夹
        # 撤销时，需要先撤销父文件夹的重命名，再撤销子文件夹的重命名
        # 因此，需要反转 history 列表，或者在存储时就按父文件夹->子文件夹的顺序存储。
        # 这里假设 rename_history 存储的是 (new_path, old_path) 格式，且是按执行顺序追加的
        # 所以撤销时，需要逆序遍历，并执行 os.rename(new_path, old_path)
        
        # 为了正确撤销，我们应该撤销最近一次批次的所有操作。
        # 这里简化处理，直接撤销所有历史记录。如果需要批次撤销，需要更复杂的历史管理。
        operations_to_undo = self.rename_history[:] # 复制一份，避免在循环中修改原列表
        self.rename_history.clear() # 清空历史，准备记录新的历史

        # 撤销时，需要先将父文件夹恢复，再恢复子文件夹，所以需要对撤销列表进行排序
        # 存储的是 (new_path, old_path)，现在要执行 os.rename(new_path, old_path)
        # 如果 new_path 是父文件夹，它的深度比子文件夹浅，应该后撤销
        # 所以这里需要按路径深度降序排序 (父文件夹路径短，深度浅)
        # 或者更简单，直接逆序遍历原始历史记录，因为它是按子文件夹->父文件夹的顺序添加的
        
        total_undo_operations = len(operations_to_undo)
        for i, (current_new_path, original_old_path) in enumerate(reversed(operations_to_undo)):
            try:
                if os.path.exists(current_new_path):
                    os.rename(current_new_path, original_old_path)
                    self.logger(f"撤销成功: '{os.path.basename(current_new_path)}' -> '{os.path.basename(original_old_path)}'", "SUCCESS")
                    undo_count += 1
                else:
                    self.logger(f"警告: 当前路径不存在，跳过撤销: {current_new_path}", "WARNING")
            except Exception as e:
                self.logger(f"撤销失败: '{current_new_path}' -> '{original_old_path}'。错误: {str(e)}", "ERROR")
                # 如果撤销失败，将该操作重新添加到历史记录中，以便下次尝试或手动处理
                self.rename_history.append((current_new_path, original_old_path)) 
            
            if progress_callback:
                progress_callback(int((i + 1) / total_undo_operations * 100), f"正在撤销操作 ({i+1}/{total_undo_operations})...")

        self.logger(f"撤销操作完成。成功撤销 {undo_count} 个操作。", "INFO")
        return undo_count, total_undo_operations - undo_count # 返回成功撤销数和失败数

    def clear_rename_history(self):
        """清空重命名历史记录。"""
        self.rename_history.clear()
        self.logger("重命名历史记录已清空。", "INFO")

    def export_analysis_report(self, stats, output_path):
        """
        导出文件夹分析报告到指定文件。
        """
        if not stats:
            self.logger("没有可导出的分析报告。", "WARNING")
            return False, "没有可导出的分析报告。"

        try:
            report_content = self._format_analysis_report(stats)
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(report_content)
            self.logger(f"分析报告已导出到: {output_path}", "SUCCESS")
            return True, f"报告已导出至: {output_path}"
        except Exception as e:
            self.logger(f"导出分析报告失败: {str(e)}", "ERROR")
            return False, f"导出报告失败: {str(e)}"

    def _format_analysis_report(self, stats):
        """
        格式化分析报告内容。
        """
        summary = f"""
{'='*50}
                   文件夹状态分析报告
{'='*50}

📁 目录信息:
    • 处理目录: {os.path.dirname(stats['folder_details'][0]['path']) if stats['folder_details'] else 'N/A'}
    • 分析时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

📊 总体统计:
    ├─ 📂 父文件夹总数: {stats['total_parents']}
    ├─ 📁 子文件夹总数: {stats['total_subfolders']}
    └─ 📭 空子文件夹数: {stats['empty_subfolders']}

🔄 状态分布:
    ├─ ✅ 无需更改: {stats['status_counts']['无需更改']}
    ├─ ➕ 需要添加后缀: {stats['status_counts']['需要添加后缀']}
    ├─ ✏️ 需要更正后缀: {stats['status_counts']['需要更正后缀']}
    └─ ➖ 需要移除后缀: {stats['status_counts']['需要移除后缀']}

🏷️ 父文件夹状态分类:
    ├─ ⚪ 全部为空: {stats['parent_status']['全部为空']}
    ├─ 🟡 部分为空: {stats['parent_status']['部分为空']}
    ├─ 🟢 全部非空: {stats['parent_status']['全部非空']}
    └─ ⚫ 无子文件夹: {stats['parent_status']['无子文件夹']}

{'='*50}
                    详细列表
{'='*50}
"""
        details = []
        for folder in stats['folder_details']:
            parent_detail = f"""
📂 父文件夹: {folder['name']} (操作: {folder['operation']})
   ├─ 路径: {folder['path']}
   ├─ 当前后缀: {folder['current_suffix'] if folder['current_suffix'] else '无'}
   ├─ 期望后缀: {folder['expected_suffix'] if folder['expected_suffix'] else '无'}
   ├─ 新名称: {folder['new_name']}
   ├─ 状态: {folder['status_text']}
   ├─ 子文件夹数: {folder['subfolders_count']}
   └─ 空子文件夹数: {folder['empty_subfolders_count']}
   
   📁 子文件夹详情:
"""
            details.append(parent_detail)
            
            if not folder['subfolder_details']:
                details.append("      (无子文件夹)\n")
            else:
                for sub in folder['subfolder_details']:
                    status_icon = "📭" if sub['is_empty'] else "📂"
                    sub_detail = f"      {status_icon} {sub['name']} (操作: {sub['operation']}, 新名称: {sub['new_name']})\n"
                    details.append(sub_detail)
            details.append("\n") # 每个父文件夹后加一个空行

        return summary + "".join(details)