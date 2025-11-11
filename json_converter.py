import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import os
import pandas as pd
from datetime import datetime
import types  # 在文件顶部导入types模块

class JSONConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("JSON转Excel/Markdown工具")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # 设置中文字体支持
        self.style = ttk.Style()
        self.style.configure("TButton", font=('SimHei', 10))
        self.style.configure("TLabel", font=('SimHei', 10))
        
        # 数据存储
        self.json_data = []
        self.file_paths = []
        
        # 创建界面
        self.create_widgets()

    def create_widgets(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame, padding="5")
        button_frame.pack(fill=tk.X, side=tk.TOP)
        
        self.add_file_btn = ttk.Button(button_frame, text="添加JSON文件", command=self.add_json_files)
        self.add_file_btn.pack(side=tk.LEFT, padx=5)
        
        self.add_folder_btn = ttk.Button(button_frame, text="添加文件夹中的JSON", command=self.add_folder_json)
        self.add_folder_btn.pack(side=tk.LEFT, padx=5)
        
        self.remove_btn = ttk.Button(button_frame, text="移除选中", command=self.remove_selected)
        self.remove_btn.pack(side=tk.LEFT, padx=5)
        
        self.clear_btn = ttk.Button(button_frame, text="清空列表", command=self.clear_all)
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        
        self.sort_label = ttk.Label(button_frame, text="排序方式:")
        self.sort_label.pack(side=tk.LEFT, padx=5, pady=5)
        
        self.sort_var = tk.StringVar(value="文件名")
        self.sort_combo = ttk.Combobox(button_frame, textvariable=self.sort_var, values=["文件名", "标题"], state="readonly", width=10)
        self.sort_combo.pack(side=tk.LEFT, padx=5)
        
        self.sort_btn = ttk.Button(button_frame, text="执行排序", command=self.sort_data)
        self.sort_btn.pack(side=tk.LEFT, padx=5)
        
        # 导出按钮
        export_frame = ttk.Frame(main_frame, padding="5")
        export_frame.pack(fill=tk.X, side=tk.TOP)
        
        self.to_excel_btn = ttk.Button(export_frame, text="导出为Excel", command=self.export_to_excel)
        self.to_excel_btn.pack(side=tk.LEFT, padx=5)
        
        self.to_markdown_btn = ttk.Button(export_frame, text="导出为Markdown", command=self.export_to_markdown)
        self.to_markdown_btn.pack(side=tk.LEFT, padx=5)
        
        # 状态标签
        self.status_var = tk.StringVar(value="就绪")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_label.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 创建表格视图
        columns = ("文件路径", "标题", "状态")
        self.tree = ttk.Treeview(main_frame, columns=columns, show="headings")
        
        # 设置列宽
        self.tree.column("文件路径", width=400)
        self.tree.column("标题", width=250)
        self.tree.column("状态", width=100)
        
        # 设置列标题
        for col in columns:
            self.tree.heading(col, text=col)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        
        # 放置表格和滚动条
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True, side=tk.TOP)

    def add_json_files(self):
        """添加JSON文件"""
        file_types = [("JSON文件", "*.json"), ("所有文件", "*.*")]
        file_paths = filedialog.askopenfilenames(title="选择JSON文件", filetypes=file_types)
        
        if file_paths:
            self.process_json_files(file_paths)

    def add_folder_json(self):
        """添加文件夹中的所有JSON文件"""
        folder_path = filedialog.askdirectory(title="选择包含JSON文件的文件夹")
        
        if folder_path:
            file_paths = []
            for root, _, files in os.walk(folder_path):
                for file in files:
                    if file.lower().endswith('.json'):
                        file_paths.append(os.path.join(root, file))
            
            if file_paths:
                self.process_json_files(file_paths)
            else:
                messagebox.showinfo("提示", f"文件夹 '{folder_path}' 中未找到JSON文件")

    def process_json_files(self, file_paths):
        """处理JSON文件并添加到列表中"""
        success_count = 0
        error_count = 0
        import types
        
        for file_path in file_paths:
            if file_path in self.file_paths:
                continue  # 跳过已添加的文件
            
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                # 验证JSON结构
                if "基础信息" not in data or "标题" not in data["基础信息"]:
                    raise ValueError("JSON结构不符合要求，缺少必要字段")
                
                # 提取必要信息
                title = data["基础信息"]["标题"]
                
                # 检查其他必要字段
                if "极简版" not in data or "简化版" not in data or "详细版" not in data:
                    raise ValueError("JSON结构不符合要求，缺少必要的事件列表字段")
                
                # 源头上防御：确保所有列表字段都不是生成器
                print(f"处理文件 {file_path}，检查字段类型")
                极简版_data = data["极简版"]
                print(f"  极简版原始类型: {type(极简版_data)}")
                # 检查并转换生成器
                if isinstance(极简版_data, types.GeneratorType) or (hasattr(极简版_data, '__iter__') and not hasattr(极简版_data, '__len__') and not isinstance(极简版_data, (str, bytes, dict))):
                    print(f"  极简版是生成器或类似对象，转换为列表")
                    极简版_data = list(极简版_data)
                elif not isinstance(极简版_data, list):
                    极简版_data = [极简版_data] if 极简版_data is not None else []
                    print(f"  极简版非列表，转换为列表")
                    
                简化版_data = data["简化版"]
                print(f"  简化版原始类型: {type(简化版_data)}")
                if isinstance(简化版_data, types.GeneratorType) or (hasattr(简化版_data, '__iter__') and not hasattr(简化版_data, '__len__') and not isinstance(简化版_data, (str, bytes, dict))):
                    print(f"  简化版是生成器或类似对象，转换为列表")
                    简化版_data = list(简化版_data)
                elif not isinstance(简化版_data, list):
                    简化版_data = [简化版_data] if 简化版_data is not None else []
                    print(f"  简化版非列表，转换为列表")
                    
                详细版_data = data["详细版"]
                print(f"  详细版原始类型: {type(详细版_data)}")
                if isinstance(详细版_data, types.GeneratorType) or (hasattr(详细版_data, '__iter__') and not hasattr(详细版_data, '__len__') and not isinstance(详细版_data, (str, bytes, dict))):
                    print(f"  详细版是生成器或类似对象，转换为列表")
                    详细版_data = list(详细版_data)
                elif not isinstance(详细版_data, list):
                    详细版_data = [详细版_data] if 详细版_data is not None else []
                    print(f"  详细版非列表，转换为列表")
                    
                # 添加到数据列表
                self.json_data.append({
                    "file_path": file_path,
                    "title": title,
                    "极简版": 极简版_data,
                    "简化版": 简化版_data,
                    "详细版": 详细版_data
                })
                
                self.file_paths.append(file_path)
                self.tree.insert("", tk.END, values=(os.path.basename(file_path), title, "成功"))
                success_count += 1
                
            except json.JSONDecodeError:
                self.tree.insert("", tk.END, values=(os.path.basename(file_path), "", "JSON格式错误"))
                error_count += 1
            except Exception as e:
                print(f"处理文件 {file_path} 时出错: {str(e)}")
                import traceback
                traceback.print_exc()
                self.tree.insert("", tk.END, values=(os.path.basename(file_path), "", f"错误: {str(e)}"))
                error_count += 1
        
        self.status_var.set(f"添加完成 - 成功: {success_count}, 失败: {error_count}")

    def remove_selected(self):
        """移除选中的项目"""
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showinfo("提示", "请先选择要移除的项目")
            return
        
        removed_count = 0
        for item in selected_items:
            file_name = self.tree.item(item, "values")[0]
            # 查找并移除对应的文件路径和数据
            for i, data in enumerate(self.json_data):
                if os.path.basename(data["file_path"]) == file_name:
                    del self.json_data[i]
                    self.file_paths.remove(data["file_path"])
                    removed_count += 1
                    break
            # 从树视图中删除
            self.tree.delete(item)
        
        self.status_var.set(f"已移除 {removed_count} 个项目")

    def clear_all(self):
        """清空所有数据"""
        if messagebox.askyesno("确认", "确定要清空所有数据吗？"):
            self.tree.delete(*self.tree.get_children())
            self.json_data = []
            self.file_paths = []
            self.status_var.set("已清空所有数据")

    def sort_data(self):
        """根据选择的方式排序数据"""
        if not self.json_data:
            messagebox.showinfo("提示", "没有数据可排序")
            return
        
        sort_by = self.sort_var.get()
        if sort_by == "文件名":
            self.json_data.sort(key=lambda x: os.path.basename(x["file_path"]))
        elif sort_by == "标题":
            self.json_data.sort(key=lambda x: x["标题"])
        
        # 更新树视图
        self.tree.delete(*self.tree.get_children())
        for data in self.json_data:
            self.tree.insert("", tk.END, values=(os.path.basename(data["file_path"]), data["标题"], "成功"))
        
        self.status_var.set(f"已按 {sort_by} 排序")

    def export_to_excel(self):
        """导出为Excel文件（按用户要求格式）"""
        print("======= 导出Excel开始 =======")
        print(f"self.json_data类型: {type(self.json_data)}")
        
        # 全局异常捕获
        try:
            # 1. 防御性编程：确保json_data是列表类型，不是生成器
            if isinstance(self.json_data, types.GeneratorType):
                print("警告：self.json_data是生成器对象，正在转换为列表")
                self.json_data = list(self.json_data)
            elif hasattr(self.json_data, '__iter__') and not hasattr(self.json_data, '__len__') and not isinstance(self.json_data, (str, bytes, dict)):
                print("警告：self.json_data可能是生成器或类似对象，正在转换为列表")
                self.json_data = list(self.json_data)
            # 确保最终是列表类型
            if not isinstance(self.json_data, list):
                self.json_data = [self.json_data] if self.json_data is not None else []
            
            if not self.json_data:
                messagebox.showinfo("提示", "没有数据可导出")
                return
            
            # 创建数据框，严格按照要求的列顺序
            data_to_export = []
            print(f"开始遍历json_data，长度: {len(self.json_data)}")
            
            # 2. 安全迭代json_data
            for index, data in enumerate(self.json_data):
                print(f"处理数据项#{index+1}: {data.get('title', '无标题')}")
                print(f"  数据项类型: {type(data)}")
                
                # 确保data是字典类型
                if not isinstance(data, dict):
                    print(f"  警告：数据项不是字典类型，跳过")
                    continue
                
                # 3. 安全获取并转换各字段
                # 极简版处理
                极简版_value = data.get("极简版", [])
                print(f"  极简版_value原始类型: {type(极简版_value)}")
                # 强制转换为列表，无论是什么类型
                try:
                    # 首先检查是否是生成器
                    if isinstance(极简版_value, types.GeneratorType):
                        print(f"  检测到生成器对象，立即转换为列表")
                        极简版_list = list(极简版_value)
                    # 检查是否是可迭代对象但不是字符串/字典
                    elif hasattr(极简版_value, '__iter__') and not isinstance(极简版_value, (str, bytes, dict)):
                        print(f"  是可迭代对象，准备转换为列表")
                        极简版_list = list(极简版_value)
                    else:
                        print(f"  非可迭代对象或字符串/字典，直接处理")
                        极简版_list = [极简版_value] if 极简版_value is not None else []
                except Exception as e:
                    print(f"  转换极简版失败: {str(e)}")
                    极简版_list = []
                
                # 简化版处理
                简化版_value = data.get("简化版", [])
                print(f"  简化版_value原始类型: {type(简化版_value)}")
                try:
                    if isinstance(简化版_value, types.GeneratorType):
                        print(f"  检测到生成器对象，立即转换为列表")
                        简化版_list = list(简化版_value)
                    elif hasattr(简化版_value, '__iter__') and not isinstance(简化版_value, (str, bytes, dict)):
                        print(f"  是可迭代对象，准备转换为列表")
                        简化版_list = list(简化版_value)
                    else:
                        print(f"  非可迭代对象或字符串/字典，直接处理")
                        简化版_list = [简化版_value] if 简化版_value is not None else []
                except Exception as e:
                    print(f"  转换简化版失败: {str(e)}")
                    简化版_list = []
                
                # 详细版处理
                详细版_value = data.get("详细版", [])
                print(f"  详细版_value原始类型: {type(详细版_value)}")
                try:
                    if isinstance(详细版_value, types.GeneratorType):
                        print(f"  检测到生成器对象，立即转换为列表")
                        详细版_list = list(详细版_value)
                    elif hasattr(详细版_value, '__iter__') and not isinstance(详细版_value, (str, bytes, dict)):
                        print(f"  是可迭代对象，准备转换为列表")
                        详细版_list = list(详细版_value)
                    else:
                        print(f"  非可迭代对象或字符串/字典，直接处理")
                        详细版_list = [详细版_value] if 详细版_value is not None else []
                except Exception as e:
                    print(f"  转换详细版失败: {str(e)}")
                    详细版_list = []
                
                # 4. 生成字符串内容 - 使用完全安全的方法
                # 极简版字符串生成
                极简版_str = ""
                try:
                    items = []
                    # 使用try-except包装整个迭代过程
                    try:
                        # 再次确保是列表
                        if isinstance(极简版_list, list):
                            safe_list = 极简版_list
                        else:
                            safe_list = []
                        
                        # 使用简单循环，避免列表推导式
                        for i, event in enumerate(safe_list, 1):
                            try:
                                items.append(f"{i}. {str(event)}")
                            except Exception as item_e:
                                print(f"  处理列表项{i}时出错: {str(item_e)}")
                                items.append(f"{i}. [数据错误]")
                    except Exception as iter_e:
                        print(f"  迭代极简版列表时出错: {str(iter_e)}")
                    
                    极简版_str = "\n".join(items)
                except Exception as e:
                    print(f"  生成极简版_str出错: {str(e)}")
                
                # 简化版字符串生成
                简化版_str = ""
                try:
                    items = []
                    try:
                        if isinstance(简化版_list, list):
                            safe_list = 简化版_list
                        else:
                            safe_list = []
                        
                        for i, event in enumerate(safe_list, 1):
                            try:
                                items.append(f"{i}. {str(event)}")
                            except Exception as item_e:
                                print(f"  处理列表项{i}时出错: {str(item_e)}")
                                items.append(f"{i}. [数据错误]")
                    except Exception as iter_e:
                        print(f"  迭代简化版列表时出错: {str(iter_e)}")
                    
                    简化版_str = "\n".join(items)
                except Exception as e:
                    print(f"  生成简化版_str出错: {str(e)}")
                
                # 详细版字符串生成
                详细版_str = ""
                try:
                    items = []
                    try:
                        if isinstance(详细版_list, list):
                            safe_list = 详细版_list
                        else:
                            safe_list = []
                        
                        for i, event in enumerate(safe_list, 1):
                            try:
                                items.append(f"{i}. {str(event)}")
                            except Exception as item_e:
                                print(f"  处理列表项{i}时出错: {str(item_e)}")
                                items.append(f"{i}. [数据错误]")
                    except Exception as iter_e:
                        print(f"  迭代详细版列表时出错: {str(iter_e)}")
                    
                    详细版_str = "\n".join(items)
                except Exception as e:
                    print(f"  生成详细版_str出错: {str(e)}")
                
                # 添加到导出数据列表
                data_to_export.append([
                    data.get("title", ""),
                    极简版_str,
                    简化版_str,
                    详细版_str
                ])
            
            # 5. 创建DataFrame并导出
            try:
                df = pd.DataFrame(data_to_export)
                
                # 保存文件
                default_filename = f"小说事件汇总_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel文件", "*.xlsx")],
                    initialfile=default_filename
                )
                
                if file_path:
                    # 使用openpyxl引擎以支持更好的格式控制
                    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                        # 不包含标题行和索引
                        df.to_excel(writer, index=False, header=False)
                        
                        # 获取工作簿和工作表对象以进一步优化格式
                        worksheet = writer.sheets['Sheet1']
                        
                        # 设置字号为12
                        from openpyxl.styles import Font
                        font = Font(size=12)
                        
                        # 防御性编程：确保worksheet.rows不是生成器
                        rows_list = list(worksheet.rows)  # 先转换为列表
                        
                        # 存储每行的最大行数，用于后续调整行高
                        max_lines_per_row = [0] * (len(rows_list) + 1)  # 行号从1开始
                        
                        # 计算每列的最大宽度
                        max_widths = [0] * len(rows_list[0]) if rows_list else []
                        
                        # 设置字体并收集每行的最大行数和每列的最大宽度
                        for row in worksheet.iter_rows():
                            for i, cell in enumerate(row):
                                if cell.value is not None:
                                    # 设置字体
                                    cell.font = font
                                    
                                    try:
                                        # 计算单元格内容的行数
                                        lines = str(cell.value).split('\n')
                                        num_lines = len(lines)
                                        
                                        # 更新该行的最大行数
                                        if num_lines > max_lines_per_row[cell.row]:
                                            max_lines_per_row[cell.row] = num_lines
                                        
                                        # 计算该单元格中最长一行的长度
                                        line_lengths = []
                                        for line in lines:
                                            try:
                                                line_lengths.append(len(line))
                                            except Exception:
                                                line_lengths.append(0)
                                        
                                        max_line_length = max(line_lengths) if line_lengths else 0
                                        # 更新该列的最大宽度
                                        if max_line_length > max_widths[i]:
                                            max_widths[i] = max_line_length
                                    except Exception:
                                        # 如果出现错误，至少设置为1行
                                        if max_lines_per_row[cell.row] < 1:
                                            max_lines_per_row[cell.row] = 1
                                        # 设置默认宽度
                                        if max_widths[i] < 10:
                                            max_widths[i] = 10
                        
                        # 设置每列的宽度
                        for i, max_width in enumerate(max_widths):
                            column_letter = chr(65 + i)  # A, B, C, ...
                            # 设置列宽，添加一些余量
                            adjusted_width = min(max_width * 1.1 + 2, 200)  # 调整系数使宽度更合适
                            worksheet.column_dimensions[column_letter].width = adjusted_width
                        
                        # 设置每行的行高，根据该行内容的最大行数
                        default_row_height = 15  # 12号字体的默认行高
                        line_height = default_row_height  # 每行文本的高度
                        
                        for row_idx in range(1, len(max_lines_per_row)):
                            if max_lines_per_row[row_idx] > 0:
                                # 根据行数调整行高
                                row_height = line_height * max_lines_per_row[row_idx]
                                worksheet.row_dimensions[row_idx].height = row_height
                    
                    self.status_var.set(f"成功导出到: {file_path} (已优化格式)")
                    messagebox.showinfo("成功", f"已成功导出到: {file_path}\n\n已按照要求格式导出：\n- 列顺序：标题、极简版、简化版、详细版\n- 无标题行\n- 每个JSON生成一行\n- 事件元素自动添加序号（如：1. 2. ）\n- 元素之间添加换行符")
                else:
                    # 用户取消了保存
                    pass
            except Exception as e:
                print(f"创建DataFrame或导出时出错: {str(e)}")
                import traceback
                traceback.print_exc()
                messagebox.showerror("错误", f"导出失败: {str(e)}")
                self.status_var.set(f"导出失败: {str(e)}")
        except Exception as e:
            print(f"致命错误 - export_to_excel函数顶层异常: {str(e)}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("错误", f"导出失败: {str(e)}")
            self.status_var.set(f"导出失败: {str(e)}")

    def export_to_markdown(self):
        """导出为Markdown文件"""
        print("======= 导出Markdown开始 =======")
        print(f"self.json_data类型: {type(self.json_data)}")
        
        # 全局异常捕获
        try:
            if not self.json_data:
                messagebox.showinfo("提示", "没有数据可导出")
                return
            
            # 防御性编程：确保json_data不是生成器
            import types
            if isinstance(self.json_data, types.GeneratorType):
                print("警告：self.json_data是生成器对象，正在转换为列表")
                self.json_data = list(self.json_data)
            elif hasattr(self.json_data, '__iter__') and not hasattr(self.json_data, '__len__') and not isinstance(self.json_data, (str, bytes, dict)):
                print("警告：self.json_data可能是生成器或类似对象，正在转换为列表")
                self.json_data = list(self.json_data)
            
            # 生成Markdown内容
            md_content = "# 小说事件汇总\n\n"
            
            for i, data in enumerate(self.json_data, 1):
                print(f"处理Markdown数据项#{i}: {data.get('title', '无标题')}")
                
                # 安全获取标题
                title = data.get('title', '未知标题')
                md_content += f"## {i}. {title}\n\n"
                
                # 添加详细版 - 包含防御性检查
                md_content += "### 详细版\n\n"
                详细版数据 = data.get('详细版', [])
                print(f"  详细版数据类型: {type(详细版数据)}")
                
                # 防御性检查：确保不是生成器
                if isinstance(详细版数据, types.GeneratorType):
                    print(f"  详细版数据是生成器，转换为列表")
                    详细版列表 = list(详细版数据)
                elif hasattr(详细版数据, '__iter__') and not isinstance(详细版数据, (str, bytes, dict)):
                    print(f"  详细版数据是可迭代对象，转换为列表")
                    详细版列表 = list(详细版数据)
                else:
                    详细版列表 = [详细版数据] if 详细版数据 is not None else []
                
                for j, event in enumerate(详细版列表, 1):
                    try:
                        md_content += f"{event}\n\n"
                    except Exception as e:
                        print(f"  添加详细版事件{j}时出错: {str(e)}")
                        md_content += f"[数据错误]\n\n"
                
                # 添加简化版 - 包含防御性检查
                md_content += "### 简化版\n\n"
                简化版数据 = data.get('简化版', [])
                print(f"  简化版数据类型: {type(简化版数据)}")
                
                # 防御性检查：确保不是生成器
                if isinstance(简化版数据, types.GeneratorType):
                    print(f"  简化版数据是生成器，转换为列表")
                    简化版列表 = list(简化版数据)
                elif hasattr(简化版数据, '__iter__') and not isinstance(简化版数据, (str, bytes, dict)):
                    print(f"  简化版数据是可迭代对象，转换为列表")
                    简化版列表 = list(简化版数据)
                else:
                    简化版列表 = [简化版数据] if 简化版数据 is not None else []
                
                for event in 简化版列表:
                    try:
                        md_content += f"- {event}\n"
                    except Exception as e:
                        print(f"  添加简化版事件时出错: {str(e)}")
                        md_content += f"- [数据错误]\n"
                md_content += "\n"
                
                # 添加极简版 - 包含防御性检查
                md_content += "### 极简版\n\n"
                极简版数据 = data.get('极简版', [])
                print(f"  极简版数据类型: {type(极简版数据)}")
                
                # 防御性检查：确保不是生成器
                if isinstance(极简版数据, types.GeneratorType):
                    print(f"  极简版数据是生成器，转换为列表")
                    极简版列表 = list(极简版数据)
                elif hasattr(极简版数据, '__iter__') and not isinstance(极简版数据, (str, bytes, dict)):
                    print(f"  极简版数据是可迭代对象，转换为列表")
                    极简版列表 = list(极简版数据)
                else:
                    极简版列表 = [极简版数据] if 极简版数据 is not None else []
                
                for event in 极简版列表:
                    try:
                        md_content += f"- {event}\n"
                    except Exception as e:
                        print(f"  添加极简版事件时出错: {str(e)}")
                        md_content += f"- [数据错误]\n"
                md_content += "\n"
            
            # 保存文件
            default_filename = f"小说事件汇总_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md"
            file_path = filedialog.asksaveasfilename(
                defaultextension=".md",
                filetypes=[("Markdown文件", "*.md")],
                initialfile=default_filename
            )
            
            if file_path:
                try:
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(md_content)
                    self.status_var.set(f"成功导出到: {file_path}")
                    messagebox.showinfo("成功", f"已成功导出到: {file_path}")
                except Exception as e:
                    messagebox.showerror("错误", f"导出失败: {str(e)}")
                    self.status_var.set(f"导出失败: {str(e)}")
        except Exception as e:
            print(f"致命错误 - export_to_markdown函数顶层异常: {str(e)}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("错误", f"导出失败: {str(e)}")
            self.status_var.set(f"导出失败: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = JSONConverterApp(root)
    root.mainloop()