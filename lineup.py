import os
import sys
import shutil
import re
import difflib
import json
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter import scrolledtext
from tkinter import ttk
import openpyxl
import webbrowser

class LineupApp:
    def __init__(self, root):
        self.root = root
        self.root.title("文件排序工具")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        
        # 设置样式
        self.setup_styles()
        
        self.folder_path = ""
        self.list_items = []
        self.similarity_threshold = 0.6  # 初始化相似度阈值
        self.auto_select_highest = tk.BooleanVar(value=False)
        self.generate_list_only = tk.BooleanVar(value=False)
        self.ignore_directories = tk.BooleanVar(value=False)
        self.output_format = tk.StringVar(value="text")
        self.filename_format = tk.StringVar(value="relative")
        self.output_folder = ""
        self.output_file = ""
        
        self.rename_mode = tk.StringVar(value="add_prefix")
        self.separator = tk.StringVar(value="-")
        self.format_str = tk.StringVar(value="[Num]")
        self.start_num = tk.IntVar(value=1)
        self.step = tk.IntVar(value=1)
        self.reverse = tk.BooleanVar(value=False)
        self.end_num = tk.IntVar(value=1)
        
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 主选项卡
        self.main_frame = ttk.Frame(self.notebook, style='Card.TFrame')
        self.notebook.add(self.main_frame, text="主界面")
        self.setup_main_frame()
        
        # 配置选项卡
        self.config_frame = ttk.Frame(self.notebook, style='Card.TFrame')
        self.notebook.add(self.config_frame, text="配置")
        self.setup_config_frame()
        
        # 关于选项卡
        self.about_frame = ttk.Frame(self.notebook, style='Card.TFrame')
        self.notebook.add(self.about_frame, text="关于")
        self.setup_about_frame()
    
    def setup_styles(self):
        style = ttk.Style()
        # 尝试使用现代主题
        try:
            style.theme_use('clam')
        except tk.TclError:
            pass  # 如果不可用，使用默认
        
        # 定义自定义样式
        style.configure('Card.TFrame', background='#f0f0f0', relief='raised', borderwidth=2)
        style.configure('TLabel', font=('Microsoft YaHei', 10), background='#f0f0f0')
        style.configure('TButton', font=('Microsoft YaHei', 10), padding=5)
        style.configure('TRadiobutton', font=('Microsoft YaHei', 10), background='#f0f0f0')
        style.configure('TEntry', font=('Microsoft YaHei', 10))
        style.configure('TText', font=('Microsoft YaHei', 10))
        style.configure('TScale', background='#f0f0f0')
        style.configure('TListbox', font=('Microsoft YaHei', 10))
        
        # 设置根窗口背景
        self.root.configure(bg='#e0e0e0')
    
    def setup_main_frame(self):
        # 文件夹选择组
        folder_group = ttk.LabelFrame(self.main_frame, text="文件夹选择", style='Card.TFrame')
        folder_group.pack(fill="x", padx=10, pady=10)
        
        ttk.Label(folder_group, text="选择文件夹:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.folder_entry = ttk.Entry(folder_group, width=50)
        self.folder_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        ttk.Button(folder_group, text="浏览", command=self.select_folder).grid(row=0, column=2, padx=10, pady=10)
        
        folder_group.columnconfigure(1, weight=1)
        
        # 列表输入组
        list_group = ttk.LabelFrame(self.main_frame, text="目的列表输入", style='Card.TFrame')
        list_group.pack(fill="x", padx=10, pady=10)
        
        self.list_mode = tk.StringVar(value="file")
        ttk.Radiobutton(list_group, text="导入列表文件", variable=self.list_mode, value="file").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        ttk.Radiobutton(list_group, text="导入Excel文件", variable=self.list_mode, value="excel").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        ttk.Radiobutton(list_group, text="手动输入", variable=self.list_mode, value="manual").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        
        self.import_button = ttk.Button(list_group, text="导入", command=self.import_list)
        self.import_button.grid(row=0, column=1, rowspan=3, padx=10, pady=10, sticky="ns")
        
        # 手动输入文本框
        self.manual_text = tk.Text(list_group, height=6, width=60, font=('Microsoft YaHei', 10), wrap=tk.WORD)
        self.manual_text.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
        
        list_group.columnconfigure(0, weight=1)
        
        # 按钮组
        button_group = ttk.Frame(self.main_frame, style='Card.TFrame')
        button_group.pack(fill="x", padx=10, pady=10)
        
        ttk.Button(button_group, text="预览", command=self.preview, style='Accent.TButton').pack(side=tk.LEFT, padx=20, pady=10)
        ttk.Button(button_group, text="运行", command=self.run_lineup, style='Accent.TButton').pack(side=tk.LEFT, padx=20, pady=10)
        
        # 结果显示组
        result_group = ttk.LabelFrame(self.main_frame, text="结果", style='Card.TFrame')
        result_group.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.result_text = tk.Text(result_group, height=12, width=80, font=('Microsoft YaHei', 10), wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(result_group, orient=tk.VERTICAL, command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=scrollbar.set)
        self.result_text.pack(side=tk.LEFT, fill="both", expand=True, padx=10, pady=10)
        scrollbar.pack(side=tk.RIGHT, fill="y")
    
    def setup_config_frame(self):
        # 创建可滚动框架
        canvas = tk.Canvas(self.config_frame)
        scrollbar = ttk.Scrollbar(self.config_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        config_group = ttk.LabelFrame(scrollable_frame, text="匹配配置", style='Card.TFrame')
        config_group.pack(fill="x", padx=20, pady=20)
        
        ttk.Label(config_group, text="相似度阈值 (0.0 - 1.0):").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.threshold_var = tk.DoubleVar(value=0.6)
        self.threshold_scale = ttk.Scale(config_group, from_=0.0, to=1.0, variable=self.threshold_var, orient="horizontal", length=300)
        self.threshold_scale.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        self.threshold_label = ttk.Label(config_group, text="0.60", font=('Microsoft YaHei', 12, 'bold'))
        self.threshold_label.grid(row=0, column=2, padx=10, pady=10)
        self.threshold_var.trace("w", self.update_threshold_label)
        
        ttk.Checkbutton(config_group, text="自动选择最高相似度（当有多个候选时）", variable=self.auto_select_highest).grid(row=1, column=0, columnspan=3, padx=10, pady=10, sticky="w")
        
        ttk.Checkbutton(config_group, text="仅生成列表（不复制文件）", variable=self.generate_list_only).grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky="w")
        
        ttk.Checkbutton(config_group, text="忽略目录（只处理文件）", variable=self.ignore_directories).grid(row=3, column=0, columnspan=3, padx=10, pady=10, sticky="w")
        
        ttk.Label(config_group, text="输出格式:").grid(row=4, column=0, padx=10, pady=10, sticky="w")
        ttk.Radiobutton(config_group, text="文本 (Result.txt)", variable=self.output_format, value="text").grid(row=5, column=0, sticky="w", padx=20, pady=5)
        ttk.Radiobutton(config_group, text="JSON (Result.json)", variable=self.output_format, value="json").grid(row=6, column=0, sticky="w", padx=20, pady=5)
        ttk.Radiobutton(config_group, text="M3U (Result.m3u)", variable=self.output_format, value="m3u").grid(row=7, column=0, sticky="w", padx=20, pady=5)
        
        ttk.Label(config_group, text="文件名格式:").grid(row=8, column=0, padx=10, pady=10, sticky="w")
        ttk.Radiobutton(config_group, text="相对路径", variable=self.filename_format, value="relative").grid(row=9, column=0, sticky="w", padx=20, pady=5)
        ttk.Radiobutton(config_group, text="绝对路径", variable=self.filename_format, value="absolute").grid(row=10, column=0, sticky="w", padx=20, pady=5)
        
        ttk.Label(config_group, text="结果输出文件夹:").grid(row=11, column=0, padx=10, pady=10, sticky="w")
        self.output_folder_entry = ttk.Entry(config_group, width=30)
        self.output_folder_entry.grid(row=11, column=1, padx=10, pady=10, sticky="ew")
        ttk.Button(config_group, text="浏览", command=self.select_output_folder).grid(row=11, column=2, padx=10, pady=10)
        
        ttk.Label(config_group, text="输出列表文件路径:").grid(row=12, column=0, padx=10, pady=10, sticky="w")
        self.output_file_entry = ttk.Entry(config_group, width=30)
        self.output_file_entry.grid(row=12, column=1, padx=10, pady=10, sticky="ew")
        ttk.Button(config_group, text="浏览", command=self.select_output_file).grid(row=12, column=2, padx=10, pady=10)
        
        config_group.columnconfigure(1, weight=1)
        
        # 重命名策略
        rename_group = ttk.LabelFrame(scrollable_frame, text="重命名策略", style='Card.TFrame')
        rename_group.pack(fill="x", padx=20, pady=20)
        
        ttk.Label(rename_group, text="模式:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        ttk.Radiobutton(rename_group, text="在文件名前加序号", variable=self.rename_mode, value="add_prefix").grid(row=1, column=0, sticky="w", padx=20, pady=5)
        ttk.Radiobutton(rename_group, text="按格式重命名", variable=self.rename_mode, value="custom_format").grid(row=2, column=0, sticky="w", padx=20, pady=5)
        
        ttk.Label(rename_group, text="分隔符:").grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.separator_entry = ttk.Entry(rename_group, textvariable=self.separator, width=10)
        self.separator_entry.grid(row=3, column=1, padx=10, pady=10, sticky="w")
        
        ttk.Label(rename_group, text="格式 (使用[Num]作为序号占位符):").grid(row=4, column=0, padx=10, pady=10, sticky="w")
        self.format_entry = ttk.Entry(rename_group, textvariable=self.format_str, width=20)
        self.format_entry.grid(row=4, column=1, padx=10, pady=10, sticky="ew")
        
        ttk.Label(rename_group, text="序号起始数:").grid(row=5, column=0, padx=10, pady=10, sticky="w")
        self.start_num_entry = ttk.Entry(rename_group, textvariable=self.start_num, width=10)
        self.start_num_entry.grid(row=5, column=1, padx=10, pady=10, sticky="w")
        
        ttk.Label(rename_group, text="跨度:").grid(row=6, column=0, padx=10, pady=10, sticky="w")
        self.step_entry = ttk.Entry(rename_group, textvariable=self.step, width=10)
        self.step_entry.grid(row=6, column=1, padx=10, pady=10, sticky="w")
        
        ttk.Checkbutton(rename_group, text="倒序", variable=self.reverse).grid(row=7, column=0, padx=10, pady=10, sticky="w")
        
        ttk.Label(rename_group, text="末尾数 (倒序时):").grid(row=8, column=0, padx=10, pady=10, sticky="w")
        self.end_num_entry = ttk.Entry(rename_group, textvariable=self.end_num, width=10)
        self.end_num_entry.grid(row=8, column=1, padx=10, pady=10, sticky="w")
        
        rename_group.columnconfigure(1, weight=1)
    
    def setup_about_frame(self):
        about_group = ttk.LabelFrame(self.about_frame, text="关于文件排序工具", style='Card.TFrame')
        about_group.pack(fill="both", expand=True, padx=20, pady=20)
        
        about_text = """
文件排序工具 (zh-lineup)

版本: 1.0
作者: GZYZhy

GitHub: https://github.com/GZYZhy/zh-lineup

使用教程:
1. 选择包含文件/目录的文件夹。
2. 选择目的列表输入方式：导入文件、Excel或手动输入。
3. 点击"预览"查看结果，或"运行"执行排序。
4. 在"配置"选项卡中调整匹配参数。

功能:
- 模糊匹配文件和目录
- 可配置相似度阈值
- 完全匹配优先
- 多种输入方式
- 预览和仅生成列表模式
- 跨平台GUI界面

许可证: Apache License 2.0
"""
        text = tk.Text(about_group, wrap=tk.WORD, font=('Microsoft YaHei', 10), height=20)
        scrollbar = ttk.Scrollbar(about_group, orient=tk.VERTICAL, command=text.yview)
        text.configure(yscrollcommand=scrollbar.set)
        text.pack(side=tk.LEFT, fill="both", expand=True, padx=10, pady=10)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        
        text.insert(tk.END, about_text.strip())
        text.config(state=tk.DISABLED)
    
    def update_threshold_label(self, *args):
        self.threshold_label.config(text=f"{self.threshold_var.get():.2f}")
        self.similarity_threshold = self.threshold_var.get()
    
    def generate_new_name(self, num, item):
        if self.rename_mode.get() == "add_prefix":
            separator = self.separator.get()
            return f"{num}{separator}{item}"
        else:
            format_str = self.format_str.get()
            base = format_str.replace("[Num]", str(num))
            # 过滤特殊字符
            base = re.sub(r'[<>:"|?*\\/]', '', base)
            # 保留扩展名
            name, ext = os.path.splitext(item)
            return base + ext
    
    def select_folder(self):
        self.folder_path = filedialog.askdirectory()
        if self.folder_path:
            self.folder_entry.delete(0, tk.END)
            self.folder_entry.insert(0, self.folder_path)
    
    def select_output_folder(self):
        self.output_folder = filedialog.askdirectory()
        if self.output_folder:
            self.output_folder_entry.delete(0, tk.END)
            self.output_folder_entry.insert(0, self.output_folder)
    
    def select_output_file(self):
        output_format = self.output_format.get()
        if output_format == "text":
            default_ext = ".txt"
        elif output_format == "json":
            default_ext = ".json"
        elif output_format == "m3u":
            default_ext = ".m3u"
        else:
            default_ext = ".txt"
        
        self.output_file = filedialog.asksaveasfilename(
            defaultextension=default_ext,
            filetypes=[
                ("文本文件", "*.txt") if output_format == "text" else ("JSON文件", "*.json") if output_format == "json" else ("M3U文件", "*.m3u"),
                ("所有文件", "*.*")
            ]
        )
        if self.output_file:
            self.output_file_entry.delete(0, tk.END)
            self.output_file_entry.insert(0, self.output_file)
    
    def import_list(self):
        mode = self.list_mode.get()
        if mode == "file":
            path = filedialog.askopenfilename(filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")])
            if path:
                with open(path, 'r', encoding='utf-8') as f:
                    self.list_items = [line.strip() for line in f if line.strip()]
                self.manual_text.delete(1.0, tk.END)
                self.manual_text.insert(tk.END, '\n'.join(self.list_items))
        elif mode == "excel":
            path = filedialog.askopenfilename(filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")])
            if path:
                wb = openpyxl.load_workbook(path)
                sheet = wb.active
                self.list_items = [str(cell.value).strip() for cell in sheet['A'] if cell.value is not None]
                self.manual_text.delete(1.0, tk.END)
                self.manual_text.insert(tk.END, '\n'.join(self.list_items))
        # For manual, list_items will be read from text box
    
    def get_list_items(self):
        if self.list_mode.get() == "manual":
            text = self.manual_text.get(1.0, tk.END).strip()
            self.list_items = [line.strip() for line in text.split('\n') if line.strip()]
        return self.list_items
    
    def preview(self):
        folder = self.folder_entry.get()
        if not folder:
            messagebox.showerror("错误", "请选择文件夹")
            return
        
        if not os.path.isdir(folder):
            messagebox.showerror("错误", f"{folder} 不是一个目录")
            return
        
        items = self.get_list_items()
        if not items:
            messagebox.showerror("错误", "请提供目的列表")
            return
        
        # 创建预览弹窗
        preview_dialog = tk.Toplevel(self.root)
        preview_dialog.title("预览结果")
        preview_dialog.geometry("600x500")
        preview_dialog.configure(bg='#f0f0f0')
        
        ttk.Label(preview_dialog, text="预览结果", font=('Microsoft YaHei', 14, 'bold')).pack(pady=10)
        
        text_frame = ttk.Frame(preview_dialog, style='Card.TFrame')
        text_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        preview_text = tk.Text(text_frame, height=20, width=70, font=('Microsoft YaHei', 10), wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=preview_text.yview)
        preview_text.configure(yscrollcommand=scrollbar.set)
        preview_text.pack(side=tk.LEFT, fill="both", expand=True, padx=10, pady=10)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        
        preview_text.insert(tk.END, "正在生成预览...\n")
        preview_dialog.update()
        
        try:
            result = self.process_lineup(folder, items, preview=True)
            preview_text.delete(1.0, tk.END)
            preview_text.insert(tk.END, result)
        except Exception as e:
            preview_text.delete(1.0, tk.END)
            preview_text.insert(tk.END, f"错误: {str(e)}")
        
        ttk.Button(preview_dialog, text="关闭", command=preview_dialog.destroy).pack(pady=10)
    
    def run_lineup(self):
        folder = self.folder_entry.get()
        if not folder:
            messagebox.showerror("错误", "请选择文件夹")
            return
        
        if not os.path.isdir(folder):
            messagebox.showerror("错误", f"{folder} 不是一个目录")
            return
        
        items = self.get_list_items()
        if not items:
            messagebox.showerror("错误", "请提供目的列表")
            return
        
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "正在处理...\n")
        self.root.update()
        
        try:
            result = self.process_lineup(folder, items, preview=False)
            self.result_text.insert(tk.END, result)
        except Exception as e:
            messagebox.showerror("错误", str(e))
    
    def process_lineup(self, folder, lines, preview=False):
        output_format = self.output_format.get()
        filename_format = self.filename_format.get()
        
        if output_format == "text":
            result_file = "Result.txt"
        elif output_format == "json":
            result_file = "Result.json"
        elif output_format == "m3u":
            result_file = "Result.m3u"
        
        def get_filename(item, base_dir):
            if filename_format == "absolute":
                return os.path.join(base_dir, item)
            else:
                return item
        
        # 获取项目
        if self.ignore_directories.get():
            items = [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
        else:
            items = os.listdir(folder)
        
        matched = []
        missed = []
        
        for i, item in enumerate(lines, 1):
            # 移除括号
            clean_item = re.sub(r'\([^)]*\)', '', item).strip()
            
            # 计算相似度
            candidates = []
            perfect_match = None
            for it in items:
                similarity = difflib.SequenceMatcher(None, clean_item, it).ratio()
                if similarity == 1.0:
                    perfect_match = it
                    break  # 找到完全匹配，直接使用
                elif similarity > self.similarity_threshold:
                    candidates.append((it, similarity))
            
            if perfect_match:
                matched_item = perfect_match
            elif candidates:
                # 按相似度排序
                candidates.sort(key=lambda x: x[1], reverse=True)
                
                if len(candidates) == 1 or self.auto_select_highest.get():
                    matched_item = candidates[0][0]
                else:
                    # 多个候选，弹出选择对话框
                    choice = self.select_candidate(item, candidates)
                    if choice is None:
                        missed.append((i, item))
                        continue
                    matched_item = choice
            else:
                missed.append((i, item))
                continue
            
            matched.append((i, matched_item))
            items.remove(matched_item)
        
        unused = items
        
        # 计算序号
        len_matched = len(matched)
        if self.reverse.get():
            end = self.end_num.get()
            nums = [end - i * self.step.get() for i in range(len_matched)]
        else:
            start = self.start_num.get()
            nums = [start + i * self.step.get() for i in range(len_matched)]
        
        if preview:
            result = "预览结果:\n\n"
            result += "匹配的项目:\n"
            for idx, (orig_idx, item) in enumerate(matched, 1):
                num = nums[idx-1]
                new_name = self.generate_new_name(num, item)
                item_type = "目录" if os.path.isdir(os.path.join(folder, item)) else "文件"
                result += f"{idx}. {new_name} ({item_type})\n"
            result += "\n"
            if missed:
                result += f"未匹配的项目 (总共 {len(missed)} 个):\n"
                for orig_idx, item in missed:
                    result += f"  - {item} (第 {orig_idx} 行)\n"
                result += "\n"
            if unused:
                result += f"文件夹中未使用的项目 (总共 {len(unused)} 个):\n"
                for u in unused:
                    item_type = "目录" if os.path.isdir(os.path.join(folder, u)) else "文件"
                    result += f"  - {u} ({item_type})\n"
            return result
        
        # 如果仅生成列表
        if self.generate_list_only.get():
            if self.output_file_entry.get().strip():
                result_list_path = self.output_file_entry.get().strip()
                result_dir = os.path.dirname(result_list_path)
            elif self.output_folder_entry.get().strip():
                result_dir = self.output_folder_entry.get().strip()
                result_list_path = os.path.join(result_dir, result_file)
            else:
                result_dir = folder
                result_list_path = os.path.join(result_dir, result_file)
        else:
            if self.output_file_entry.get().strip():
                result_list_path = self.output_file_entry.get().strip()
                result_dir = os.path.dirname(result_list_path)
            elif self.output_folder_entry.get().strip():
                result_dir = self.output_folder_entry.get().strip()
                result_list_path = os.path.join(result_dir, result_file)
            else:
                result_dir = os.path.join(folder, 'Result')
                # 检查是否存在Result文件夹
                if os.path.exists(result_dir):
                    choice = messagebox.askyesno("确认", f"选择的文件夹中已存在 'Result' 文件夹。\n\n选择 '是' 以覆盖该文件夹，选择 '否' 将其当作排序项目并指定新的输出文件夹名称。")
                    if not choice:
                        # 要求输入新的输出文件夹名称
                        new_name = simpledialog.askstring("输入新的输出文件夹名称", "请输入新的输出文件夹名称:")
                        if new_name and new_name.strip():
                            result_dir = os.path.join(folder, new_name.strip())
                            result_list_path = os.path.join(result_dir, result_file)
                        else:
                            return "操作已取消。"
                    else:
                        result_list_path = os.path.join(result_dir, result_file)
                else:
                    result_list_path = os.path.join(result_dir, result_file)
        
        # 检查输出文件是否已存在
        if os.path.exists(result_list_path):
            choice = messagebox.askyesno("确认", f"输出文件 '{os.path.basename(result_list_path)}' 已存在。\n\n是否覆盖该文件？")
            if not choice:
                return "操作已取消。"
        
        if self.generate_list_only.get():
            with open(result_list_path, 'w', encoding='utf-8') as f:
                if output_format == "text":
                    f.write(f"# 文件排序 for 文件夹 {os.path.basename(folder)}\n")
                    f.write(f"# 使用配置 相似度阈值 {self.similarity_threshold} (仅生成列表)\n")
                    for idx, (orig_idx, item) in enumerate(matched, 1):
                        f.write(f"{get_filename(item, folder)}\n")
                    if missed:
                        f.write(f"# 未匹配项目 (总共 {len(missed)} 个未匹配)\n")
                        for orig_idx, item in missed:
                            f.write(f"# {item}(第 {orig_idx} 行)\n")
                    if unused:
                        f.write(f"# 文件夹中未使用的项目 (总共 {len(unused)} 个项目)\n")
                        for u in unused:
                            item_type = "目录" if os.path.isdir(os.path.join(folder, u)) else "文件"
                            f.write(f"# {get_filename(u, folder)} ({item_type})\n")
                elif output_format == "json":
                    data = {
                        "folder": os.path.basename(folder),
                        "threshold": self.similarity_threshold,
                        "mode": "list_only",
                        "matched": [get_filename(item, folder) for _, item in matched],
                        "missed": [{"line": orig_idx, "item": item} for orig_idx, item in missed],
                        "unused": [{"item": get_filename(u, folder), "type": "目录" if os.path.isdir(os.path.join(folder, u)) else "文件"} for u in unused]
                    }
                    json.dump(data, f, ensure_ascii=False, indent=2)
                elif output_format == "m3u":
                    f.write("#EXTM3U\n")
                    for idx, (orig_idx, item) in enumerate(matched, 1):
                        f.write(f"#EXTINF:-1,{item}\n")
                        f.write(f"{get_filename(item, folder)}\n")
            return f"列表生成完成！{os.path.basename(result_list_path)} 保存在 {result_dir}\n"
        
        # 确保输出目录存在
        os.makedirs(result_dir, exist_ok=True)
        
        # 计算新名称
        new_names = []
        for idx, (orig_idx, item) in enumerate(matched, 1):
            num = nums[idx-1]
            new_name = self.generate_new_name(num, item)
            new_names.append(new_name)
        
        # 复制或创建项目
        for idx, (orig_idx, item) in enumerate(matched, 1):
            new_name = new_names[idx-1]
            src_path = os.path.join(folder, item)
            dst_path = os.path.join(result_dir, new_name)
            if os.path.isfile(src_path):
                shutil.copy(src_path, dst_path)
            elif os.path.isdir(src_path):
                shutil.copytree(src_path, dst_path)
        
        # 生成Result.list
        with open(result_list_path, 'w', encoding='utf-8') as f:
            if output_format == "text":
                f.write(f"# 文件排序 for 文件夹 {os.path.basename(folder)}\n")
                f.write(f"# 使用配置 相似度阈值 {self.similarity_threshold}\n")
                for new_name in new_names:
                    item_type = "目录" if os.path.isdir(os.path.join(result_dir, new_name)) else "文件"
                    f.write(f"{get_filename(new_name, result_dir)} ({item_type})\n")
                if missed:
                    f.write(f"# 未匹配项目 (总共 {len(missed)} 个未匹配)\n")
                    for orig_idx, item in missed:
                        f.write(f"# {item}(第 {orig_idx} 行)\n")
                if unused:
                    f.write(f"# 文件夹中未使用的项目 (总共 {len(unused)} 个项目)\n")
                    for u in unused:
                        item_type = "目录" if os.path.isdir(os.path.join(folder, u)) else "文件"
                        f.write(f"# {get_filename(u, folder)} ({item_type})\n")
            elif output_format == "json":
                data = {
                    "folder": os.path.basename(folder),
                    "threshold": self.similarity_threshold,
                    "mode": "full",
                    "matched": [get_filename(new_name, result_dir) for new_name in new_names],
                    "missed": [{"line": orig_idx, "item": item} for orig_idx, item in missed],
                    "unused": [{"item": get_filename(u, folder), "type": "目录" if os.path.isdir(os.path.join(folder, u)) else "文件"} for u in unused]
                }
                json.dump(data, f, ensure_ascii=False, indent=2)
            elif output_format == "m3u":
                f.write("#EXTM3U\n")
                for idx, (orig_idx, item) in enumerate(matched, 1):
                    new_name = new_names[idx-1]
                    f.write(f"#EXTINF:-1,{item}\n")
                    f.write(f"{get_filename(new_name, result_dir)}\n")
        
        return f"处理完成！结果保存在 {result_dir}\n"
    
    def select_candidate(self, item, candidates):
        # 创建选择对话框
        dialog = tk.Toplevel(self.root)
        dialog.title(f"选择匹配文件 - {item}")
        dialog.geometry("500x400")
        dialog.configure(bg='#f0f0f0')
        
        ttk.Label(dialog, text=f"为 '{item}' 选择匹配的文件:", font=('Microsoft YaHei', 12)).pack(pady=10)
        
        frame = ttk.Frame(dialog, style='Card.TFrame')
        frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        listbox = tk.Listbox(frame, height=15, font=('Microsoft YaHei', 10), selectbackground='#cce7ff')
        for file, sim in candidates:
            listbox.insert(tk.END, f"{file} (相似度: {sim:.2f})")
        listbox.pack(fill="both", expand=True, padx=10, pady=10)
        
        button_frame = ttk.Frame(dialog, style='Card.TFrame')
        button_frame.pack(fill="x", padx=10, pady=10)
        
        selected = [None]
        
        def on_select():
            if listbox.curselection():
                idx = listbox.curselection()[0]
                selected[0] = candidates[idx][0]
            dialog.destroy()
        
        def on_skip():
            selected[0] = None
            dialog.destroy()
        
        ttk.Button(button_frame, text="选择", command=on_select, style='Accent.TButton').pack(side=tk.LEFT, padx=20)
        ttk.Button(button_frame, text="跳过", command=on_skip).pack(side=tk.RIGHT, padx=20)
        
        dialog.wait_window()
        return selected[0]
    
    def show_about(self):
        about_text = """
文件排序工具 (zh-lineup)

版本: 1.0
作者: GZYZhy

GitHub: https://github.com/GZYZhy/zh-lineup

使用教程:
1. 选择包含文件/目录的文件夹。
2. 选择目的列表输入方式：导入文件、Excel或手动输入。
3. 点击“预览”查看结果，或“运行”执行排序。
4. 在“配置”选项卡中调整匹配参数。

功能:
- 模糊匹配文件和目录
- 可配置相似度阈值
- 完全匹配优先
- 多种输入方式
- 预览和仅生成列表模式
- 跨平台GUI界面

许可证: Apache License 2.0
"""
        about_dialog = tk.Toplevel(self.root)
        about_dialog.title("关于")
        about_dialog.geometry("500x400")
        about_dialog.configure(bg='#f0f0f0')
        
        text = tk.Text(about_dialog, wrap=tk.WORD, font=('Microsoft YaHei', 10))
        scrollbar = ttk.Scrollbar(about_dialog, orient=tk.VERTICAL, command=text.yview)
        text.configure(yscrollcommand=scrollbar.set)
        text.pack(side=tk.LEFT, fill="both", expand=True, padx=10, pady=10)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        
        text.insert(tk.END, about_text.strip())
        text.config(state=tk.DISABLED)
        
        ttk.Button(about_dialog, text="关闭", command=about_dialog.destroy).pack(pady=10)

if __name__ == "__main__":
    root = tk.Tk()
    app = LineupApp(root)
    root.mainloop()