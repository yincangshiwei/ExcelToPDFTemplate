import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from ttkthemes import ThemedTk
import os
import sys
import socket
from pathlib import Path
import threading
from datetime import datetime
from core import ExcelToPDFProcessor
from font_manager import FontManagerWindow


class ExcelToPDFGUI:
    # 类级别的标志，防止重复检测
    _network_checked = False

    def __init__(self):
        # 企业环境检测
        # if not ExcelToPDFGUI._network_checked:
        #     ExcelToPDFGUI._network_checked = True
        #     if not self.check_network_connection():
        #         return
        # 创建主窗口
        self.root = ThemedTk(theme="arc")
        self.root.title("Excel转PDF模板（套打-电子签）")
        self.root.geometry("1200x700")  # 增加宽度以适应左右布局
        self.root.resizable(True, True)
        
        # 初始化处理器
        self.processor = ExcelToPDFProcessor()
        # 设置处理器的日志回调
        self.processor.set_gui_log_callback(self.add_operation_log)
        
        # 初始化字体相关变量
        self.default_font_combo = None
        self.chinese_font_combo = None
        
        # 初始化变量
        self.excel_path_var = tk.StringVar()
        self.pdf_template_var = tk.StringVar()
        self.output_folder_var = tk.StringVar(value=self.processor.get_desktop_path())
        self.sheet_name_var = tk.StringVar()
        self.title_row_var = tk.IntVar(value=3)
        self.start_row_var = tk.IntVar(value=4)
        self.filename_column_var = tk.StringVar()  # 新增：文件名指定列
        self.flatten_form_var = tk.BooleanVar(value=True)
        self.output_png_var = tk.BooleanVar(value=False)
        self.output_ppt_var = tk.BooleanVar(value=False)
        
        # 字体选择相关变量
        self.default_font_var = tk.StringVar()
        self.chinese_font_var = tk.StringVar()
        
        # 字段映射存储
        self.field_mapping_widgets = {}
        self.pdf_fields = []
        
        # 操作日志相关
        self.operation_logs = []
        self.max_logs = 1000  # 最大日志条数
        
        self.setup_ui()
        
        # 初始化字体列表
        self.refresh_fonts()
        
    def setup_ui(self):
        """设置用户界面"""
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建左右两个主要区域，使用place精确控制比例
        left_frame = ttk.Frame(main_frame)
        left_frame.place(relx=0, rely=0, relwidth=0.65, relheight=1)
        left_frame.columnconfigure(1, weight=1)
        
        right_frame = ttk.Frame(main_frame)
        right_frame.place(relx=0.65, rely=0, relwidth=0.35, relheight=1)
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(0, weight=1)
        
        # 设置左侧功能区域
        self.setup_left_panel(left_frame)
        
        # 设置右侧操作日志面板
        self.setup_operation_log_panel(right_frame)
        
    def setup_left_panel(self, parent):
        """设置左侧功能面板"""
        row = 0
        
        # 预设管理区域
        self.create_preset_section(parent, row)
        row += 1
        
        # 分隔线
        ttk.Separator(parent, orient='horizontal').grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        row += 1
        
        # 文件选择区域
        self.create_file_selection_section(parent, row)
        row += 4
        
        # 配置区域
        self.create_config_section(parent, row)
        row += 3
        
        # 字段映射区域
        self.create_field_mapping_section(parent, row)
        row += 1
        
        # 处理按钮和进度条
        self.create_process_section(parent, row)
        
    def setup_operation_log_panel(self, parent):
        """设置右侧操作日志面板"""
        # 创建日志面板框架
        log_frame = ttk.LabelFrame(parent, text="操作日志", padding="5")
        log_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(1, weight=1)
        
        # 日志控制按钮区域
        control_frame = ttk.Frame(log_frame)
        control_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        ttk.Button(control_frame, text="清空日志", command=self.clear_operation_logs).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(control_frame, text="导出日志", command=self.export_operation_logs).pack(side=tk.LEFT, padx=5)
        
        # 日志显示区域
        self.create_log_display_area(log_frame)
        
    def create_log_display_area(self, parent):
        """创建日志显示区域"""
        # 创建Text组件和滚动条
        text_frame = ttk.Frame(parent)
        text_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        self.log_text = tk.Text(text_frame, wrap=tk.WORD, state=tk.DISABLED, 
                               font=('Consolas', 9), bg='#f8f8f8')
        log_scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 配置文本标签样式
        self.log_text.tag_configure("info", foreground="#2E8B57")
        self.log_text.tag_configure("warning", foreground="#FF8C00")
        self.log_text.tag_configure("error", foreground="#DC143C")
        self.log_text.tag_configure("success", foreground="#228B22")
        
        # 添加欢迎信息
        self.add_operation_log("系统启动", "info", "Excel转PDF表单填充工具已启动")
        
    def add_operation_log(self, operation, level="info", message=""):
        """添加操作日志"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = {
            "timestamp": timestamp,
            "operation": operation,
            "level": level,
            "message": message
        }
        
        # 添加到日志列表
        self.operation_logs.append(log_entry)
        
        # 限制日志数量
        if len(self.operation_logs) > self.max_logs:
            self.operation_logs.pop(0)
        
        # 更新显示
        self.update_log_display(log_entry)
        
    def update_log_display(self, log_entry):
        """更新日志显示"""
        self.log_text.config(state=tk.NORMAL)
        
        # 格式化日志条目
        log_line = f"[{log_entry['timestamp']}] {log_entry['operation']}"
        if log_entry['message']:
            log_line += f": {log_entry['message']}"
        log_line += "\n"
        
        # 插入日志
        self.log_text.insert(tk.END, log_line, log_entry['level'])
        
        # 自动滚动到底部
        self.log_text.see(tk.END)
        
        self.log_text.config(state=tk.DISABLED)
        
    def clear_operation_logs(self):
        """清空操作日志"""
        self.operation_logs.clear()
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.add_operation_log("清空日志", "info", "操作日志已清空")
        
    def export_operation_logs(self):
        """导出操作日志"""
        if not self.operation_logs:
            messagebox.showwarning("警告", "没有日志可以导出")
            return
            
        filename = filedialog.asksaveasfilename(
            title="导出操作日志",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write("Excel转PDF表单填充工具 - 操作日志\n")
                    f.write("=" * 50 + "\n\n")
                    
                    for log in self.operation_logs:
                        f.write(f"[{log['timestamp']}] [{log['level'].upper()}] {log['operation']}")
                        if log['message']:
                            f.write(f": {log['message']}")
                        f.write("\n")
                        
                self.add_operation_log("导出日志", "success", f"日志已导出到: {filename}")
                messagebox.showinfo("成功", f"日志已成功导出到:\n{filename}")
            except Exception as e:
                self.add_operation_log("导出日志", "error", f"导出失败: {str(e)}")
                messagebox.showerror("错误", f"导出日志失败:\n{str(e)}")
        
    def create_preset_section(self, parent, row):
        """创建预设管理区域"""
        preset_frame = ttk.LabelFrame(parent, text="预设管理", padding="5")
        preset_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        
        # 预设按钮 - 左对齐
        ttk.Button(preset_frame, text="加载预设", command=self.load_preset).grid(row=0, column=0, padx=(0, 5), sticky=tk.W)
        ttk.Button(preset_frame, text="保存预设", command=self.save_preset).grid(row=0, column=1, padx=5, sticky=tk.W)
        ttk.Button(preset_frame, text="恢复默认", command=self.reset_to_default).grid(row=0, column=2, padx=5, sticky=tk.W)
        
    def create_file_selection_section(self, parent, row):
        """创建文件选择区域"""
        # Excel文件选择
        ttk.Label(parent, text="Excel文件:").grid(row=row, column=0, sticky=tk.W, pady=2)
        ttk.Entry(parent, textvariable=self.excel_path_var).grid(row=row, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(parent, text="浏览", command=self.browse_excel_file).grid(row=row, column=2, padx=5)
        
        # PDF模板选择
        ttk.Label(parent, text="PDF模板:").grid(row=row+1, column=0, sticky=tk.W, pady=2)
        ttk.Entry(parent, textvariable=self.pdf_template_var).grid(row=row+1, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(parent, text="浏览", command=self.browse_pdf_template).grid(row=row+1, column=2, padx=5)
        
        # 输出目录选择
        ttk.Label(parent, text="输出目录:").grid(row=row+2, column=0, sticky=tk.W, pady=2)
        ttk.Entry(parent, textvariable=self.output_folder_var).grid(row=row+2, column=1, sticky=(tk.W, tk.E), padx=5)
        ttk.Button(parent, text="浏览", command=self.browse_output_folder).grid(row=row+2, column=2, padx=5)
        
    def create_config_section(self, parent, row):
        """创建配置区域"""
        config_frame = ttk.LabelFrame(parent, text="配置参数", padding="5")
        config_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        config_frame.columnconfigure(1, weight=1)
        config_frame.columnconfigure(3, weight=1)
        config_frame.columnconfigure(5, weight=1)
        config_frame.columnconfigure(7, weight=1)
        
        # 第一行：Sheet名称、表头行、数据起始行、文件名列（紧凑排列）
        ttk.Label(config_frame, text="Sheet名称:").grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Entry(config_frame, textvariable=self.sheet_name_var, width=12).grid(row=0, column=1, sticky=tk.W, padx=(2,10))
        
        ttk.Label(config_frame, text="表头行:").grid(row=0, column=2, sticky=tk.W, pady=2)
        ttk.Spinbox(config_frame, from_=1, to=100, textvariable=self.title_row_var, width=8).grid(row=0, column=3, sticky=tk.W, padx=(2,10))
        
        ttk.Label(config_frame, text="数据起始行:").grid(row=0, column=4, sticky=tk.W, pady=2)
        ttk.Spinbox(config_frame, from_=1, to=100, textvariable=self.start_row_var, width=8).grid(row=0, column=5, sticky=tk.W, padx=(2,10))
        
        ttk.Label(config_frame, text="文件名列:").grid(row=0, column=6, sticky=tk.W, pady=2)
        filename_entry = ttk.Entry(config_frame, textvariable=self.filename_column_var, width=8)
        filename_entry.grid(row=0, column=7, sticky=tk.W, padx=(2,0))
        
        # 添加文件名列的提示
        ttk.Label(config_frame, text="(可选，如A列。不填则使用数字编号)", font=('TkDefaultFont', 8)).grid(row=0, column=8, sticky=tk.W, padx=(5,0), pady=2)
        
        # 第二行：字体配置
        font_frame = ttk.Frame(config_frame)
        font_frame.grid(row=1, column=0, columnspan=9, sticky=(tk.W, tk.E), pady=(8,2))
        
        ttk.Button(font_frame, text="字体管理", command=self.open_font_manager).pack(side=tk.LEFT, padx=(0,20))
        
        ttk.Label(font_frame, text="默认字体:").pack(side=tk.LEFT)
        self.default_font_combo = ttk.Combobox(font_frame, textvariable=self.default_font_var, width=15, state="readonly")
        self.default_font_combo.pack(side=tk.LEFT, padx=(2,10))
        
        ttk.Label(font_frame, text="中文字体:").pack(side=tk.LEFT)
        self.chinese_font_combo = ttk.Combobox(font_frame, textvariable=self.chinese_font_var, width=15, state="readonly")
        self.chinese_font_combo.pack(side=tk.LEFT, padx=(2,0))
        
        # 第三行：输出选项（紧凑排列）
        options_frame = ttk.Frame(config_frame)
        options_frame.grid(row=2, column=0, columnspan=9, sticky=(tk.W, tk.E), pady=(8,2))
        
        ttk.Checkbutton(options_frame, text="扁平化表单(将表单转为静态文本)", variable=self.flatten_form_var).pack(side=tk.LEFT)
        ttk.Checkbutton(options_frame, text="输出PNG图片", variable=self.output_png_var).pack(side=tk.LEFT, padx=(20,0))
        ttk.Checkbutton(options_frame, text="输出PPT", variable=self.output_ppt_var).pack(side=tk.LEFT, padx=(20,0))
        
    def create_field_mapping_section(self, parent, row):
        """创建字段映射区域"""
        mapping_frame = ttk.LabelFrame(parent, text="字段映射配置", padding="5")
        mapping_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        mapping_frame.columnconfigure(0, weight=1)
        mapping_frame.rowconfigure(1, weight=1)
        
        # 加载字段按钮
        button_frame = ttk.Frame(mapping_frame)
        button_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        ttk.Button(button_frame, text="加载PDF表单字段", command=self.load_pdf_fields).pack(side=tk.LEFT)
        ttk.Label(button_frame, text="选择PDF模板后点击此按钮加载表单字段").pack(side=tk.LEFT, padx=10)
        
        # 字段映射滚动区域
        self.create_scrollable_mapping_area(mapping_frame)
        
    def create_scrollable_mapping_area(self, parent):
        """创建可滚动的字段映射区域"""
        # 创建Canvas和Scrollbar
        canvas = tk.Canvas(parent, height=200)
        v_scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        h_scrollbar = ttk.Scrollbar(parent, orient="horizontal", command=canvas.xview)
        self.mapping_frame = ttk.Frame(canvas)
        
        # 配置滚动
        self.mapping_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.mapping_frame, anchor="nw")
        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # 布局
        canvas.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=2, column=0, sticky=(tk.W, tk.E))
        
        # 鼠标滚轮绑定
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind("<MouseWheel>", _on_mousewheel)
        
        # Shift+鼠标滚轮进行水平滚动
        def _on_shift_mousewheel(event):
            canvas.xview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind("<Shift-MouseWheel>", _on_shift_mousewheel)
        
    def create_process_section(self, parent, row):
        """创建处理按钮和进度条区域"""
        process_frame = ttk.Frame(parent)
        process_frame.grid(row=row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        process_frame.columnconfigure(1, weight=1)
        
        # 处理按钮
        self.process_button = tk.Button(process_frame, text="开始处理", command=self.start_processing,
                                      fg="white", bg="#0078D7", font=('Microsoft YaHei UI', 10, 'bold'),
                                      padx=10, pady=5, relief='flat', borderwidth=0,
                                      activebackground="#005A9E", activeforeground="white",
                                      highlightthickness=0, bd=0)
        self.process_button.grid(row=0, column=0, padx=5)
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(process_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=5)
        
        # 状态标签
        self.status_label = ttk.Label(process_frame, text="就绪")
        self.status_label.grid(row=0, column=2, padx=5)
        
    def browse_excel_file(self):
        """浏览Excel文件"""
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if filename:
            self.processor.logger.info(f"用户选择Excel文件: {filename}")
            self.excel_path_var.set(filename)
            self.add_operation_log("选择Excel文件", "info", f"已选择: {os.path.basename(filename)}")
        else:
            self.add_operation_log("选择Excel文件", "warning", "用户取消了文件选择")
            
    def browse_pdf_template(self):
        """浏览PDF模板文件"""
        # 默认路径为resources/template目录
        initial_dir = os.path.join(os.getcwd(), "resources", "template")
        if not os.path.exists(initial_dir):
            initial_dir = os.getcwd()
            
        filename = filedialog.askopenfilename(
            title="选择PDF模板文件",
            initialdir=initial_dir,
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if filename:
            self.processor.logger.info(f"用户选择PDF模板文件: {filename}")
            self.pdf_template_var.set(filename)
            self.add_operation_log("选择PDF模板", "info", f"已选择: {os.path.basename(filename)}")
        else:
            self.add_operation_log("选择PDF模板", "warning", "用户取消了文件选择")
            
    def browse_output_folder(self):
        """浏览输出目录"""
        folder = filedialog.askdirectory(
            title="选择输出目录",
            initialdir=self.output_folder_var.get()
        )
        if folder:
            self.processor.logger.info(f"用户选择输出目录: {folder}")
            self.output_folder_var.set(folder)
            self.add_operation_log("选择输出目录", "info", f"已选择: {folder}")
        else:
            self.add_operation_log("选择输出目录", "warning", "用户取消了目录选择")
            
    def load_preset(self):
        """加载预设"""
        # 默认路径为presets目录
        initial_dir = os.path.join(os.getcwd(), "presets")
        if not os.path.exists(initial_dir):
            initial_dir = os.getcwd()
            
        filename = filedialog.askopenfilename(
            title="选择预设文件",
            initialdir=initial_dir,
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
        )
        if filename:
            self.processor.logger.info(f"用户选择加载预设文件: {filename}")
            self.add_operation_log("加载预设", "info", f"正在加载: {os.path.basename(filename)}")
            success, message = self.processor.load_preset(filename)
            if success:
                self.update_ui_from_processor()
                self.add_operation_log("加载预设", "success", f"预设加载成功: {os.path.basename(filename)}")
                messagebox.showinfo("成功", message)
            else:
                self.add_operation_log("加载预设", "error", f"预设加载失败: {message}")
                messagebox.showerror("错误", message)
        else:
            self.add_operation_log("加载预设", "warning", "用户取消了预设加载")
                
    def save_preset(self):
        """保存预设"""
        # 默认路径为presets目录
        initial_dir = os.path.join(os.getcwd(), "presets")
        if not os.path.exists(initial_dir):
            os.makedirs(initial_dir, exist_ok=True)
        
        # 如果有PDF模板文件，使用其文件名作为默认文件名
        initial_filename = ""
        pdf_template_path = self.pdf_template_var.get().strip()
        if pdf_template_path and os.path.exists(pdf_template_path):
            # 获取PDF文件名（不含扩展名）
            pdf_filename = os.path.splitext(os.path.basename(pdf_template_path))[0]
            initial_filename = f"{pdf_filename}.json"
            
        filename = filedialog.asksaveasfilename(
            title="保存预设文件",
            initialdir=initial_dir,
            initialfile=initial_filename,
            defaultextension=".json",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
        )
        if filename:
            self.processor.logger.info(f"用户选择保存预设文件: {filename}")
            self.add_operation_log("保存预设", "info", f"正在保存: {os.path.basename(filename)}")
            self.update_processor_from_ui()
            success, message = self.processor.save_preset(filename)
            if success:
                self.add_operation_log("保存预设", "success", f"预设保存成功: {os.path.basename(filename)}")
                messagebox.showinfo("成功", message)
            else:
                self.add_operation_log("保存预设", "error", f"预设保存失败: {message}")
                messagebox.showerror("错误", message)
        else:
            self.add_operation_log("保存预设", "warning", "用户取消了预设保存")
                
    def reset_to_default(self):
        """恢复默认配置"""
        if messagebox.askyesno("确认", "确定要恢复默认配置吗？这将清除当前所有设置。"):
            self.processor.logger.info("用户选择恢复默认配置")
            self.add_operation_log("恢复默认配置", "info", "正在恢复默认配置...")
            message = self.processor.reset_to_default()
            self.update_ui_from_processor()
            self.clear_field_mapping()
            self.add_operation_log("恢复默认配置", "success", "已恢复到默认配置")
            messagebox.showinfo("成功", message)
        else:
            self.add_operation_log("恢复默认配置", "warning", "用户取消了恢复默认配置操作")
    
    def refresh_fonts(self):
        """刷新字体列表"""
        try:
            # 重新加载字体库
            self.processor.load_available_fonts()
            
            # 获取默认字体和中文字体列表
            default_fonts = self.processor.get_default_fonts()
            chinese_fonts = self.processor.get_chinese_fonts()
            
            # 更新默认字体下拉框
            if self.default_font_combo:
                self.default_font_combo['values'] = default_fonts
                if default_fonts:
                    # 尝试找到包含calibri的字体作为默认选择
                    default_selection = None
                    for font in default_fonts:
                        if 'calibri' in font.lower():
                            default_selection = font
                            break
                    if not default_selection:
                        default_selection = default_fonts[0]
                    self.default_font_var.set(default_selection)
            
            # 更新中文字体下拉框
            if self.chinese_font_combo:
                self.chinese_font_combo['values'] = chinese_fonts
                if chinese_fonts:
                    # 尝试找到包含simhei的字体作为默认选择
                    chinese_selection = None
                    for font in chinese_fonts:
                        if 'simhei' in font.lower():
                            chinese_selection = font
                            break
                    if not chinese_selection:
                        chinese_selection = chinese_fonts[0]
                    self.chinese_font_var.set(chinese_selection)
            
            total_fonts = len(default_fonts) + len(chinese_fonts)
            self.add_operation_log("刷新字体", "success", f"已加载 {total_fonts} 个字体文件 (默认:{len(default_fonts)}, 中文:{len(chinese_fonts)})")
            
        except Exception as e:
            self.add_operation_log("刷新字体", "error", f"刷新字体失败: {str(e)}")
    
    def open_font_manager(self):
        """打开字体管理窗口"""
        FontManagerWindow(self.root, self.processor, self.refresh_fonts, self.add_operation_log)
            
    def load_pdf_fields(self):
        """加载PDF表单字段"""
        pdf_path = self.pdf_template_var.get().strip()
        if not pdf_path:
            self.add_operation_log("加载PDF字段", "warning", "请先选择PDF模板文件")
            messagebox.showwarning("警告", "请先选择PDF模板文件")
            return
            
        if not os.path.exists(pdf_path):
            self.add_operation_log("加载PDF字段", "error", "PDF模板文件不存在")
            messagebox.showerror("错误", "PDF模板文件不存在")
            return
            
        try:
            self.processor.logger.info(f"用户开始加载PDF表单字段: {pdf_path}")
            self.add_operation_log("加载PDF字段", "info", f"正在分析PDF模板: {os.path.basename(pdf_path)}")
            self.pdf_fields = self.processor.get_pdf_form_keys(pdf_path)
            if not self.pdf_fields:
                self.add_operation_log("加载PDF字段", "warning", "PDF文件中没有找到表单字段")
                messagebox.showwarning("警告", "PDF文件中没有找到表单字段")
                return
                
            self.create_field_mapping_widgets()
            success_msg = f"成功加载了 {len(self.pdf_fields)} 个表单字段"
            self.processor.logger.info(f"PDF字段加载成功: {success_msg}")
            self.add_operation_log("加载PDF字段", "success", success_msg)
            messagebox.showinfo("成功", success_msg)
            
        except Exception as e:
            error_msg = f"加载PDF字段失败: {e}"
            self.processor.logger.error(error_msg)
            self.add_operation_log("加载PDF字段", "error", error_msg)
            messagebox.showerror("错误", error_msg)
            
    def create_field_mapping_widgets(self):
        """创建字段映射控件"""
        # 清除现有控件
        for widget in self.mapping_frame.winfo_children():
            widget.destroy()
        self.field_mapping_widgets.clear()
        
        # 计算分列显示
        total_fields = len(self.pdf_fields)
        fields_per_column = (total_fields + 1) // 2  # 向上取整
        
        # 创建左列表头
        ttk.Label(self.mapping_frame, text="PDF字段名", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        ttk.Label(self.mapping_frame, text="类型", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(self.mapping_frame, text="映射值", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        
        # 创建右列表头（如果有足够的字段需要分列）
        if total_fields > fields_per_column:
            ttk.Label(self.mapping_frame, text="PDF字段名", font=("Arial", 10, "bold")).grid(row=0, column=4, padx=(20, 5), pady=5, sticky=tk.W)
            ttk.Label(self.mapping_frame, text="类型", font=("Arial", 10, "bold")).grid(row=0, column=5, padx=5, pady=5)
            ttk.Label(self.mapping_frame, text="映射值", font=("Arial", 10, "bold")).grid(row=0, column=6, padx=5, pady=5, sticky=tk.W)
        
        # 为每个PDF字段创建映射控件
        for i, field_name in enumerate(self.pdf_fields):
            # 确定显示位置（左列还是右列）
            if i < fields_per_column:
                # 左列
                row_pos = i + 1
                col_offset = 0
            else:
                # 右列
                row_pos = i - fields_per_column + 1
                col_offset = 4
            
            # 字段名标签
            padx_left = (20, 5) if col_offset == 4 else 5
            ttk.Label(self.mapping_frame, text=field_name).grid(row=row_pos, column=col_offset, padx=padx_left, pady=2, sticky=tk.W)
            
            # 类型选择
            type_var = tk.StringVar(value="Excel列")
            type_combo = ttk.Combobox(self.mapping_frame, textvariable=type_var, values=["Excel列", "Excel列-图片", "自定义值"], width=12, state="readonly")
            type_combo.grid(row=row_pos, column=col_offset+1, padx=5, pady=2)
            
            # 映射值输入
            value_var = tk.StringVar()
            value_entry = ttk.Entry(self.mapping_frame, textvariable=value_var)
            value_entry.grid(row=row_pos, column=col_offset+2, padx=5, pady=2, sticky=(tk.W, tk.E))
            
            # 存储控件引用
            self.field_mapping_widgets[field_name] = {
                "type_var": type_var,
                "value_var": value_var,
                "type_combo": type_combo,
                "value_entry": value_entry
            }
            
            # 绑定类型变化事件
            type_combo.bind("<<ComboboxSelected>>", lambda e, fn=field_name: self.on_type_changed(fn))
            
        # 配置列权重
        self.mapping_frame.columnconfigure(2, weight=1)
        if total_fields > fields_per_column:
            self.mapping_frame.columnconfigure(6, weight=1)
        
    def on_type_changed(self, field_name):
        """当映射类型改变时的处理"""
        widgets = self.field_mapping_widgets[field_name]
        type_val = widgets["type_var"].get()
        
        if type_val == "Excel列":
            widgets["value_entry"].configure(state="normal")
            if not widgets["value_var"].get():
                widgets["value_var"].set("A")  # 默认值
        elif type_val == "Excel列-图片":
            widgets["value_entry"].configure(state="normal")
            if not widgets["value_var"].get() or widgets["value_var"].get() in ["A", "B", "C"]:
                widgets["value_var"].set("A")  # 默认值
        else:
            widgets["value_entry"].configure(state="normal")
            if widgets["value_var"].get() in ["A", "B", "C"]:
                widgets["value_var"].set("")  # 清除默认Excel列值
                
    def clear_field_mapping(self):
        """清除字段映射"""
        for widget in self.mapping_frame.winfo_children():
            widget.destroy()
        self.field_mapping_widgets.clear()
        self.pdf_fields.clear()
        
    def update_ui_from_processor(self):
        """从处理器更新UI"""
        self.excel_path_var.set(self.processor.excel_path)
        self.pdf_template_var.set(self.processor.pdf_template_path)
        self.output_folder_var.set(self.processor.output_folder)
        self.sheet_name_var.set(self.processor.sheet_name or "")
        self.title_row_var.set(self.processor.title_row)
        self.start_row_var.set(self.processor.start_row)
        self.filename_column_var.set(getattr(self.processor, 'filename_column', '') or "")  # 新增：文件名列
        self.flatten_form_var.set(self.processor.flatten_form)
        self.output_png_var.set(self.processor.output_png)
        self.output_ppt_var.set(self.processor.output_ppt)
        
        # 更新字体配置
        self.default_font_var.set(getattr(self.processor, 'default_font', ''))
        self.chinese_font_var.set(getattr(self.processor, 'chinese_font', ''))
        
        # 如果有字段映射，尝试加载PDF字段并设置映射
        if self.processor.field_mapping and self.processor.pdf_template_path:
            try:
                self.pdf_fields = self.processor.get_pdf_form_keys(self.processor.pdf_template_path)
                self.create_field_mapping_widgets()
                
                # 设置字段映射值
                for field_name, mapping in self.processor.field_mapping.items():
                    if field_name in self.field_mapping_widgets:
                        widgets = self.field_mapping_widgets[field_name]
                        if isinstance(mapping, dict):
                            is_excel_col = mapping.get("is_excel_col", True)
                            is_excel_image = mapping.get("is_excel_image", False)
                            val = mapping.get("val", "")
                            if is_excel_image:
                                widgets["type_var"].set("Excel列-图片")
                            elif is_excel_col:
                                widgets["type_var"].set("Excel列")
                            else:
                                widgets["type_var"].set("自定义值")
                            widgets["value_var"].set(val)
                        else:
                            # 向后兼容
                            widgets["type_var"].set("Excel列")
                            widgets["value_var"].set(str(mapping))
            except:
                pass  # 忽略错误，可能PDF文件不存在
                
    def update_processor_from_ui(self):
        """从UI更新处理器"""
        self.processor.excel_path = self.excel_path_var.get().strip()
        self.processor.pdf_template_path = self.pdf_template_var.get().strip()
        self.processor.output_folder = self.output_folder_var.get().strip()
        self.processor.sheet_name = self.sheet_name_var.get().strip() or None
        self.processor.title_row = self.title_row_var.get()
        self.processor.start_row = self.start_row_var.get()
        self.processor.filename_column = self.filename_column_var.get().strip() or None  # 新增：文件名列
        self.processor.flatten_form = self.flatten_form_var.get()
        self.processor.output_png = self.output_png_var.get()
        self.processor.output_ppt = self.output_ppt_var.get()
        
        # 更新字体配置
        self.processor.default_font = self.default_font_var.get().strip()
        self.processor.chinese_font = self.chinese_font_var.get().strip()
        
        # 更新字段映射
        self.processor.field_mapping = {}
        for field_name, widgets in self.field_mapping_widgets.items():
            type_val = widgets["type_var"].get()
            value_val = widgets["value_var"].get()
            
            # 只对Excel列进行strip，自定义值保留原始格式（包括前后空格）
            if type_val in ["Excel列", "Excel列-图片"]:
                value_val = value_val.strip()
            
            if value_val:  # 只有非空值才添加
                self.processor.field_mapping[field_name] = {
                    "is_excel_col": type_val == "Excel列",
                    "is_excel_image": type_val == "Excel列-图片",
                    "val": value_val
                }
                
    def start_processing(self):
        """开始处理"""
        # 验证输入
        if not self.excel_path_var.get().strip():
            self.add_operation_log("开始处理", "warning", "请选择Excel文件")
            messagebox.showwarning("警告", "请选择Excel文件")
            return
            
        if not self.pdf_template_var.get().strip():
            self.add_operation_log("开始处理", "warning", "请选择PDF模板文件")
            messagebox.showwarning("警告", "请选择PDF模板文件")
            return
            
        if not self.output_folder_var.get().strip():
            self.add_operation_log("开始处理", "warning", "请选择输出目录")
            messagebox.showwarning("警告", "请选择输出目录")
            return
            
        if not self.field_mapping_widgets:
            self.add_operation_log("开始处理", "warning", "请先加载PDF表单字段")
            messagebox.showwarning("警告", "请先加载PDF表单字段")
            return
            
        # 检查是否有字段映射
        has_mapping = False
        for widgets in self.field_mapping_widgets.values():
            value = widgets["value_var"].get()
            type_val = widgets["type_var"].get()
            
            # 对于Excel列，检查strip后的值；对于自定义值，检查原始值
            if type_val == "Excel列":
                if value.strip():
                    has_mapping = True
                    break
            else:
                if value:  # 自定义值保留原始格式，只要不是空字符串就算有映射
                    has_mapping = True
                    break
                
        if not has_mapping:
            self.add_operation_log("开始处理", "warning", "请至少设置一个字段映射")
            messagebox.showwarning("警告", "请至少设置一个字段映射")
            return
            
        # 更新处理器配置
        self.update_processor_from_ui()
        
        # 记录处理开始信息
        excel_file = os.path.basename(self.excel_path_var.get())
        pdf_template = os.path.basename(self.pdf_template_var.get())
        mapping_count = sum(1 for widgets in self.field_mapping_widgets.values() 
                           if widgets["value_var"].get().strip())
        
        self.add_operation_log("开始处理", "info", 
                              f"Excel文件: {excel_file}, PDF模板: {pdf_template}, 字段映射: {mapping_count}个")
        
        # 禁用处理按钮
        self.process_button.configure(state="disabled")
        self.progress_var.set(0)
        self.status_label.configure(text="处理中...")
        
        # 在新线程中处理
        def process_thread():
            try:
                # 记录开始处理的日志
                self.processor.logger.info("用户开始处理Excel转PDF任务")
                result = self.processor.process_excel_to_pdf(self.update_progress)
                
                # 在主线程中更新UI
                self.root.after(0, lambda: self.process_completed(result))
            except Exception as e:
                error_msg = str(e)
                self.processor.logger.error(f"处理线程异常: {error_msg}")
                self.root.after(0, lambda: self.process_error(error_msg))
                
        threading.Thread(target=process_thread, daemon=True).start()
        
    def update_progress(self, progress, status):
        """更新进度"""
        self.root.after(0, lambda: self._update_progress_ui(progress, status))
        
    def _update_progress_ui(self, progress, status):
        """在主线程中更新进度UI"""
        self.progress_var.set(progress)
        self.status_label.configure(text=status)
        
    def process_completed(self, result):
        """处理完成"""
        self.process_button.configure(state="normal")
        self.progress_var.set(100)
        
        if result["success"]:
            message = f"处理完成！\n总行数: {result['total_rows']}\n成功: {result['success_count']}\n失败: {result['error_count']}"
            if result["error_messages"]:
                message += "\n\n错误详情:\n" + "\n".join(result["error_messages"][:5])  # 只显示前5个错误
                if len(result["error_messages"]) > 5:
                    message += f"\n... 还有 {len(result['error_messages']) - 5} 个错误"
            
            self.status_label.configure(text="处理完成")
            self.processor.logger.info(f"处理任务完成 - 成功: {result['success_count']}, 失败: {result['error_count']}")
            
            # 添加操作日志
            log_msg = f"总行数: {result['total_rows']}, 成功: {result['success_count']}, 失败: {result['error_count']}"
            if result['error_count'] == 0:
                self.add_operation_log("处理完成", "success", log_msg)
            else:
                self.add_operation_log("处理完成", "warning", log_msg)
            
            messagebox.showinfo("处理完成", message)
        else:
            self.status_label.configure(text="处理失败")
            self.processor.logger.error(f"处理任务失败: {result['error']}")
            self.add_operation_log("处理失败", "error", result["error"])
            messagebox.showerror("处理失败", result["error"])
            
    def process_error(self, error_msg):
        """处理错误"""
        self.process_button.configure(state="normal")
        self.progress_var.set(0)
        self.status_label.configure(text="处理失败")
        self.processor.logger.error(f"处理过程发生错误: {error_msg}")
        self.add_operation_log("处理错误", "error", f"处理过程中发生错误: {error_msg}")
        messagebox.showerror("处理错误", f"处理过程中发生错误: {error_msg}")

    def check_network_connection(self):
        """检查网络连接"""
        # 创建检测中的提示窗口
        check_window = tk.Tk()
        check_window.title("网络检测")
        check_window.geometry("300x100")
        check_window.resizable(False, False)

        # 居中显示窗口
        check_window.eval('tk::PlaceWindow . center')

        # 添加检测中的标签
        label = tk.Label(check_window, text="企业环境检测中...", font=('Microsoft YaHei UI', 12))
        label.pack(expand=True)

        # 更新窗口显示
        check_window.update()

        try:
            # 创建socket连接测试10.0.182.21:22
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(3)  # 设置3秒超时
            result = sock.connect_ex(('10.0.182.21', 22))
            sock.close()

            # 关闭检测窗口
            check_window.destroy()

            if result == 0:
                return True
            else:
                self.show_network_error()
                return False
        except Exception:
            # 关闭检测窗口
            check_window.destroy()
            self.show_network_error()
            return False

    def show_network_error(self):
        """显示网络错误提示窗口"""
        try:
            # 创建一个临时的根窗口用于显示错误消息
            temp_root = tk.Tk()
            temp_root.withdraw()  # 隐藏主窗口

            # 显示错误消息
            messagebox.showerror(
                "网络连接错误",
                "企业工具请在企业网络环境使用！",
                parent=temp_root
            )

            # 销毁临时窗口
            temp_root.destroy()
        except Exception:
            # 如果GUI创建失败，直接退出
            pass
        finally:
            # 确保程序退出
            sys.exit(1)
        
    def run(self):
        """运行GUI"""
        self.root.mainloop()


if __name__ == "__main__":
    app = ExcelToPDFGUI()
    app.run()