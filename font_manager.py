import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import shutil
import subprocess
import sys
from pathlib import Path


class FontManagerWindow:
    """字体管理窗口"""
    
    def __init__(self, parent, processor, refresh_callback, log_callback):
        self.parent = parent
        self.processor = processor
        self.refresh_callback = refresh_callback
        self.log_callback = log_callback
        
        # 创建字体管理窗口
        self.window = tk.Toplevel(parent)
        self.window.title("字体管理")
        self.window.geometry("600x350")
        self.window.resizable(True, True)
        
        # 设置窗口图标（如果有的话）
        try:
            self.window.iconbitmap(parent.iconbitmap())
        except:
            pass
        
        # 使窗口模态
        self.window.transient(parent)
        self.window.grab_set()
        
        # 居中显示窗口
        self.center_window()
        
        # 初始化变量
        self.font_base_path_var = tk.StringVar(value=self.processor.font_base_path)
        
        # 设置UI
        self.setup_ui()
        
    def center_window(self):
        """居中显示窗口"""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f"{width}x{height}+{x}+{y}")
        
    def setup_ui(self):
        """设置用户界面"""
        # 创建主框架
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 字体库路径配置区域
        self.create_path_config_section(main_frame)
        
        # 分隔线
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=10)
        
        # 字体上传区域
        self.create_upload_section(main_frame)
        
        # 按钮区域
        self.create_button_section(main_frame)
        
    def create_path_config_section(self, parent):
        """创建字体库路径配置区域"""
        path_frame = ttk.LabelFrame(parent, text="字体库路径配置", padding="10")
        path_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 当前字体库路径
        ttk.Label(path_frame, text="字体库路径:").grid(row=0, column=0, sticky=tk.W, pady=5)
        path_entry = ttk.Entry(path_frame, textvariable=self.font_base_path_var, width=50)
        path_entry.grid(row=0, column=1, sticky="we", padx=(10, 5), pady=5)
        ttk.Button(path_frame, text="浏览", command=self.browse_font_path).grid(row=0, column=2, padx=5, pady=5)
        ttk.Button(path_frame, text="打开", command=self.open_font_path).grid(row=0, column=3, padx=5, pady=5)
        
        # 配置列权重
        path_frame.columnconfigure(1, weight=1)
        
        # 说明文字
        info_label = ttk.Label(path_frame, text="字体库包含两个子目录：default（默认字体）和 zh（中文字体）", 
                              font=('TkDefaultFont', 8), foreground='gray')
        info_label.grid(row=1, column=0, columnspan=4, sticky=tk.W, pady=(5, 0))
        
    def create_upload_section(self, parent):
        """创建字体上传区域"""
        upload_frame = ttk.LabelFrame(parent, text="字体文件上传", padding="10")
        upload_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 上传按钮区域
        button_frame = ttk.Frame(upload_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(button_frame, text="上传默认字体", command=lambda: self.upload_fonts("default")).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="上传中文字体", command=lambda: self.upload_fonts("zh")).pack(side=tk.LEFT)
        
        # 说明文字
        info_text = "支持的字体格式：.ttf, .otf\n支持多选文件批量上传，上传的字体文件将复制到对应的字体库目录中"
        ttk.Label(upload_frame, text=info_text, font=('TkDefaultFont', 8), foreground='gray').pack(anchor=tk.W, pady=(10, 0))
        
    def create_button_section(self, parent):
        """创建按钮区域"""
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=10)
        
        # 左侧按钮
        left_frame = ttk.Frame(button_frame)
        left_frame.pack(side=tk.LEFT)
        
        ttk.Button(left_frame, text="重置为默认", command=self.reset_to_default).pack(side=tk.LEFT)
        
        # 右侧按钮
        right_frame = ttk.Frame(button_frame)
        right_frame.pack(side=tk.RIGHT)
        
        ttk.Button(right_frame, text="取消", command=self.cancel).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(right_frame, text="保存", command=self.save_settings).pack(side=tk.LEFT)
        
    def browse_font_path(self):
        """浏览字体库路径"""
        current_path = self.font_base_path_var.get()
        
        # 如果是相对路径，转换为绝对路径用于初始目录
        if not os.path.isabs(current_path):
            initial_dir = os.path.abspath(current_path)
        else:
            initial_dir = current_path
            
        folder = filedialog.askdirectory(
            title="选择字体库目录",
            initialdir=initial_dir if os.path.exists(initial_dir) else os.getcwd()
        )
        if folder:
            self.font_base_path_var.set(folder)
            
    def open_font_path(self):
        """打开字体库路径"""
        font_path = self.font_base_path_var.get()
        
        # 如果是相对路径，转换为绝对路径
        if not os.path.isabs(font_path):
            font_path = os.path.abspath(font_path)
            
        try:
            if os.path.exists(font_path):
                if sys.platform == "win32":
                    subprocess.run(["explorer", font_path])
                elif sys.platform == "darwin":  # macOS
                    subprocess.run(["open", font_path])
                else:  # Linux
                    subprocess.run(["xdg-open", font_path])
            else:
                messagebox.showerror("错误", f"字体库路径不存在: {font_path}")
        except Exception as e:
            messagebox.showerror("错误", f"无法打开字体库路径: {str(e)}")
            
    def upload_fonts(self, font_type):
        """上传字体文件到指定类型目录（支持多选）"""
        filetypes = [("字体文件", "*.ttf *.otf"), ("TrueType字体", "*.ttf"), ("OpenType字体", "*.otf"), ("所有文件", "*.*")]
        filenames = filedialog.askopenfilenames(
            title=f"选择{'默认' if font_type == 'default' else '中文'}字体文件（支持多选）",
            filetypes=filetypes
        )
        
        if not filenames:
            return
            
        # 确保目标目录存在
        font_base_path = self.font_base_path_var.get()
        if not os.path.isabs(font_base_path):
            font_base_path = os.path.abspath(font_base_path)
            
        target_dir = os.path.join(font_base_path, font_type)
        os.makedirs(target_dir, exist_ok=True)
        
        success_count = 0
        error_count = 0
        
        for filename in filenames:
            try:
                # 检查文件扩展名
                if not filename.lower().endswith(('.ttf', '.otf')):
                    self.log_callback("上传字体", "warning", f"跳过非字体文件: {os.path.basename(filename)}")
                    continue
                    
                # 复制文件到目标目录
                target_path = os.path.join(target_dir, os.path.basename(filename))
                
                # 如果文件已存在，询问是否覆盖
                if os.path.exists(target_path):
                    if not messagebox.askyesno("文件已存在", f"字体文件 '{os.path.basename(filename)}' 已存在，是否覆盖？"):
                        continue
                        
                shutil.copy2(filename, target_path)
                success_count += 1
                self.log_callback("上传字体", "success", f"已上传: {os.path.basename(filename)}")
                
            except Exception as e:
                error_count += 1
                error_msg = f"上传 {os.path.basename(filename)} 失败: {str(e)}"
                self.log_callback("上传字体", "error", error_msg)
                
        # 显示结果
        if success_count > 0:
            messagebox.showinfo("上传完成", f"成功上传 {success_count} 个字体文件到{'默认' if font_type == 'default' else '中文'}字体目录" + 
                              (f"，{error_count} 个失败" if error_count > 0 else ""))
        elif error_count > 0:
            messagebox.showerror("上传失败", f"所有 {error_count} 个文件上传失败")
            
    def reset_to_default(self):
        """重置为默认字体库路径"""
        if messagebox.askyesno("确认重置", "确定要重置为默认字体库路径吗？"):
            default_path = "resources/fonts"
            self.font_base_path_var.set(default_path)
            self.log_callback("重置字体路径", "info", f"已重置为默认路径: {default_path}")
            
    def save_settings(self):
        """保存字体管理设置"""
        try:
            # 更新处理器的字体库路径
            new_path = self.font_base_path_var.get().strip()
            if not new_path:
                messagebox.showerror("错误", "字体库路径不能为空")
                return
                
            # 如果是相对路径，检查相对于当前工作目录是否存在
            if not os.path.isabs(new_path):
                abs_path = os.path.abspath(new_path)
            else:
                abs_path = new_path
                
            # 检查路径是否存在，如果不存在则创建
            if not os.path.exists(abs_path):
                if messagebox.askyesno("创建目录", f"字体库路径不存在，是否创建？\n\n{abs_path}"):
                    os.makedirs(abs_path, exist_ok=True)
                    # 创建子目录
                    os.makedirs(os.path.join(abs_path, "default"), exist_ok=True)
                    os.makedirs(os.path.join(abs_path, "zh"), exist_ok=True)
                else:
                    return
                    
            # 更新处理器配置（保持原始路径格式，相对或绝对）
            self.processor.font_base_path = new_path
            
            # 重新加载字体库
            self.processor.load_available_fonts()
            
            # 刷新主界面的字体列表
            if self.refresh_callback:
                self.refresh_callback()
                
            self.log_callback("保存字体设置", "success", f"字体库路径已更新: {new_path}")
            messagebox.showinfo("保存成功", "字体管理设置已保存，字体列表已刷新")
            
            # 关闭窗口
            self.window.destroy()
            
        except Exception as e:
            error_msg = f"保存字体设置失败: {str(e)}"
            self.log_callback("保存字体设置", "error", error_msg)
            messagebox.showerror("保存失败", error_msg)
            
    def cancel(self):
        """取消并关闭窗口"""
        self.window.destroy()