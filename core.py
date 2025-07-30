import pandas as pd
import fitz  # PyMuPDF
import os
import re
import json
import logging
from pathlib import Path
from datetime import datetime
from PIL import Image
from pptx import Presentation
from pptx.util import Inches
import CatchExcelImageTool


class ExcelToPDFProcessor:
    """Excel转PDF表单处理核心类"""
    
    def __init__(self):
        self.excel_path = ""
        self.pdf_template_path = ""
        self.output_folder = ""
        self.sheet_name = None
        self.title_row = 3
        self.start_row = 4
        self.filename_column = None  # 新增：文件名指定列
        self.field_mapping = {}
        self.col_separator = " "
        self.flatten_form = False
        self.output_png = False
        self.output_ppt = False
        
        # GUI日志回调函数
        self.gui_log_callback = None
        
        # 设置日志记录
        self.setup_logging()
        
    def excel_col_letter_to_index(self, col):
        """将Excel列字母如'A'转为列号索引，从0开始"""
        result = 0
        for c in col:
            if 'A' <= c <= 'Z':
                result = result * 26 + (ord(c) - ord('A') + 1)
            elif 'a' <= c <= 'z':
                result = result * 26 + (ord(c) - ord('a') + 1)
            else:
                return None
        return result - 1
    
    def setup_logging(self):
        """设置日志记录"""
        # 创建日志记录器
        self.logger = logging.getLogger('ExcelToPDFProcessor')
        self.logger.setLevel(logging.DEBUG)
        
        # 如果已经有处理器，先清除
        if self.logger.handlers:
            self.logger.handlers.clear()
        
        # 创建文件处理器
        log_file = 'app.log'
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        
        # 创建格式化器
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler.setFormatter(formatter)
        
        # 添加处理器到记录器
        self.logger.addHandler(file_handler)
        
        self.logger.info("日志系统初始化完成")
    
    def set_gui_log_callback(self, callback):
        """设置GUI日志回调函数"""
        self.gui_log_callback = callback
    
    def log_to_gui(self, operation, level="info", message=""):
        """发送日志到GUI（如果回调函数存在）"""
        if self.gui_log_callback:
            try:
                self.gui_log_callback(operation, level, message)
            except Exception as e:
                self.logger.error(f"GUI日志回调失败: {e}")

    def is_excel_col_pattern(self, val):
        """判断字符串是不是Excel列格式（如'A'或'A,B'）"""
        return re.fullmatch(r'([A-Za-z]+)(\s*,\s*[A-Za-z]+)*', val.strip()) is not None
    
    def extract_dispimg_id(self, cell_value):
        """从DISPIMG函数中提取图片ID"""
        if not cell_value or not isinstance(cell_value, str):
            return None
        
        # 匹配DISPIMG函数格式：=DISPIMG("ID_xxx",1) 或 =_xlfn.DISPIMG("ID_xxx",1)
        pattern = r'=(?:_xlfn\.)?DISPIMG\("([^"]+)"'
        match = re.search(pattern, str(cell_value))
        if match:
            return match.group(1)
        return None

    def get_pdf_form_keys(self, pdf_path):
        """获取PDF表单字段列表"""
        self.logger.info(f"开始获取PDF表单字段: {pdf_path}")
        
        if not os.path.exists(pdf_path):
            error_msg = f"PDF文件不存在: {pdf_path}"
            self.logger.error(error_msg)
            raise FileNotFoundError(error_msg)
            
        doc = fitz.open(pdf_path)
        fields = []
        
        try:
            for page_num in range(len(doc)):
                page = doc[page_num]
                widgets = page.widgets()
                
                for widget in widgets:
                    field_name = widget.field_name
                    if field_name and field_name not in fields:
                        fields.append(field_name)
            
            self.logger.info(f"成功获取到 {len(fields)} 个PDF表单字段: {fields}")
        except Exception as e:
            error_msg = f"检查表单字段时出错: {e}"
            self.logger.error(error_msg)
            raise Exception(error_msg)
        finally:
            doc.close()
        
        return fields
    
    def fill_pdf_image_field(self, page, widget, image_path):
        """在PDF表单域中填充图片"""
        try:
            # 获取字段的矩形区域
            field_rect = widget.rect
            
            # 插入图片到字段区域
            page.insert_image(field_rect, filename=image_path, overlay=True)
            
            self.logger.debug(f"成功在字段 '{widget.field_name}' 中插入图片: {image_path}")
            return True
        except Exception as e:
            self.logger.error(f"在字段 '{widget.field_name}' 中插入图片失败: {e}")
            return False

    def fill_pdf_form(self, input_pdf_path, output_pdf_path, data_dict, image_dict=None, flatten_form=False):
        """使用PyMuPDF填充PDF表单"""
        self.logger.debug(f"开始填充PDF表单: {output_pdf_path}")
        self.logger.debug(f"填充数据: {data_dict}")
        
        try:
            doc = fitz.open(input_pdf_path)
            
            filled_count = 0
            for page_num in range(len(doc)):
                page = doc[page_num]
                widgets = page.widgets()
                
                for widget in widgets:
                    field_name = widget.field_name
                    
                    # 处理图片字段
                    if image_dict and field_name in image_dict:
                        image_path = image_dict[field_name]
                        if image_path and os.path.exists(image_path):
                            if self.fill_pdf_image_field(page, widget, image_path):
                                filled_count += 1
                        continue
                    
                    # 处理普通文本字段
                    if field_name in data_dict:
                        value = str(data_dict[field_name])
                        try:
                            # 记录字段详细信息
                            field_type = widget.field_type
                            field_flags = widget.field_flags
                            self.logger.debug(f"字段 '{field_name}' 信息: 类型={field_type}, 标志={field_flags}")
                            
                            # 尝试多种方法设置字段值
                            widget.field_value = value
                            widget.update()
                            
                            # 对于文本字段，尝试特殊处理以保留空格
                            if field_type == fitz.PDF_WIDGET_TYPE_TEXT:
                                # 检查是否包含多个连续空格
                                if '  ' in value:  # 如果包含多个空格
                                    self.logger.info(f"检测到字段 '{field_name}' 包含多个空格，尝试保留")
                                    
                                    # 方法1: 使用非断行空格替换普通空格
                                    value_with_nbsp = value.replace(' ', '\u00A0')
                                    widget.field_value = value_with_nbsp
                                    widget.update()
                                    
                                    # 记录尝试结果
                                    result_nbsp = widget.field_value
                                    self.logger.debug(f"非断行空格方法结果: '{result_nbsp}'")
                                    
                                    # 如果非断行空格不工作，尝试其他方法
                                    if result_nbsp.replace('\u00A0', ' ') != value:
                                        # 方法2: 使用下划线临时替换空格（用户可以手动替换）
                                        value_with_underscore = value.replace('  ', '_')
                                        widget.field_value = value_with_underscore
                                        widget.update()
                                        
                                        result_underscore = widget.field_value
                                        self.logger.debug(f"下划线替换方法结果: '{result_underscore}'")
                                        
                                        if result_underscore == value_with_underscore:
                                            self.logger.warning(f"字段 '{field_name}' 使用下划线替换多个空格，请手动替换为空格")
                                        else:
                                            # 方法3: 恢复原始值
                                            widget.field_value = value
                                            widget.update()
                                else:
                                    # 普通情况，直接设置
                                    widget.field_value = value
                                    widget.update()
                            
                            # 验证设置是否成功
                            actual_value = widget.field_value
                            if actual_value != value:
                                self.logger.warning(f"字段 '{field_name}' 值被修改: 期望='{value}' (长度:{len(value)}), 实际='{actual_value}' (长度:{len(actual_value)})")
                                # 尝试使用字节方式重新设置
                                try:
                                    widget.field_value = value.encode('utf-8').decode('utf-8')
                                    widget.update()
                                    final_value = widget.field_value
                                    self.logger.debug(f"重新设置后的值: '{final_value}'")
                                except:
                                    pass
                            
                            filled_count += 1
                            self.logger.debug(f"成功填充字段 '{field_name}': 原始值='{value}', 最终值='{widget.field_value}'")
                        except Exception as e:
                            error_msg = f"填充字段 '{field_name}' 失败: {e}"
                            self.logger.error(error_msg)
                            print(error_msg)
            
            # 确保输出目录存在
            os.makedirs(os.path.dirname(output_pdf_path), exist_ok=True)
            
            # 保存文档
            if flatten_form:
                # 扁平化表单：将表单字段转换为静态文本
                self.logger.debug("正在扁平化表单...")
                pdfbytes = doc.convert_to_pdf()  # 这会扁平化所有表单字段
                flattened_doc = fitz.open("pdf", pdfbytes)
                flattened_doc.save(output_pdf_path, deflate=True, clean=True)
                flattened_doc.close()
                self.logger.debug("表单扁平化完成")
            else:
                doc.save(output_pdf_path, incremental=False, encryption=fitz.PDF_ENCRYPT_NONE)
            
            doc.close()
            success_msg = f"成功填充了 {filled_count} 个字段"
            self.logger.info(f"PDF填充完成: {output_pdf_path}, {success_msg}")
            return True, success_msg
            
        except Exception as e:
            error_msg = f"PDF填充失败: {e}"
            self.logger.error(f"填充PDF失败 {output_pdf_path}: {error_msg}")
            return False, error_msg

    def process_excel_to_pdf(self, progress_callback=None):
        """主要处理函数：将Excel数据填充到PDF表单"""
        self.logger.info("开始处理Excel转PDF任务")
        self.logger.info(f"Excel文件: {self.excel_path}")
        self.logger.info(f"PDF模板: {self.pdf_template_path}")
        self.logger.info(f"输出目录: {self.output_folder}")
        self.logger.info(f"配置参数: 表头行={self.title_row}, 数据起始行={self.start_row}, 扁平化={self.flatten_form}")
        
        try:
            # 验证输入
            if not os.path.exists(self.excel_path):
                error_msg = f"Excel文件不存在: {self.excel_path}"
                self.logger.error(error_msg)
                raise FileNotFoundError(error_msg)
            if not os.path.exists(self.pdf_template_path):
                error_msg = f"PDF模板不存在: {self.pdf_template_path}"
                self.logger.error(error_msg)
                raise FileNotFoundError(error_msg)
            if not self.field_mapping:
                error_msg = "字段映射不能为空"
                self.logger.error(error_msg)
                raise ValueError(error_msg)
            
            # 创建输出目录
            os.makedirs(self.output_folder, exist_ok=True)
            self.logger.info(f"输出目录已创建: {self.output_folder}")
            
            # 读取Excel数据
            # 处理sheet_name为空的情况
            sheet_to_read = self.sheet_name if self.sheet_name and self.sheet_name.strip() else 0
            self.logger.info(f"读取Excel sheet: {sheet_to_read if sheet_to_read != 0 else '第一个sheet'}")
            
            df = pd.read_excel(self.excel_path, sheet_name=sheet_to_read, header=None)
            
            # 如果返回的是字典（多个sheet），取第一个
            if isinstance(df, dict):
                sheet_names = list(df.keys())
                df = df[sheet_names[0]]
                self.logger.info(f"读取到多个sheet，使用第一个: {sheet_names[0]}")
            
            title_row_idx = self.title_row - 1
            data_start_idx = self.start_row - 1
            self.logger.info(f"Excel数据读取完成，总行数: {len(df)}")
            
            # 获取PDF表单字段
            pdf_keys = self.get_pdf_form_keys(self.pdf_template_path)
            
            # 检查字段映射
            missing_fields = []
            for pdf_key in self.field_mapping.keys():
                if pdf_key not in pdf_keys:
                    missing_fields.append(pdf_key)
            
            if missing_fields:
                warning_msg = f"警告：以下字段在PDF中不存在: {missing_fields}"
                self.logger.warning(warning_msg)
                print(warning_msg)
            
            # 处理每一行数据
            total_rows = len(df.iloc[data_start_idx:])
            success_count = 0
            error_messages = []
            generated_pdf_paths = []  # 存储成功生成的PDF文件路径
            self.logger.info(f"开始处理数据，共 {total_rows} 行")
            self.log_to_gui("开始数据处理", "info", f"共 {total_rows} 行数据待处理")
            
            for idx, row in df.iloc[data_start_idx:].iterrows():
                row_num = idx + 1
                actual_row_num = idx - data_start_idx + 1
                
                try:
                    self.logger.info(f"开始处理第 {row_num} 行数据 (数据行 {actual_row_num}/{total_rows})")
                    self.log_to_gui(f"处理第{row_num}行", "info", f"开始处理 (数据行 {actual_row_num}/{total_rows})")
                    
                    # 构建数据字典和图片字典
                    data = {}
                    image_data = {}
                    field_details = []  # 记录字段映射详情
                    
                    for pdf_key, map_config in self.field_mapping.items():
                        if isinstance(map_config, dict):
                            is_excel_col = map_config.get("is_excel_col", True)
                            is_excel_image = map_config.get("is_excel_image", False)
                            map_val = str(map_config.get("val", ""))
                            # 只有Excel列才需要strip，自定义值保留原始格式（包括空格）
                            if is_excel_col or is_excel_image:
                                map_val = map_val.strip()
                        else:
                            # 向后兼容旧格式
                            map_val = str(map_config)
                            # 判断是否为Excel列格式，如果是才strip
                            if self.is_excel_col_pattern(map_val.strip()):
                                map_val = map_val.strip()
                                is_excel_col = True
                                is_excel_image = False
                            else:
                                is_excel_col = False
                                is_excel_image = False
                        
                        if is_excel_image:
                            # 处理Excel图片列
                            col_idx = self.excel_col_letter_to_index(map_val.upper())
                            if col_idx is not None and col_idx < len(row):
                                cell_val = row[col_idx]
                                # 对于图片字段，即使单元格为空也要尝试提取浮动图片
                                should_extract_image = True
                                
                                if should_extract_image:
                                    # 创建临时图片目录
                                    temp_image_dir = os.path.join(self.output_folder, "temp_images")
                                    os.makedirs(temp_image_dir, exist_ok=True)
                                    
                                    # 首先尝试提取DISPIMG中的图片ID（仅当单元格不为空时）
                                    image_id = None
                                    image_path = None
                                    
                                    if not pd.isna(cell_val) and cell_val:
                                        image_id = self.extract_dispimg_id(str(cell_val))
                                    
                                    if image_id:
                                        # DISPIMG格式图片
                                        image_path = CatchExcelImageTool.extract_image_by_id(
                                            self.excel_path, image_id, temp_image_dir
                                        )
                                        if image_path:
                                            image_data[pdf_key] = image_path
                                            field_details.append(f"{pdf_key}={map_val}(DISPIMG-ID:{image_id})→图片已提取")
                                        else:
                                            field_details.append(f"{pdf_key}={map_val}(DISPIMG-ID:{image_id})→图片提取失败")
                                    
                                    # 如果没有DISPIMG或DISPIMG提取失败，尝试提取浮动图片
                                    if not image_path:
                                        # 尝试提取浮动图片
                                        # 将列索引转换为单元格地址
                                        from openpyxl.utils import get_column_letter
                                        cell_address = f"{get_column_letter(col_idx + 1)}{row_num}"
                                        
                                        # 获取工作表名称
                                        sheet_to_use = self.sheet_name if self.sheet_name and self.sheet_name.strip() else None
                                        if not sheet_to_use:
                                            # 如果没有指定工作表名，尝试获取第一个工作表名
                                            try:
                                                import openpyxl
                                                wb_temp = openpyxl.load_workbook(self.excel_path, read_only=True)
                                                sheet_to_use = wb_temp.sheetnames[0]
                                                wb_temp.close()
                                            except:
                                                sheet_to_use = "Sheet1"  # 默认工作表名
                                        
                                        self.logger.debug(f"尝试从单元格 {cell_address} (工作表: {sheet_to_use}) 提取浮动图片")
                                        
                                        # 先检查是否有浮动图片存在
                                        floating_images = CatchExcelImageTool._extract_floating_images(self.excel_path)
                                        self.logger.debug(f"Excel文件中发现的浮动图片: {floating_images}")
                                        
                                        image_path = CatchExcelImageTool.extract_image_from_cell(
                                            self.excel_path, sheet_to_use, cell_address, temp_image_dir
                                        )
                                        
                                        if image_path:
                                            image_data[pdf_key] = image_path
                                            field_details.append(f"{pdf_key}={map_val}(浮动图片-{cell_address})→图片已提取")
                                            self.logger.info(f"成功提取浮动图片: {image_path}")
                                        else:
                                            # 提供详细的诊断信息，但不使用备用方案
                                            if floating_images:
                                                self.logger.warning(f"发现{len(floating_images)}个浮动图片但位置与单元格{cell_address}不匹配")
                                                field_details.append(f"{pdf_key}={map_val}(单元格{cell_address})→发现{len(floating_images)}个浮动图片但位置不匹配")
                                            else:
                                                field_details.append(f"{pdf_key}={map_val}(单元格{cell_address})→未发现任何图片")
                                                self.logger.warning(f"单元格 {cell_address} 未发现任何图片")
                                # 如果最终没有提取到任何图片，记录相应信息
                                if not image_path:
                                    if pd.isna(cell_val) or not cell_val:
                                        field_details.append(f"{pdf_key}={map_val}→单元格为空且未找到浮动图片")
                                    else:
                                        field_details.append(f"{pdf_key}={map_val}→单元格有值但未找到图片")
                            else:
                                field_details.append(f"{pdf_key}={map_val}→列索引无效")
                        elif is_excel_col:
                            # 处理Excel列值
                            col_list = [col.strip().upper() for col in map_val.split(",")]
                            cell_values = []
                            for col in col_list:
                                col_idx = self.excel_col_letter_to_index(col)
                                if col_idx is not None and col_idx < len(row):
                                    val = row[col_idx]
                                    cell_val = "" if pd.isna(val) else str(val)
                                    cell_values.append(cell_val)
                                else:
                                    cell_values.append("")
                            final_value = self.col_separator.join(cell_values)
                            data[pdf_key] = final_value
                            field_details.append(f"{pdf_key}={map_val}→'{final_value}'")
                        else:
                            # 直接使用自定义值，保留空格
                            data[pdf_key] = map_val
                            field_details.append(f"{pdf_key}=自定义值→'{map_val}'")
                    
                    self.logger.debug(f"第 {row_num} 行字段映射详情: {'; '.join(field_details)}")
                    
                    # 生成输出文件名
                    if self.filename_column and self.filename_column.strip():
                        # 使用指定列的数据作为文件名
                        try:
                            filename_col_idx = self.excel_col_letter_to_index(self.filename_column.strip())
                            if filename_col_idx is not None and filename_col_idx < len(row):
                                filename_value = str(row.iloc[filename_col_idx]).strip()
                                # 清理文件名中的非法字符
                                filename_value = re.sub(r'[<>:"/\\|?*]', '_', filename_value)
                                if filename_value:
                                    output_filename = f"{filename_value}.pdf"
                                else:
                                    output_filename = f"{actual_row_num}.pdf"  # 去掉filled_前缀，只保留数字
                            else:
                                output_filename = f"{actual_row_num}.pdf"  # 去掉filled_前缀，只保留数字
                        except:
                            output_filename = f"{actual_row_num}.pdf"  # 去掉filled_前缀，只保留数字
                    else:
                        # 使用数字编号，去掉filled_前缀
                        output_filename = f"{actual_row_num}.pdf"
                    
                    output_pdf_path = os.path.join(self.output_folder, output_filename)
                    
                    # 填充PDF表单
                    success, message = self.fill_pdf_form(
                        self.pdf_template_path, output_pdf_path, data, image_data, self.flatten_form
                    )
                    
                    if success:
                        success_count += 1
                        generated_pdf_paths.append(output_pdf_path)  # 记录成功生成的PDF路径
                        self.logger.info(f"第 {row_num} 行处理成功 → {os.path.basename(output_pdf_path)}")
                        self.log_to_gui(f"第{row_num}行", "success", f"处理成功 → {os.path.basename(output_pdf_path)}")
                    else:
                        error_msg = f"第{row_num}行: {message}"
                        error_messages.append(error_msg)
                        self.logger.error(f"第 {row_num} 行处理失败: {message}")
                        self.log_to_gui(f"第{row_num}行", "error", f"处理失败: {message}")
                    
                    # 更新进度
                    if progress_callback:
                        progress = actual_row_num / total_rows * 100
                        progress_callback(progress, f"处理第 {row_num} 行")
                        
                except Exception as e:
                    error_msg = f"第{row_num}行处理失败: {e}"
                    error_messages.append(error_msg)
                    self.logger.error(f"第 {row_num} 行处理异常: {str(e)}")
                    self.log_to_gui(f"第{row_num}行", "error", f"处理异常: {str(e)}")
            
            # 记录处理结果
            self.logger.info(f"处理完成 - 总行数: {total_rows}, 成功: {success_count}, 失败: {len(error_messages)}")
            if error_messages:
                self.logger.warning(f"处理过程中的错误: {error_messages}")
            
            # 发送汇总日志到GUI
            if len(error_messages) == 0:
                self.log_to_gui("处理完成", "success", f"全部 {total_rows} 行数据处理成功")
            else:
                self.log_to_gui("处理完成", "warning", f"总行数: {total_rows}, 成功: {success_count}, 失败: {len(error_messages)}")
                # 发送前几个错误详情到GUI
                for i, error_msg in enumerate(error_messages[:3]):
                    self.log_to_gui("错误详情", "error", error_msg)
                if len(error_messages) > 3:
                    self.log_to_gui("错误详情", "warning", f"还有 {len(error_messages) - 3} 个错误未显示")
            
            # 处理PNG和PPT转换
            png_paths = []
            ppt_path = None
            
            if generated_pdf_paths:
                # 转换为PNG图片
                if self.output_png:
                    self.log_to_gui("PNG转换", "info", "开始转换PDF为PNG图片...")
                    for pdf_path in generated_pdf_paths:
                        png_path = self.convert_pdf_to_png(pdf_path, self.output_folder)
                        if png_path:
                            png_paths.append(png_path)
                    
                    if png_paths:
                        self.log_to_gui("PNG转换", "success", f"成功转换 {len(png_paths)} 个PNG图片")
                    else:
                        self.log_to_gui("PNG转换", "error", "PNG转换失败")
                
                # 创建PPT文件
                if self.output_ppt:
                    self.log_to_gui("PPT创建", "info", "开始创建PPT文件...")
                    ppt_path = self.create_ppt_from_pdfs(generated_pdf_paths, self.output_folder)
                    
                    if ppt_path:
                        self.log_to_gui("PPT创建", "success", "PPT文件创建成功")
                    else:
                        self.log_to_gui("PPT创建", "error", "PPT文件创建失败")
            
            return {
                "success": True,
                "total_rows": total_rows,
                "success_count": success_count,
                "error_count": len(error_messages),
                "error_messages": error_messages,
                "generated_pdf_paths": generated_pdf_paths,
                "png_paths": png_paths,
                "ppt_path": ppt_path
            }
            
        except Exception as e:
            error_msg = str(e)
            self.logger.error(f"处理过程中发生严重错误: {error_msg}")
            return {
                "success": False,
                "error": error_msg
            }

    def save_preset(self, preset_path):
        """保存预设配置到JSON文件"""
        self.logger.info(f"开始保存预设配置到: {preset_path}")
        
        preset_data = {
            "excel_path": self.excel_path,
            "pdf_template_path": self.pdf_template_path,
            "output_folder": self.output_folder,
            "sheet_name": self.sheet_name,
            "title_row": self.title_row,
            "start_row": self.start_row,
            "filename_column": self.filename_column,  # 新增：文件名列
            "field_mapping": self.field_mapping,
            "col_separator": self.col_separator,
            "flatten_form": self.flatten_form,
            "output_png": self.output_png,
            "output_ppt": self.output_ppt
        }
        
        try:
            with open(preset_path, 'w', encoding='utf-8') as f:
                json.dump(preset_data, f, ensure_ascii=False, indent=2)
            success_msg = "预设保存成功"
            self.logger.info(f"预设配置保存成功: {preset_path}")
            return True, success_msg
        except Exception as e:
            error_msg = f"预设保存失败: {e}"
            self.logger.error(f"保存预设配置失败 {preset_path}: {error_msg}")
            return False, error_msg

    def load_preset(self, preset_path):
        """从JSON文件加载预设配置"""
        self.logger.info(f"开始加载预设配置: {preset_path}")
        
        try:
            with open(preset_path, 'r', encoding='utf-8') as f:
                preset_data = json.load(f)
            
            self.excel_path = preset_data.get("excel_path", "")
            self.pdf_template_path = preset_data.get("pdf_template_path", "")
            self.output_folder = preset_data.get("output_folder", "")
            self.sheet_name = preset_data.get("sheet_name", None)
            self.title_row = preset_data.get("title_row", 3)
            self.start_row = preset_data.get("start_row", 4)
            self.filename_column = preset_data.get("filename_column", None)  # 新增：文件名列
            self.field_mapping = preset_data.get("field_mapping", {})
            self.col_separator = preset_data.get("col_separator", " ")
            self.flatten_form = preset_data.get("flatten_form", False)
            self.output_png = preset_data.get("output_png", False)
            self.output_ppt = preset_data.get("output_ppt", False)
            
            success_msg = "预设加载成功"
            self.logger.info(f"预设配置加载成功: {preset_path}")
            self.logger.debug(f"加载的配置: {preset_data}")
            return True, success_msg
        except Exception as e:
            error_msg = f"预设加载失败: {e}"
            self.logger.error(f"加载预设配置失败 {preset_path}: {error_msg}")
            return False, error_msg

    def reset_to_default(self):
        """恢复默认配置"""
        self.logger.info("开始恢复默认配置")
        
        self.excel_path = ""
        self.pdf_template_path = ""
        self.output_folder = str(Path.home() / "Desktop")
        self.sheet_name = None
        self.title_row = 3
        self.start_row = 4
        self.filename_column = None  # 新增：文件名列重置
        self.field_mapping = {}
        self.col_separator = " "
        self.flatten_form = False
        self.output_png = False
        self.output_ppt = False
        
        self.logger.info("默认配置恢复完成")
        return "已恢复默认配置"

    def get_desktop_path(self):
        """获取桌面路径"""
        return str(Path.home() / "Desktop")
    
    def convert_pdf_to_png(self, pdf_path, output_folder):
        """将PDF转换为PNG图片"""
        try:
            pdf_name = Path(pdf_path).stem
            png_path = os.path.join(output_folder, f"{pdf_name}.png")
            
            # 打开PDF文件
            pdf_doc = fitz.open(pdf_path)
            
            # 获取第一页（假设PDF只有一页）
            page = pdf_doc[0]
            
            # 设置缩放比例以获得高质量图片
            zoom = 2.0  # 缩放比例
            mat = fitz.Matrix(zoom, zoom)
            
            # 渲染页面为图片
            pix = page.get_pixmap(matrix=mat)
            
            # 保存为PNG
            pix.save(png_path)
            
            pdf_doc.close()
            
            self.logger.info(f"PDF转PNG成功: {png_path}")
            self.log_to_gui("PDF转PNG", "info", f"成功转换: {pdf_name}.png")
            
            return png_path
            
        except Exception as e:
            error_msg = f"PDF转PNG失败: {str(e)}"
            self.logger.error(error_msg)
            self.log_to_gui("PDF转PNG", "error", error_msg)
            return None
    
    def create_ppt_from_pdfs(self, pdf_paths, output_folder):
        """将多个PDF文件合并为一个PPT文件"""
        try:
            # 创建PPT演示文稿
            prs = Presentation()
            
            # 设置幻灯片尺寸为A4比例
            prs.slide_width = Inches(10)
            prs.slide_height = Inches(7.5)
            
            for pdf_path in pdf_paths:
                try:
                    # 打开PDF文件
                    pdf_doc = fitz.open(pdf_path)
                    
                    # 获取第一页
                    page = pdf_doc[0]
                    
                    # 设置缩放比例
                    zoom = 2.0
                    mat = fitz.Matrix(zoom, zoom)
                    
                    # 渲染页面为图片
                    pix = page.get_pixmap(matrix=mat)
                    
                    # 将图片数据转换为PIL Image
                    img_data = pix.tobytes("png")
                    
                    # 创建临时图片文件
                    temp_img_path = os.path.join(output_folder, f"temp_{Path(pdf_path).stem}.png")
                    with open(temp_img_path, "wb") as f:
                        f.write(img_data)
                    
                    # 添加新幻灯片
                    slide_layout = prs.slide_layouts[6]  # 空白布局
                    slide = prs.slides.add_slide(slide_layout)
                    
                    # 添加图片到幻灯片
                    left = Inches(0.5)
                    top = Inches(0.5)
                    width = Inches(9)
                    height = Inches(6.5)
                    
                    slide.shapes.add_picture(temp_img_path, left, top, width, height)
                    
                    # 删除临时图片文件
                    os.remove(temp_img_path)
                    
                    pdf_doc.close()
                    
                except Exception as e:
                    self.logger.warning(f"处理PDF文件 {pdf_path} 时出错: {str(e)}")
                    continue
            
            # 保存PPT文件
            ppt_path = os.path.join(output_folder, "merged_pdfs.pptx")
            prs.save(ppt_path)
            
            self.logger.info(f"PPT创建成功: {ppt_path}")
            self.log_to_gui("创建PPT", "info", f"成功创建PPT: merged_pdfs.pptx")
            
            return ppt_path
            
        except Exception as e:
            error_msg = f"创建PPT失败: {str(e)}"
            self.logger.error(error_msg)
            self.log_to_gui("创建PPT", "error", error_msg)
            return None