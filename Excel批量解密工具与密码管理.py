"""
Excel批量加密解密工具
版本: 3.0
功能: 批量加密/解密Excel文件，支持密码本，完全静默处理
作者: AI Assistant
日期: 2024-01-01
"""

import sys
import os
import csv
import traceback
from datetime import datetime
import pythoncom
import win32com.client
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *


class ProcessingThread(QThread):
    """处理线程 - 负责实际的Excel文件处理"""
    
    progress_signal = pyqtSignal(int, int)
    log_signal = pyqtSignal(str, str)
    file_status_signal = pyqtSignal(str, str, bool, str)
    finished_signal = pyqtSignal(dict)
    
    def __init__(self, input_path, output_path, file_list, is_encrypt, parent=None):
        super().__init__(parent)
        self.input_path = os.path.abspath(input_path)
        self.output_path = os.path.abspath(output_path)
        self.file_list = file_list
        self.is_encrypt = is_encrypt
        self.is_cancelled = False
    
    def run(self):
        """线程运行"""
        results = {
            'success_count': 0,
            'total_count': len(self.file_list),
            'failed_files': []
        }
        
        try:
            pythoncom.CoInitialize()
            
            # 初始化Excel - 完全静默模式
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AskToUpdateLinks = False
            excel.AlertBeforeOverwriting = False
            excel.AutomationSecurity = 1
            excel.EnableEvents = False
            excel.Interactive = False
            excel.ScreenUpdating = False
            
            # 清理临时文件
            self.clean_temp_files()
            
            # 设置进度
            self.progress_signal.emit(0, len(self.file_list))
            
            success_count = 0
            
            for i, file_info in enumerate(self.file_list):
                if self.is_cancelled:
                    self.log_signal.emit("处理已取消", "warning")
                    break
                    
                filename = file_info['filename']
                new_filename = file_info['new_filename']
                password = file_info['password']
                notes = file_info['notes']
                input_filepath = os.path.join(self.input_path, filename)
                output_filepath = os.path.join(self.output_path, new_filename)
                # print(file_info)
                # print(self.input_path,self.output_path)
                
                # 更新进度
                self.progress_signal.emit(i + 1, len(self.file_list))
                
                try:
                    if self.is_encrypt:
                        result = self.encrypt_file(excel, filename, input_filepath, output_filepath, password, notes)
                    else:
                        result = self.decrypt_file(excel, filename, input_filepath, output_filepath, password, notes)
                        
                    if result[0]:  # 成功
                        success_count += 1
                        self.file_status_signal.emit(filename, result[1], True, result[2])
                        self.log_signal.emit(f"{filename}: {result[1]} → {new_filename}", "success")
                    else:  # 失败
                        self.file_status_signal.emit(filename, result[1], False, result[2])
                        self.log_signal.emit(f"{filename}: {result[1]}", "error")
                        results['failed_files'].append({
                            'filename': filename,
                            'error': result[1]
                        })
                        
                except Exception as e:
                    error_msg = str(e)
                    self.file_status_signal.emit(filename, f"错误: {error_msg[:50]}", False, "处理异常")
                    self.log_signal.emit(f"{filename}: 处理失败 - {error_msg}", "error")
                    results['failed_files'].append({
                        'filename': filename,
                        'error': error_msg
                    })
                    
            # 关闭Excel
            try:
                excel.Quit()
                del excel
            except:
                pass
            
            # 更新结果
            results['success_count'] = success_count
            
            # 发送完成信号
            self.finished_processing(results)
            
        except Exception as e:
            self.log_signal.emit(f"线程错误: {str(e)}", "error")
            self.log_signal.emit(traceback.format_exc(), "error")
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    
    def clean_temp_files(self):
        """清理临时文件"""
        for f in os.listdir(self.input_path):
            if f.startswith('~$'):
                try:
                    os.remove(os.path.join(self.input_path, f))
                except:
                    pass
    
    def encrypt_file(self, excel, filename, input_filepath, output_filepath, password, notes):
        """加密文件"""
        try:
            if not password:
                return (False, "未设置密码", "密码为空")
                
            if not os.path.exists(input_filepath):
                return (False, "输入文件不存在", "文件路径错误")
            
            # 确保输出目录存在
            output_dir = os.path.dirname(output_filepath)
            if not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            
            wb = None
            try:
                # 打开原始工作簿
                wb = excel.Workbooks.Open(input_filepath)
                
                # 设置密码
                wb.Password = password
                wb.WritePassword = password
                
                # 另存为新文件
                wb.SaveAs(output_filepath)
                
                # 验证保存成功
                if os.path.exists(output_filepath):
                    return (True, "加密成功", notes)
                else:
                    return (False, "保存失败", "文件未创建")
                    
            except Exception as e:
                error_msg = str(e)
                if "password" in error_msg.lower() or "密码" in error_msg:
                    return (False, "密码设置错误", "密码设置失败")
                else:
                    return (False, f"加密失败: {error_msg[:50]}", "加密异常")
                    
            finally:
                if wb:
                    try:
                        wb.Close(SaveChanges=False)
                    except:
                        pass
                        
        except Exception as e:
            error_msg = str(e)
            return (False, f"加密失败: {error_msg[:50]}", "系统错误")
    
    def decrypt_file(self, excel, filename, input_filepath, output_filepath, password, notes):
        """解密文件 - 仅移除密码保护"""
        try:
            if not os.path.exists(input_filepath):
                return (False, "输入文件不存在", "文件路径错误")
            
            # 确保输出目录存在
            output_dir = os.path.dirname(output_filepath)
            if not os.path.exists(output_dir):
                os.makedirs(output_dir, exist_ok=True)
            
            wb = None
            try:
                # 尝试使用密码打开文件
                try:
                    if password:
                        wb = excel.Workbooks.Open(input_filepath, False, True, None, password)
                    else:
                        wb = excel.Workbooks.Open(input_filepath)
                except Exception as open_error:
                    error_msg = str(open_error)
                    if "password" in error_msg.lower() or "密码" in error_msg:
                        return (False, "密码错误或密码不匹配", "密码不正确")
                    else:
                        return (False, f"无法打开文件: {error_msg[:50]}", "打开失败")
                
                # 移除密码（仅此操作，不修改其他保护）
                try:
                    wb.Password = ""
                    wb.WritePassword = ""
                except:
                    pass  # 如果本来就没有密码，继续执行
                
                # 保存到指定位置
                try:
                    # 如果输出文件已存在，先删除
                    if os.path.exists(output_filepath):
                        try:
                            os.remove(output_filepath)
                        except:
                            pass
                    
                    # 保存解密后的文件
                    wb.SaveAs(output_filepath)
                    
                    # 验证保存成功
                    if os.path.exists(output_filepath):
                        return (True, "解密成功", notes)
                    else:
                        return (False, "保存失败", "文件未创建")
                        
                except Exception as save_error:
                    error_msg = str(save_error)
                    return (False, f"保存失败: {error_msg}", "保存错误")
                    
            finally:
                if wb:
                    try:
                        wb.Close(SaveChanges=False)
                    except:
                        pass
                        
        except Exception as e:
            error_msg = str(e)
            return (False, f"解密失败: {error_msg[:50]}", "系统错误")
    
    def finished_processing(self, results):
        """处理完成"""
        self.progress_signal.emit(len(self.file_list), len(self.file_list))
        self.finished_signal.emit(results)
    
    def cancel(self):
        """取消处理"""
        self.is_cancelled = True


class ExcelProtectorGUI(QMainWindow):
    """主界面类 - 负责用户交互"""
    
    def __init__(self):
        super().__init__()
        self.excel_app = None
        self.process_thread = None
        self.is_cancelled = False
        self.is_processing = False
        self.password_dict = {}
        self.init_ui()
    
    def init_ui(self):
        """初始化界面"""
        self.setWindowTitle("Excel批量加密解密工具 v3.0")
        self.setGeometry(100, 100, 1000, 800)
        
        try:
            self.setWindowIcon(QIcon.fromTheme("document-encrypt"))
        except:
            pass
        
        # 创建中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # 1. 文件夹设置区域
        folder_group = QGroupBox("文件夹设置")
        folder_layout = QVBoxLayout()
        
        # 输入文件夹
        input_folder_layout = QHBoxLayout()
        input_folder_layout.addWidget(QLabel("输入文件夹:"))
        self.input_folder_edit = QLineEdit(os.getcwd())
        self.input_folder_edit.setPlaceholderText("选择要处理的Excel文件所在文件夹...")
        input_folder_layout.addWidget(self.input_folder_edit)
        
        self.input_browse_button = QPushButton("浏览...")
        self.input_browse_button.clicked.connect(lambda: self.browse_folder(self.input_folder_edit))
        input_folder_layout.addWidget(self.input_browse_button)
        
        folder_layout.addLayout(input_folder_layout)
        
        # 输出文件夹
        output_folder_layout = QHBoxLayout()
        output_folder_layout.addWidget(QLabel("输出文件夹:"))
        self.output_folder_edit = QLineEdit(os.path.join(os.getcwd(), "output"))
        self.output_folder_edit.setPlaceholderText("处理后的文件保存位置...")
        output_folder_layout.addWidget(self.output_folder_edit)
        
        self.output_browse_button = QPushButton("浏览...")
        self.output_browse_button.clicked.connect(lambda: self.browse_folder(self.output_folder_edit))
        output_folder_layout.addWidget(self.output_browse_button)
        
        self.same_folder_check = QCheckBox("使用相同文件夹")
        self.same_folder_check.setChecked(False)
        self.same_folder_check.stateChanged.connect(self.on_same_folder_changed)
        output_folder_layout.addWidget(self.same_folder_check)
        
        folder_layout.addLayout(output_folder_layout)
        folder_group.setLayout(folder_layout)
        main_layout.addWidget(folder_group)
        
        # 2. 功能设置区域
        function_group = QGroupBox("功能设置")
        function_layout = QVBoxLayout()
        
        # 功能选择
        radio_layout = QHBoxLayout()
        self.encrypt_radio = QRadioButton("批量加密")
        self.decrypt_radio = QRadioButton("批量解密")
        self.encrypt_radio.setChecked(True)
        self.encrypt_radio.toggled.connect(self.on_function_changed)
        radio_layout.addWidget(self.encrypt_radio)
        radio_layout.addWidget(self.decrypt_radio)
        function_layout.addLayout(radio_layout)
        
        # 文件名后缀设置
        suffix_layout = QHBoxLayout()
        suffix_layout.addWidget(QLabel("文件名后缀:"))
        self.suffix_edit = QLineEdit()
        self.suffix_edit.setPlaceholderText("为空则不添加后缀")
        suffix_layout.addWidget(self.suffix_edit)
        
        suffix_layout.addWidget(QLabel("示例:"))
        self.suffix_example = QLabel("原文件.xlsx → 原文件_加密.xlsx")
        self.suffix_example.setStyleSheet("color: #666; font-style: italic;")
        suffix_layout.addWidget(self.suffix_example)
        
        function_layout.addLayout(suffix_layout)
        
        # 自动匹配密码本复选框
        self.auto_match_check = QCheckBox("自动从密码本匹配密码")
        self.auto_match_check.stateChanged.connect(self.on_auto_match_changed)
        function_layout.addWidget(self.auto_match_check)
        
        function_group.setLayout(function_layout)
        main_layout.addWidget(function_group)
        
        # 3. 密码设置区域
        password_group = QGroupBox("密码设置")
        password_layout = QVBoxLayout()
        
        # 密码本选择
        password_book_layout = QHBoxLayout()
        password_book_layout.addWidget(QLabel("密码本文件:"))
        self.password_book_edit = QLineEdit()
        self.password_book_edit.setPlaceholderText("选择包含文件-密码对应的CSV文件")
        password_book_layout.addWidget(self.password_book_edit)
        
        self.browse_password_button = QPushButton("浏览...")
        self.browse_password_button.clicked.connect(self.browse_password_book)
        password_book_layout.addWidget(self.browse_password_button)
        
        self.load_password_button = QPushButton("加载密码本")
        self.load_password_button.clicked.connect(self.load_password_book)
        password_book_layout.addWidget(self.load_password_button)
        
        # 导出模板按钮
        self.export_template_button = QPushButton("导出模板")
        self.export_template_button.clicked.connect(self.export_password_template)
        password_book_layout.addWidget(self.export_template_button)
        
        password_layout.addLayout(password_book_layout)
        
        # 单密码输入（用于加密）
        single_password_layout = QHBoxLayout()
        single_password_layout.addWidget(QLabel("统一密码:"))
        self.single_password_edit = QLineEdit()
        self.single_password_edit.setPlaceholderText("为所有文件设置相同的密码")
        self.single_password_edit.setEchoMode(QLineEdit.Password)
        self.single_password_edit.textChanged.connect(self.on_unified_password_changed)
        single_password_layout.addWidget(self.single_password_edit)
        
        self.show_password_check = QCheckBox("显示密码")
        self.show_password_check.stateChanged.connect(self.toggle_password_visibility)
        single_password_layout.addWidget(self.show_password_check)
        
        password_layout.addLayout(single_password_layout)
        
        # CSV格式说明
        csv_info = QLabel("CSV格式要求: 第一列为文件名, 第二列为密码")
        csv_info.setStyleSheet("color: #666; font-style: italic;")
        password_layout.addWidget(csv_info)
        
        password_group.setLayout(password_layout)
        main_layout.addWidget(password_group)
        
        # 4. 文件列表区域
        files_group = QGroupBox("文件列表")
        files_layout = QVBoxLayout()
        
        self.files_table = QTableWidget()
        self.files_table.setColumnCount(6)
        self.files_table.setHorizontalHeaderLabels(["选择", "文件名", "状态", "密码", "备注", "新文件名"])
        self.files_table.horizontalHeader().setStretchLastSection(True)
        
        # 设置列宽
        self.files_table.setColumnWidth(0, 50)    # 选择
        self.files_table.setColumnWidth(1, 200)   # 文件名
        self.files_table.setColumnWidth(2, 100)   # 状态
        self.files_table.setColumnWidth(3, 120)   # 密码
        self.files_table.setColumnWidth(4, 100)   # 备注
        self.files_table.setColumnWidth(5, 200)   # 新文件名
        
        files_layout.addWidget(self.files_table)
        
        # 按钮区域
        table_buttons_layout = QHBoxLayout()
        self.refresh_button = QPushButton("刷新文件列表")
        self.refresh_button.clicked.connect(self.scan_files)
        table_buttons_layout.addWidget(self.refresh_button)
        
        self.select_all_check = QCheckBox("全选")
        self.select_all_check.stateChanged.connect(self.select_all_files)
        table_buttons_layout.addWidget(self.select_all_check)
        
        self.update_passwords_button = QPushButton("更新选中文件的密码")
        self.update_passwords_button.clicked.connect(self.update_selected_passwords)
        table_buttons_layout.addWidget(self.update_passwords_button)
        
        self.preview_button = QPushButton("预览新文件名")
        self.preview_button.clicked.connect(self.preview_new_filenames)
        table_buttons_layout.addWidget(self.preview_button)
        
        table_buttons_layout.addStretch()
        files_layout.addLayout(table_buttons_layout)
        
        files_group.setLayout(files_layout)
        main_layout.addWidget(files_group)
        
        # 5. 进度和日志区域
        progress_group = QGroupBox("进度和日志")
        progress_layout = QVBoxLayout()
        
        self.progress_bar = QProgressBar()
        progress_layout.addWidget(self.progress_bar)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        progress_layout.addWidget(self.log_text)
        
        progress_group.setLayout(progress_layout)
        main_layout.addWidget(progress_group)
        
        # 6. 按钮区域
        button_layout = QHBoxLayout()
        
        self.start_button = QPushButton("开始执行")
        self.start_button.clicked.connect(self.start_processing)
        self.start_button.setStyleSheet("background-color: #4CAF50; color: white; padding: 8px;")
        
        self.cancel_button = QPushButton("取消")
        self.cancel_button.clicked.connect(self.cancel_processing)
        self.cancel_button.setEnabled(False)
        
        self.clear_log_button = QPushButton("清空日志")
        self.clear_log_button.clicked.connect(self.clear_log)
        
        self.export_log_button = QPushButton("导出日志")
        self.export_log_button.clicked.connect(self.export_log)
        
        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.cancel_button)
        button_layout.addWidget(self.clear_log_button)
        button_layout.addWidget(self.export_log_button)
        button_layout.addStretch()
        
        main_layout.addLayout(button_layout)
        
        # 状态栏
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪")
        
        # 初始化示例后缀
        self.on_function_changed()
        
        # 初始化文件列表
        QTimer.singleShot(100, self.scan_files)
    
    # ==================== 界面操作方法 ====================
    
    def browse_folder(self, line_edit):
        """浏览文件夹"""
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹", line_edit.text())
        if folder:
            line_edit.setText(folder)
            if line_edit == self.input_folder_edit:
                self.scan_files()
    
    def on_same_folder_changed(self, state):
        """使用相同文件夹复选框状态改变"""
        if state == Qt.Checked:
            self.output_folder_edit.setText(self.input_folder_edit.text())
            self.output_folder_edit.setEnabled(False)
            self.output_browse_button.setEnabled(False)
        else:
            self.output_folder_edit.setEnabled(True)
            self.output_browse_button.setEnabled(True)
    
    def on_function_changed(self):
        """功能选择改变时更新示例"""
        if self.encrypt_radio.isChecked():
            self.suffix_example.setText("原文件.xlsx → 原文件_加密.xlsx")
            self.suffix_edit.setPlaceholderText("如：_加密（为空则不添加后缀）")
        else:
            self.suffix_example.setText("原文件.xlsx → 原文件_解密.xlsx")
            self.suffix_edit.setPlaceholderText("如：_解密（为空则不添加后缀）")
    
    def on_auto_match_changed(self, state):
        """自动匹配复选框状态改变"""
        if state == Qt.Checked and self.password_dict:
            match_count = self.match_passwords_from_book()
            if match_count > 0:
                self.log_message(f"自动匹配完成，成功匹配 {match_count} 个文件", "success")
    
    def browse_password_book(self):
        """浏览密码本文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "选择密码本文件", 
            "", 
            "CSV文件 (*.csv);;所有文件 (*.*)"
        )
        if file_path:
            self.password_book_edit.setText(file_path)
    
    def load_password_book(self):
        """加载密码本并自动匹配"""
        password_book_path = self.password_book_edit.text()
        if not password_book_path or not os.path.exists(password_book_path):
            QMessageBox.warning(self, "警告", "请先选择有效的密码本文件！")
            return
            
        try:
            self.password_dict = {}
            row_count = 0
            
            # 支持的编码列表
            encodings = ['utf-8-sig', 'utf-8', 'gbk', 'gb2312', 'latin1', 'cp936']
            
            for encoding in encodings:
                try:
                    with open(password_book_path, 'r', encoding=encoding, newline='') as f:
                        # 读取第一行判断分隔符
                        first_line = f.readline()
                        f.seek(0)
                        
                        delimiter = '\t' if '\t' in first_line else ','
                        reader = csv.reader(f, delimiter=delimiter)
                        
                        for row in reader:
                            # 跳过空行和注释行
                            if not row or (len(row) > 0 and str(row[0]).strip().startswith('#')):
                                continue
                                
                            if len(row) >= 2:
                                filename = str(row[0]).strip()
                                password = str(row[1]).strip()
                                
                                if filename:
                                    self.password_dict[filename] = password
                                    row_count += 1
                        
                    self.log_message(f"使用编码 {encoding} 读取成功，共 {row_count} 条记录", "success")
                    break
                        
                except UnicodeDecodeError:
                    continue
                except Exception as e:
                    self.log_message(f"编码 {encoding} 读取失败: {str(e)}", "warning")
                    continue
            else:
                raise Exception("无法解码CSV文件，请检查文件编码或格式")
                    
            # 自动匹配密码
            if self.auto_match_check.isChecked():
                match_count = self.match_passwords_from_book()
                self.log_message(f"自动匹配完成，成功匹配 {match_count} 个文件", "success")
                
        except Exception as e:
            self.log_message(f"加载密码本失败: {str(e)}", "error")
            QMessageBox.critical(self, "错误", 
                f"加载密码本失败:\n{str(e)}\n\n请确保CSV文件格式正确：\n1. 使用UTF-8或GBK编码\n2. 第一列是文件名，第二列是密码\n3. 使用逗号或制表符分隔")
    
    def match_passwords_from_book(self):
        """从密码本自动匹配密码"""
        match_count = 0
        for row in range(self.files_table.rowCount()):
            filename_item = self.files_table.item(row, 1)
            if filename_item:
                filename = filename_item.text()
                if filename in self.password_dict:
                    password = self.password_dict[filename]
                    password_item = QTableWidgetItem(password)
                    self.files_table.setItem(row, 3, password_item)
                    
                    notes_item = QTableWidgetItem("来自密码本")
                    notes_item.setForeground(QColor("blue"))
                    self.files_table.setItem(row, 4, notes_item)
                    
                    match_count += 1
        
        return match_count
    
    def export_password_template(self):
        """导出密码本模板"""
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "导出密码本模板",
            f"excel_passwords_{datetime.now().strftime('%Y%m%d')}.csv",
            "CSV文件 (*.csv);;所有文件 (*.*)"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8-sig', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(['文件名', '密码', '备注'])
                    writer.writerow(['文件1.xlsx', 'password123', '示例1'])
                    writer.writerow(['文件2.xlsx', 'abc@2024', '示例2'])
                    writer.writerow(['财务表.xlsx', 'YS2026-XM', '重要文件'])
                    writer.writerow(['# 说明：', '', ''])
                    writer.writerow(['# 1. 第一列：Excel文件名（包含扩展名）', '', ''])
                    writer.writerow(['# 2. 第二列：对应的密码', '', ''])
                    writer.writerow(['# 3. 以#开头的行会被忽略', '', ''])
                    writer.writerow(['# 4. 可以使用逗号或制表符分隔', '', ''])
                
                self.log_message(f"密码本模板已导出到: {file_path}", "success")
                QMessageBox.information(self, "成功", f"密码本模板已导出到:\n{file_path}")
                
            except Exception as e:
                self.log_message(f"导出模板失败: {str(e)}", "error")
                QMessageBox.critical(self, "错误", f"导出模板失败:\n{str(e)}")
    
    def on_unified_password_changed(self, text):
        """统一密码改变时更新选中文件的密码"""
        if not text or not self.single_password_edit.isEnabled():
            return
            
        if self.auto_match_check.isChecked():
            return
            
        updated_count = 0
        for row in range(self.files_table.rowCount()):
            widget = self.files_table.cellWidget(row, 0)
            if widget:
                checkbox = widget.layout().itemAt(0).widget()
                if checkbox.isChecked():
                    password_item = QTableWidgetItem(text)
                    self.files_table.setItem(row, 3, password_item)
                    
                    notes_item = QTableWidgetItem("统一密码")
                    notes_item.setForeground(QColor("green"))
                    self.files_table.setItem(row, 4, notes_item)
                    updated_count += 1
        
        if updated_count > 0:
            self.log_message(f"已为 {updated_count} 个选中文件更新统一密码", "info")
    
    def toggle_password_visibility(self, state):
        """切换密码显示状态"""
        if state == Qt.Checked:
            self.single_password_edit.setEchoMode(QLineEdit.Normal)
        else:
            self.single_password_edit.setEchoMode(QLineEdit.Password)
    
    def scan_files(self):
        """扫描文件夹中的Excel文件"""
        folder_path = self.input_folder_edit.text()
        if not os.path.exists(folder_path):
            QMessageBox.warning(self, "警告", "输入文件夹不存在！")
            return
        
        # 确保输出文件夹存在
        output_path = self.output_folder_edit.text()
        if not os.path.exists(output_path):
            try:
                os.makedirs(output_path, exist_ok=True)
                self.log_message(f"创建输出文件夹: {output_path}", "info")
            except Exception as e:
                self.log_message(f"创建输出文件夹失败: {str(e)}", "error")
        
        # 清空表格
        self.files_table.setRowCount(0)
        
        # 扫描Excel文件
        excel_files = []
        for filename in os.listdir(folder_path):
            if filename.lower().endswith(('.xlsx', '.xls', '.xlsm', '.xlsb')):
                if not filename.startswith('~$') and not filename.startswith('~'):
                    excel_files.append(filename)
        
        if not excel_files:
            self.log_message(f"警告：在文件夹 {folder_path} 中未找到Excel文件", "warning")
            QMessageBox.warning(self, "提示", 
                f"在文件夹中未找到Excel文件：\n{folder_path}\n\n请检查：\n1. 文件夹路径是否正确\n2. 文件夹中是否有Excel文件（.xlsx, .xls）")
        
        # 更新表格
        for filename in excel_files:
            row = self.files_table.rowCount()
            self.files_table.insertRow(row)
            
            # 选择复选框
            checkbox = QCheckBox()
            checkbox.setChecked(True)
            checkbox_widget = QWidget()
            checkbox_layout = QHBoxLayout(checkbox_widget)
            checkbox_layout.addWidget(checkbox)
            checkbox_layout.setAlignment(Qt.AlignCenter)
            checkbox_layout.setContentsMargins(0, 0, 0, 0)
            self.files_table.setCellWidget(row, 0, checkbox_widget)
            
            # 文件名
            filename_item = QTableWidgetItem(filename)
            self.files_table.setItem(row, 1, filename_item)
            
            # 状态
            status_item = QTableWidgetItem("待处理")
            status_item.setForeground(QColor("gray"))
            self.files_table.setItem(row, 2, status_item)
            
            # 密码（初始为空）
            password_item = QTableWidgetItem("")
            self.files_table.setItem(row, 3, password_item)
            
            # 备注
            notes_item = QTableWidgetItem("")
            self.files_table.setItem(row, 4, notes_item)
            
            # 新文件名
            new_filename_item = QTableWidgetItem("")
            self.files_table.setItem(row, 5, new_filename_item)
        
        self.status_bar.showMessage(f"找到 {len(excel_files)} 个Excel文件")
        self.log_message(f"扫描完成，找到 {len(excel_files)} 个Excel文件", "success")
        
        # 如果密码本已加载且启用了自动匹配，进行匹配
        if self.password_dict and self.auto_match_check.isChecked():
            match_count = self.match_passwords_from_book()
            if match_count > 0:
                self.log_message(f"自动匹配完成，成功匹配 {match_count} 个文件", "success")
        
        # 预览新文件名
        self.preview_new_filenames()
    
    def select_all_files(self, state):
        """全选/取消全选文件"""
        for row in range(self.files_table.rowCount()):
            widget = self.files_table.cellWidget(row, 0)
            if widget:
                checkbox = widget.layout().itemAt(0).widget()
                checkbox.setChecked(state == Qt.Checked)
    
    def update_selected_passwords(self):
        """更新选中文件的密码"""
        password, ok = QInputDialog.getText(
            self, 
            "设置密码", 
            "请输入密码:", 
            QLineEdit.Password
        )
        
        if ok and password:
            updated_count = 0
            for row in range(self.files_table.rowCount()):
                widget = self.files_table.cellWidget(row, 0)
                if widget:
                    checkbox = widget.layout().itemAt(0).widget()
                    if checkbox.isChecked():
                        status_item = self.files_table.item(row, 2)
                        if status_item and status_item.text() != "待处理":
                            reply = QMessageBox.question(
                                self,
                                "确认覆盖",
                                f"文件 {self.files_table.item(row, 1).text()} 已经处理过，是否重新设置密码？",
                                QMessageBox.Yes | QMessageBox.No
                            )
                            if reply != QMessageBox.Yes:
                                continue
                        
                        password_item = QTableWidgetItem(password)
                        self.files_table.setItem(row, 3, password_item)
                        
                        notes_item = QTableWidgetItem("手动设置")
                        notes_item.setForeground(QColor("orange"))
                        self.files_table.setItem(row, 4, notes_item)
                        updated_count += 1
            
            if updated_count > 0:
                self.log_message(f"已为 {updated_count} 个选中文件设置密码: {password}", "success")
            else:
                self.log_message("没有选中任何文件", "warning")
    
    def preview_new_filenames(self):
        """预览新文件名"""
        suffix = self.suffix_edit.text().strip()
        
        for row in range(self.files_table.rowCount()):
            filename_item = self.files_table.item(row, 1)
            if filename_item:
                filename = filename_item.text()
                name, ext = os.path.splitext(filename)
                
                if suffix:
                    new_filename = f"{name}{suffix}{ext}"
                else:
                    new_filename = filename
                    
                new_filename_item = QTableWidgetItem(new_filename)
                self.files_table.setItem(row, 5, new_filename_item)
    
    def get_selected_files(self):
        """获取选中的文件及其信息"""
        selected_files = []
        for row in range(self.files_table.rowCount()):
            widget = self.files_table.cellWidget(row, 0)
            if widget:
                checkbox = widget.layout().itemAt(0).widget()
                if checkbox.isChecked():
                    filename = self.files_table.item(row, 1).text()
                    password_item = self.files_table.item(row, 3)
                    password = password_item.text() if password_item else ""
                    notes_item = self.files_table.item(row, 4)
                    notes = notes_item.text() if notes_item else ""
                    new_filename_item = self.files_table.item(row, 5)
                    new_filename = new_filename_item.text() if new_filename_item else filename
                    
                    selected_files.append({
                        'filename': filename,
                        'new_filename': new_filename,
                        'password': password,
                        'notes': notes
                    })
        return selected_files
    
    def validate_selection(self):
        """验证选择的有效性"""
        selected_files = self.get_selected_files()
        if not selected_files:
            return False, "请至少选择一个文件！"
        
        # 检查输出文件夹
        output_path = self.output_folder_edit.text()
        if not output_path:
            return False, "请设置输出文件夹！"
        
        try:
            os.makedirs(output_path, exist_ok=True)
        except Exception as e:
            return False, f"无法创建输出文件夹: {str(e)}"
        
        # 检查加密模式下的密码
        if self.encrypt_radio.isChecked():
            empty_password_files = []
            for file_info in selected_files:
                if not file_info['password']:
                    empty_password_files.append(file_info['filename'])
            
            if empty_password_files:
                file_list = "\n".join(empty_password_files[:5])
                if len(empty_password_files) > 5:
                    file_list += f"\n...等 {len(empty_password_files)} 个文件"
                return False, f"以下文件没有设置密码：\n{file_list}"
        
        return True, ""
    
    def start_processing(self):
        """开始处理"""
        if self.is_processing:
            QMessageBox.warning(self, "警告", "处理正在进行中，请等待完成或取消！")
            return
        
        # 验证输入
        input_path = self.input_folder_edit.text()
        print(input_path)
        if not os.path.exists(input_path):
            QMessageBox.warning(self, "警告", "请选择有效的输入文件夹！")
            return
        
        # 验证选择
        is_valid, error_msg = self.validate_selection()
        if not is_valid:
            QMessageBox.warning(self, "警告", error_msg)
            return
        
        # 获取选中的文件
        selected_files = self.get_selected_files()
        output_path = self.output_folder_edit.text()
        suffix = self.suffix_edit.text().strip()
        
        if self.encrypt_radio.isChecked():
            action = "加密"
        else:
            action = "解密"
        
        # 确认对话框
        confirm_msg = f"确定要开始批量{action}吗？\n\n"
        confirm_msg += f"处理文件数: {len(selected_files)}\n"
        confirm_msg += f"输入文件夹: {input_path}\n"
        confirm_msg += f"输出文件夹: {output_path}\n"
        
        if suffix:
            confirm_msg += f"文件名后缀: {suffix}\n"
        
        confirm_msg += "\n重要：原文件不会被修改，处理后的文件将保存到输出文件夹！"
        
        reply = QMessageBox.question(
            self,
            "确认开始",
            confirm_msg,
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # 禁用界面控件
        self.set_controls_enabled(False)
        self.is_cancelled = False
        self.is_processing = True
        
        # 创建处理线程
        self.process_thread = ProcessingThread(
            input_path,
            output_path,
            selected_files,
            self.encrypt_radio.isChecked(),
            self
        )
        self.process_thread.progress_signal.connect(self.update_progress)
        self.process_thread.log_signal.connect(self.log_message)
        self.process_thread.file_status_signal.connect(self.update_file_status)
        self.process_thread.finished_signal.connect(self.processing_finished)
        
        self.process_thread.start()
    
    def cancel_processing(self):
        """取消处理"""
        if not self.is_processing:
            return
        
        self.is_cancelled = True
        if self.process_thread and self.process_thread.isRunning():
            self.process_thread.cancel()
        self.log_message("正在取消处理...", "warning")
    
    def set_controls_enabled(self, enabled):
        """设置控件启用状态"""
        self.is_processing = not enabled
        
        self.input_folder_edit.setEnabled(enabled)
        self.input_browse_button.setEnabled(enabled)
        self.output_folder_edit.setEnabled(enabled and not self.same_folder_check.isChecked())
        self.output_browse_button.setEnabled(enabled and not self.same_folder_check.isChecked())
        self.same_folder_check.setEnabled(enabled)
        self.encrypt_radio.setEnabled(enabled)
        self.decrypt_radio.setEnabled(enabled)
        self.suffix_edit.setEnabled(enabled)
        self.auto_match_check.setEnabled(enabled)
        self.password_book_edit.setEnabled(enabled)
        self.browse_password_button.setEnabled(enabled)
        self.load_password_button.setEnabled(enabled)
        self.export_template_button.setEnabled(enabled)
        self.single_password_edit.setEnabled(enabled)
        self.show_password_check.setEnabled(enabled)
        self.refresh_button.setEnabled(enabled)
        self.select_all_check.setEnabled(enabled)
        self.update_passwords_button.setEnabled(enabled)
        self.preview_button.setEnabled(enabled)
        self.start_button.setEnabled(enabled)
        self.clear_log_button.setEnabled(enabled)
        self.export_log_button.setEnabled(enabled)
        self.cancel_button.setEnabled(not enabled)
        
        # 禁用/启用表格编辑
        for row in range(self.files_table.rowCount()):
            password_item = self.files_table.item(row, 3)
            if password_item:
                if enabled:
                    status_item = self.files_table.item(row, 2)
                    if status_item and status_item.text() == "待处理":
                        password_item.setFlags(password_item.flags() | Qt.ItemIsEditable)
                    else:
                        password_item.setFlags(password_item.flags() & ~Qt.ItemIsEditable)
                else:
                    password_item.setFlags(password_item.flags() & ~Qt.ItemIsEditable)
    
    def update_progress(self, value, maximum=None):
        """更新进度条"""
        if maximum is not None:
            self.progress_bar.setMaximum(maximum)
        self.progress_bar.setValue(value)
    
    def log_message(self, message, message_type="info"):
        """添加日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        
        if message_type == "error":
            color = "#ff4444"
            prefix = "[错误]"
        elif message_type == "warning":
            color = "#ff9900"
            prefix = "[警告]"
        elif message_type == "success":
            color = "#44aa44"
            prefix = "[成功]"
        else:
            color = "#666666"
            prefix = "[信息]"
        
        log_entry = f'<font color="#888888">{timestamp}</font> <font color="{color}">{prefix}</font> {message}'
        self.log_text.append(log_entry)
        
        # 自动滚动到底部
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def update_file_status(self, filename, status, is_success=True, notes=""):
        """更新文件状态"""
        for row in range(self.files_table.rowCount()):
            if self.files_table.item(row, 1).text() == filename:
                # 更新状态
                status_item = QTableWidgetItem(status)
                if is_success:
                    status_item.setForeground(QColor("#44aa44"))
                else:
                    status_item.setForeground(QColor("#ff4444"))
                self.files_table.setItem(row, 2, status_item)
                
                # 更新备注
                if notes:
                    notes_item = QTableWidgetItem(notes)
                    self.files_table.setItem(row, 4, notes_item)
                
                # 处理完成后禁用密码编辑
                password_item = self.files_table.item(row, 3)
                if password_item:
                    password_item.setFlags(password_item.flags() & ~Qt.ItemIsEditable)
                
                break
    
    def processing_finished(self, results):
        """处理完成"""
        success_count = results.get('success_count', 0)
        total_count = results.get('total_count', 0)
        
        self.set_controls_enabled(True)
        
        # 显示完成消息
        self.log_message(f"处理完成！成功：{success_count}/{total_count}", "success")
        self.status_bar.showMessage(f"处理完成 - 成功：{success_count}/{total_count}")
        
        # 显示完成对话框
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("完成")
        msg_box.setText(f"处理完成！\n总文件数：{total_count}\n成功处理：{success_count}\n失败：{total_count - success_count}")
        
        # 如果有失败文件，显示详情
        failed_files = results.get('failed_files', [])
        if failed_files:
            msg_box.setDetailedText("\n".join([f"{f['filename']}: {f['error']}" for f in failed_files]))
        
        msg_box.exec_()
    
    def clear_log(self):
        """清空日志"""
        self.log_text.clear()
    
    def export_log(self):
        """导出日志"""
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "导出日志",
            f"excel_tool_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
            "文本文件 (*.txt);;所有文件 (*.*)"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.toPlainText())
                self.log_message(f"日志已导出到：{file_path}", "success")
            except Exception as e:
                self.log_message(f"导出日志失败：{str(e)}", "error")
    
    def closeEvent(self, event):
        """关闭窗口事件"""
        if self.is_processing:
            reply = QMessageBox.question(
                self,
                "确认退出",
                "处理仍在进行中，确定要退出吗？",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.cancel_processing()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()


def main():
    """主函数"""
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    try:
        app.setWindowIcon(QIcon.fromTheme("applications-office"))
    except:
        pass
    
    window = ExcelProtectorGUI()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()