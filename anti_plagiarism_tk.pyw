#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# nuitka-project: --mode=onefile
# nuitka-project: --enable-plugin=tk-inter
# nuitka-project: --windows-console-mode=disable
# nuitka-project: --assume-yes-for-downloads
"""
综合反抄袭检测工具 - 图形界面版本
使用 tkinter 实现易于使用的图形界面
"""

# Python标准库导入
import os
import platform
import shutil
import subprocess
import threading
from io import StringIO

# Tkinter模块导入
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# 自定义模块导入
from anti_plagiarism_tool import (
    analyze_document_metadata,
    detect_text_plagiarism,
    detect_image_plagiarism
)
from anti_plagiarism_enhanced import run_enhanced_detection

class AntiPlagiarismGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("综合反抄袭检测工具")
        self.root.geometry("1000x700")
        self.root.resizable(True, True)
        
        # 存储检测路径
        self.docx_directory = ""
        
        # 创建界面
        self.create_widgets()
        
    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="we")
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)
        
        # 标题
        title_label = ttk.Label(main_frame, text="综合反抄袭检测工具", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # 文件夹选择区域
        folder_label = ttk.Label(main_frame, text="选择文档文件夹:")
        folder_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 10))
        
        self.folder_var = tk.StringVar()
        folder_entry = ttk.Entry(main_frame, textvariable=self.folder_var, width=50)
        folder_entry.grid(row=1, column=1, sticky="we", padx=(0, 10))
        
        browse_button = ttk.Button(main_frame, text="浏览...", command=self.browse_folder)
        browse_button.grid(row=1, column=2)

        # 增强版检测的可选模板白名单；全部检测模块均会依次执行
        enhanced_frame = ttk.LabelFrame(main_frame, text="增强检测参数", padding="5")
        enhanced_frame.grid(row=2, column=0, columnspan=3, sticky="we", pady=(10, 0))
        enhanced_frame.columnconfigure(2, weight=1)

        ttk.Label(enhanced_frame, text="模板白名单（可选）:").grid(
            row=0, column=0, sticky=tk.W)
        self.template_var = tk.StringVar()
        template_entry = ttk.Entry(enhanced_frame, textvariable=self.template_var)
        template_entry.grid(row=0, column=1, columnspan=2, sticky="we",
                            padx=(5, 5))
        template_button = ttk.Button(
            enhanced_frame, text="选择文件...", command=self.browse_template)
        template_button.grid(row=0, column=3)
        
        # 操作按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=20)
        
        self.start_button = ttk.Button(button_frame, text="开始检测", command=self.start_detection)
        self.start_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.exit_button = ttk.Button(button_frame, text="退出", command=self.root.quit)
        self.exit_button.pack(side=tk.LEFT)
        
        # 输出区域
        output_label = ttk.Label(main_frame, text="检测结果:")
        output_label.grid(row=4, column=0, sticky=tk.W)
        
        self.output_text = scrolledtext.ScrolledText(main_frame, height=20, state=tk.DISABLED)
        self.output_text.grid(row=5, column=0, columnspan=3, sticky="wens", pady=(5, 0))
        
        # 配置输出文本框的标签
        self.output_text.tag_config("info", foreground="blue")
        self.output_text.tag_config("success", foreground="green")
        self.output_text.tag_config("error", foreground="red")
        self.output_text.tag_config("title", foreground="black", font=("Arial", 10, "bold"))
        
        # 配置网格权重
        main_frame.rowconfigure(5, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 操作按钮框架（检测完成后显示）
        self.action_frame = ttk.Frame(main_frame)
        self.action_frame.grid(row=6, column=0, columnspan=3, pady=(10, 0))
        self.action_frame.grid_remove()  # 初始隐藏
        
    def create_action_buttons(self):
        """创建操作按钮"""
        # 清除之前的按钮
        for widget in self.action_frame.winfo_children():
            widget.destroy()
        
        # 显示操作按钮框架
        self.action_frame.grid()
        
        button_groups = [
            ("基本检测结果", [
            ("打开图片临时文件夹", self.open_temp_folder),
            ("删除图片临时文件夹", self.delete_temp_folder),
            ("打开元信息CSV", self.open_metadata_csv),
            ("打开文字检测CSV", self.open_text_csv),
            ("打开图片检测CSV", self.open_image_csv)
            ]),
            ("增强检测结果", [
            ("打开增强模板CSV", self.open_enhanced_template_csv),
            ("打开增强精确匹配CSV", self.open_enhanced_exact_csv),
            ("打开增强指纹CSV", self.open_enhanced_fingerprint_csv),
            ("打开增强句级匹配CSV", self.open_enhanced_sentence_csv),
            ("打开增强rsid分析CSV", self.open_enhanced_rsid_csv)
            ])
        ]

        for row, (group_title, buttons) in enumerate(button_groups):
            group_frame = ttk.LabelFrame(
                self.action_frame, text=group_title, padding=(5, 2))
            group_frame.grid(row=row, column=0, sticky="w", pady=(0, 5))
            for column, (text, command) in enumerate(buttons):
                button = ttk.Button(group_frame, text=text, command=command)
                button.grid(row=0, column=column, padx=(0, 5))
        
    def open_temp_folder(self):
        """打开图片临时文件夹"""
        temp_folder = os.path.normpath('temp_word_images')
        self.open_path(temp_folder)
        
    def delete_temp_folder(self):
        """删除图片临时文件夹"""
        temp_folder = os.path.normpath('temp_word_images')
        if os.path.exists(temp_folder):
            try:
                import shutil
                shutil.rmtree(temp_folder)
                messagebox.showinfo("成功", "临时文件夹已删除")
            except Exception as e:
                messagebox.showerror("错误", f"删除失败: {str(e)}")
        else:
            messagebox.showwarning("警告", "临时文件夹不存在")
            
    def open_metadata_csv(self):
        """打开元信息CSV"""
        csv_file = os.path.normpath("plagiarism_metadata_analysis.csv")
        self.open_path(csv_file)
        
    def open_text_csv(self):
        """打开文字检测CSV"""
        csv_file = os.path.normpath("plagiarism_combined_results.csv")
        self.open_path(csv_file)
        
    def open_image_csv(self):
        """打开图片检查CSV"""
        csv_file = os.path.normpath("plagiarism_image_detection.csv")
        self.open_path(csv_file)

    def open_enhanced_template_csv(self):
        self.open_path("enhanced_template_paragraphs.csv")

    def open_enhanced_exact_csv(self):
        self.open_path("enhanced_exact_match.csv")

    def open_enhanced_fingerprint_csv(self):
        self.open_path("enhanced_fingerprint.csv")

    def open_enhanced_sentence_csv(self):
        self.open_path("enhanced_sentence_match.csv")

    def open_enhanced_rsid_csv(self):
        self.open_path("enhanced_rsid_analysis.csv")
        
    def open_path(self, path):
        """在文件管理器中打开路径或文件"""
        if not os.path.exists(path):
            messagebox.showwarning("警告", f"路径不存在: {path}")
            return
            
        try:
            system = platform.system()
            if system == "Windows":
                os.startfile(path)
            elif system == "Darwin":  # macOS
                subprocess.Popen(["open", path])
            else:  # Linux
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            messagebox.showerror("错误", f"无法打开路径: {str(e)}")
        
    def browse_folder(self):
        """浏览文件夹对话框"""
        folder_path = filedialog.askdirectory(title="选择包含Word文档的文件夹")
        if folder_path:
            self.folder_var.set(folder_path)

    def browse_template(self):
        """选择增强检测的模板白名单文件"""
        template_path = filedialog.askopenfilename(
            title="选择模板白名单文件",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        if template_path:
            self.template_var.set(template_path)
            
    def start_detection(self):
        """开始检测按钮事件处理"""
        folder_path = self.folder_var.get().strip()
        template_path = self.template_var.get().strip()
        
        if not folder_path:
            messagebox.showwarning("警告", "请先选择一个文件夹")
            return
            
        if not os.path.isdir(folder_path):
            messagebox.showerror("错误", f"文件夹不存在: {folder_path}")
            return

        if template_path and not os.path.isfile(template_path):
            messagebox.showerror("错误", f"模板白名单文件不存在: {template_path}")
            return
            
        # 保存检测路径
        self.docx_directory = folder_path
            
        # 禁用开始按钮
        self.start_button.config(state=tk.DISABLED)
        
        # 在新线程中运行检测，防止界面冻结
        detection_thread = threading.Thread(
            target=self.run_detection,
            args=(folder_path, template_path)
        )
        detection_thread.daemon = True
        detection_thread.start()
        
    def run_detection(self, folder_path, template_path=""):
        """在后台线程中运行检测"""
        # 重定向print输出到GUI
        import sys
        
        # 保存原始stdout
        old_stdout = sys.stdout
        
        try:
            # 创建StringIO对象捕获输出
            captured_output = StringIO()
            sys.stdout = captured_output
            
            # 执行检测
            self.log_message("开始执行综合反抄袭检测...\n", "title")
            
            # 基础检测优先执行
            # 1. 文档元数据分析
            self.log_message("[第一部分] 文档元数据分析\n", "title")
            analyze_document_metadata(folder_path)
            
            # 2. 文字抄袭检测
            self.log_message("\n[第二部分] 文字抄袭检测\n", "title")
            detect_text_plagiarism(folder_path)
            
            # 3. 图片抄袭检测
            self.log_message("\n[第三部分] 图片抄袭检测\n", "title")
            detect_image_plagiarism(folder_path)

            # 基础检测完成后执行增强文字检测
            self.log_message("\n[第四部分] 增强文字检测\n", "title")
            if not run_enhanced_detection(folder_path, template_path or None):
                print("增强文字检测未完成，请确认至少有两个有效的 docx 文档。")
            
            self.log_message("\n所有检测已完成！\n", "success")
            
            # 恢复原始stdout
            sys.stdout = old_stdout
            
            # 获取捕获的输出
            output = captured_output.getvalue()
            
            # 在GUI中显示输出
            self.display_output(output)
            
        except Exception as e:
            # 恢复原始stdout
            sys.stdout = old_stdout
            self.log_message(f"检测过程中发生错误: {str(e)}\n", "error")
        finally:
            # 恢复按钮状态和进度条
            self.root.after(0, self.detection_finished)
            
    def detection_finished(self):
        """检测完成后恢复界面状态"""
        self.start_button.config(state=tk.NORMAL)
        # 显示操作按钮
        self.create_action_buttons()
        
    def display_output(self, output):
        """在输出区域显示文本"""
        self.root.after(0, self._display_output, output)
        
    def _display_output(self, output):
        """在主线程中更新输出区域"""
        self.output_text.config(state=tk.NORMAL)
        self.output_text.delete(1.0, tk.END)
        self.output_text.insert(tk.END, output)
        self.output_text.config(state=tk.DISABLED)
        self.output_text.see(tk.END)
        
    def log_message(self, message, tag=None):
        """向输出区域添加消息"""
        self.root.after(0, self._log_message, message, tag)
        
    def _log_message(self, message, tag=None):
        """在主线程中添加消息到输出区域"""
        self.output_text.config(state=tk.NORMAL)
        if tag:
            self.output_text.insert(tk.END, message, tag)
        else:
            self.output_text.insert(tk.END, message)
        self.output_text.config(state=tk.DISABLED)
        self.output_text.see(tk.END)

def main():
    """主函数"""
    root = tk.Tk()
    app = AntiPlagiarismGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()