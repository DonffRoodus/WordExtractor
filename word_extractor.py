import os
import sys
import argparse
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import win32com.client
except ImportError:
    print("pywin32库未安装，正在尝试安装...")
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])
    import win32com.client


class WordExtractor:
    def __init__(self):
        self.word_app = None

    def extract_pages(self, input_file, output_file, start_page, end_page=None):
        """
        Extract specified page range from a Word document with full format preservation.
        :param input_file: Path to input file
        :param output_file: Path to output file
        :param start_page: Starting page number (1-based)
        :param end_page: Ending page number, defaults to start_page if None
        :return: True if successful, False otherwise
        """
        if end_page is None:
            end_page = start_page

        if start_page < 1 or end_page < start_page:
            print("错误：无效的页码范围")
            return False

        return self._extract_with_win32com(input_file, output_file, start_page, end_page)

    def _extract_with_win32com(self, input_file, output_file, start_page, end_page):
        """
        Extract pages using win32com with full format preservation.
        """
        print("使用win32com提取页面...")
        word_app = None
        try:
            # Initialize Word application
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = False
            word_app.DisplayAlerts = False

            # Open source document
            doc = word_app.Documents.Open(os.path.abspath(input_file))
            doc.Repaginate()  # Ensure accurate page calculation

            # Get total pages
            total_pages = doc.ComputeStatistics(2)  # 2 represents pages
            if total_pages <= 0:
                total_pages = doc.Windows(1).Panes(1).Pages.Count
            print(f"文档总页数: {total_pages}")

            if end_page > total_pages:
                print(f"警告：文档只有{total_pages}页，将提取到最后一页")
                end_page = total_pages

            # Create new document
            new_doc = word_app.Documents.Add()

            # Copy styles with dependency order
            self._copy_styles_with_dependencies(doc, new_doc)

            # Select and copy content
            doc.Activate()
            word_app.Selection.HomeKey(6)  # Move to document start
            word_app.Selection.GoTo(1, 1, start_page)  # Go to start page
            start_pos = word_app.Selection.Start

            if end_page < total_pages:
                word_app.Selection.GoTo(1, 1, end_page + 1)
                end_pos = word_app.Selection.Start - 1 if word_app.Selection.Start > 0 else word_app.Selection.EndKey(6)
            else:
                word_app.Selection.EndKey(6)  # Move to document end
                end_pos = word_app.Selection.Start

            if start_pos >= end_pos:
                print("错误：无法确定有效的选择范围")
                return False

            # Copy and paste range
            doc_range = doc.Range(start_pos, end_pos)
            doc_range.Copy()
            new_doc.Activate()
            new_doc.Range(0).Paste()

            print(f"成功提取从第 {start_page} 页到第 {end_page} 页的内容")

            # Remove trailing blank paragraphs
            self._remove_trailing_blank_paragraphs(new_doc)

            # Save new document
            file_ext = os.path.splitext(output_file)[1].lower()
            save_format = 0 if file_ext == '.doc' else 16  # 0=wdFormatDocument, 16=wdFormatXMLDocument
            new_doc.SaveAs(os.path.abspath(output_file), FileFormat=save_format)
            new_doc.Close(0)
            doc.Close(0)
            return True

        except Exception as e:
            print(f"使用win32com提取时出错: {str(e)}")
            return False
        finally:
            if word_app is not None:
                try:
                    word_app.Quit()
                except:
                    pass

    def _copy_styles_with_dependencies(self, source_doc, target_doc):
        """
        Copy styles from source to target document, respecting dependencies.
        """
        print("正在复制样式...")
        styles_to_copy = list(source_doc.Styles)
        style_dependencies = {}
        for style in styles_to_copy:
            try:
                style_name = style.NameLocal
                based_on = style.BasedOn.NameLocal if style.BasedOn else None
                style_dependencies[style_name] = based_on
            except Exception as e:
                print(f"获取样式 {style.NameLocal} 依赖关系时出错: {str(e)}")

        # Topological sort to ensure base styles are copied first
        copied_styles = set()
        sorted_styles = []

        while styles_to_copy:
            style = styles_to_copy.pop(0)
            style_name = style.NameLocal
            based_on = style_dependencies.get(style_name)

            if based_on is None or based_on in copied_styles:
                sorted_styles.append(style)
                copied_styles.add(style_name)
            else:
                styles_to_copy.append(style)

        # Copy sorted styles
        for style in sorted_styles:
            try:
                if style.BuiltIn:
                    target_style = target_doc.Styles(style.NameLocal)
                    target_style.ParagraphFormat = style.ParagraphFormat.Duplicate
                    target_style.Font = style.Font.Duplicate
                else:
                    if style.NameLocal not in [s.NameLocal for s in target_doc.Styles]:
                        new_style = target_doc.Styles.Add(style.NameLocal, style.Type)
                        new_style.ParagraphFormat = style.ParagraphFormat.Duplicate
                        new_style.Font = style.Font.Duplicate
                    else:
                        print(f"样式 '{style.NameLocal}' 已存在于目标文档，跳过添加")
            except Exception as e:
                print(f"复制样式 {style.NameLocal} 时出错: {str(e)}")

    def _remove_trailing_blank_paragraphs(self, doc):
        """
        Remove trailing blank paragraphs to prevent extra blank pages.
        """
        paragraphs = doc.Paragraphs
        last_non_empty_index = -1
        for i in range(paragraphs.Count, 0, -1):
            paragraph = paragraphs.Item(i)
            text = paragraph.Range.Text.strip()
            if text and text != '\r':
                last_non_empty_index = i
                break
        if last_non_empty_index != -1 and last_non_empty_index < paragraphs.Count:
            for i in range(paragraphs.Count, last_non_empty_index, -1):
                paragraphs.Item(i).Range.Delete()

    def __del__(self):
        if self.word_app is not None:
            try:
                self.word_app.Quit()
            except:
                pass


class WordExtractorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Word文档页面提取器")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        
        self.extractor = WordExtractor()
        self.setup_ui()
    
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
        file_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(file_frame, text="输入文件:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.input_file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.input_file_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="浏览...", command=self.browse_input_file).grid(row=0, column=2, padx=5, pady=5)
        
        ttk.Label(file_frame, text="输出文件:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.output_file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.output_file_var, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="浏览...", command=self.browse_output_file).grid(row=1, column=2, padx=5, pady=5)
        
        page_frame = ttk.LabelFrame(main_frame, text="页面范围", padding="10")
        page_frame.pack(fill=tk.X, pady=5)
        
        self.mode_var = tk.StringVar(value="single")
        ttk.Radiobutton(page_frame, text="单页模式", variable=self.mode_var, value="single", command=self.update_page_inputs).grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Radiobutton(page_frame, text="范围模式", variable=self.mode_var, value="range", command=self.update_page_inputs).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        self.page_frame_inner = ttk.Frame(page_frame)
        self.page_frame_inner.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=5)
        
        self.single_page_frame = ttk.Frame(self.page_frame_inner)
        ttk.Label(self.single_page_frame, text="页码:").pack(side=tk.LEFT, padx=5)
        self.single_page_var = tk.StringVar(value="1")
        ttk.Entry(self.single_page_frame, textvariable=self.single_page_var, width=10).pack(side=tk.LEFT, padx=5)
        
        self.range_page_frame = ttk.Frame(self.page_frame_inner)
        ttk.Label(self.range_page_frame, text="起始页:").pack(side=tk.LEFT, padx=5)
        self.start_page_var = tk.StringVar(value="1")
        ttk.Entry(self.range_page_frame, textvariable=self.start_page_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(self.range_page_frame, text="结束页:").pack(side=tk.LEFT, padx=5)
        self.end_page_var = tk.StringVar(value="1")
        ttk.Entry(self.range_page_frame, textvariable=self.end_page_var, width=10).pack(side=tk.LEFT, padx=5)
        
        self.single_page_frame.pack()
        
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(button_frame, text="提取页面", command=self.extract_pages).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="退出", command=self.root.quit).pack(side=tk.RIGHT, padx=5)
        
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        log_frame = ttk.LabelFrame(main_frame, text="日志", padding="5")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.log_text = tk.Text(log_frame, height=10, width=70, wrap=tk.WORD)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        self.log("程序已启动，请选择Word文档和页面范围进行提取。")
    
    def update_page_inputs(self):
        if self.mode_var.get() == "single":
            self.range_page_frame.pack_forget()
            self.single_page_frame.pack()
        else:
            self.single_page_frame.pack_forget()
            self.range_page_frame.pack()
    
    def browse_input_file(self):
        file_path = filedialog.askopenfilename(
            title="选择Word文档",
            filetypes=[("Word文档", "*.docx;*.doc"), ("所有文件", "*.*")]
        )
        if file_path:
            self.input_file_var.set(file_path)
            input_path = Path(file_path)
            output_path = input_path.parent / f"{input_path.stem}_提取{input_path.suffix}"
            self.output_file_var.set(str(output_path))
    
    def browse_output_file(self):
        file_path = filedialog.asksaveasfilename(
            title="保存提取的文档",
            filetypes=[("Word文档", "*.docx"), ("旧版Word文档", "*.doc")],
            defaultextension=".docx"
        )
        if file_path:
            self.output_file_var.set(file_path)
    
    def extract_pages(self):
        input_file = self.input_file_var.get().strip()
        output_file = self.output_file_var.get().strip()
        
        if not input_file:
            messagebox.showerror("错误", "请选择输入文件")
            return
        
        if not output_file:
            messagebox.showerror("错误", "请指定输出文件")
            return
        
        try:
            if self.mode_var.get() == "single":
                start_page = int(self.single_page_var.get())
                end_page = start_page
            else:
                start_page = int(self.start_page_var.get())
                end_page = int(self.end_page_var.get())
            
            if start_page < 1 or end_page < start_page:
                messagebox.showerror("错误", "无效的页码范围")
                return
        except ValueError:
            messagebox.showerror("错误", "页码必须是数字")
            return
        
        self.status_var.set("正在提取页面...")
        self.log(f"开始提取: 从 {input_file} 提取第 {start_page} 到 {end_page} 页到 {output_file}")
        self.root.update()
        
        success = self.extractor.extract_pages(input_file, output_file, start_page, end_page)
        
        if success:
            self.status_var.set("提取完成")
            self.log("页面提取成功完成！")
            messagebox.showinfo("成功", f"已成功提取页面并保存到:\n{output_file}")
        else:
            self.status_var.set("提取失败")
            self.log("页面提取失败，请查看详细错误信息。")
            messagebox.showerror("失败", "页面提取失败，请查看日志了解详情。")
    
    def log(self, message):
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)


def main_gui():
    root = tk.Tk()
    app = WordExtractorGUI(root)
    root.mainloop()


def main_cli():
    parser = argparse.ArgumentParser(description="从Word文档中提取指定页面")
    parser.add_argument("input_file", help="输入Word文档路径")
    parser.add_argument("output_file", help="输出Word文档路径")
    parser.add_argument("--start", type=int, required=True, help="起始页码")
    parser.add_argument("--end", type=int, help="结束页码（可选，默认与起始页相同）")
    
    args = parser.parse_args()
    
    extractor = WordExtractor()
    success = extractor.extract_pages(args.input_file, args.output_file, args.start, args.end)
    
    if success:
        print("页面提取成功完成！")
        return 0
    else:
        print("页面提取失败！")
        return 1


if __name__ == "__main__":
    if len(sys.argv) > 1:
        sys.exit(main_cli())
    else:
        main_gui()