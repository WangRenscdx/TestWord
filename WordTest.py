import random
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import messagebox, ttk, filedialog
import os
import sys
from fpdf import FPDF
from openpyxl import load_workbook, Workbook
from pathlib import Path


class VocabularyTestApp:
    def __init__(self, root):
        self.root = root
        self.root.title("单词测验程序")
        self.root.geometry("800x600")
        self.root.resizable(True, True)

        self.setup_fonts()
        self.mode = tk.StringVar(value="word_to_meaning")
        self.correct_count = 0
        self.total_count = 0
        self.wrong_answers = []
        self.test_questions = []
        self.question_index = 0
        self.selected_var = tk.StringVar()
        self.current_word = None

        # 优化文件路径检测
        self.desktop_path = Path.home() / "Desktop" / "单词本.xlsx"
        self.file_path = self.desktop_path if self.desktop_path.exists() else Path("单词本.xlsx")

        self.wb, self.sheet, self.words = self.load_words()
        self.create_ui()
        if self.words:
            self.show_welcome()

    def setup_fonts(self):
        """设置支持中文的字体"""
        if sys.platform.startswith('win'):
            self.default_font = ('Microsoft YaHei UI', 10)
            self.title_font = ('Microsoft YaHei UI', 16, 'bold')
            self.question_font = ('Microsoft YaHei UI', 14)
            self.button_font = ('Microsoft YaHei UI', 12)
            self.option_font = ('Microsoft YaHei UI', 16)  # 调整字体大小为16
        else:
            self.default_font = ('SimHei', 10)
            self.title_font = ('SimHei', 16, 'bold')
            self.question_font = ('SimHei', 14)
            self.button_font = ('SimHei', 12)
            self.option_font = ('SimHei', 16)  # 调整字体大小为16

    def load_words(self):
        """加载单词数据"""
        try:
            if not os.path.exists(self.file_path):
                return self.create_new_workbook()

            wb = load_workbook(self.file_path)
            sheet = wb.active
            words = []

            if sheet.max_row < 2:
                messagebox.showinfo("提示", "单词本为空，请添加单词后再开始测试。")
                return wb, sheet, []

            for row in sheet.iter_rows(min_row=2, values_only=True):
                if len(row) < 5:
                    continue
                word, meaning, last_correct_time, correct_ratio = row[0], row[2], row[3], row[4] or 0
                if word and meaning:
                    if last_correct_time and not isinstance(last_correct_time, datetime):
                        try:
                            last_correct_time = datetime.strptime(str(last_correct_time), '%Y-%m-%d %H:%M:%S')
                        except:
                            last_correct_time = None
                    words.append((word, meaning, last_correct_time, correct_ratio))
            return wb, sheet, words
        except Exception as e:
            messagebox.showerror("错误", f"加载单词本时出现错误: {e}")
            return self.create_new_workbook()

    def create_new_workbook(self):
        """创建新的单词本工作簿"""
        try:
            wb = Workbook()
            sheet = wb.active
            sheet.title = "单词表"
            sheet['A1'] = "单词"
            sheet['B1'] = "发音"
            sheet['C1'] = "释义"
            sheet['D1'] = "最后正确时间"
            sheet['E1'] = "正确率"

            wb.save(self.file_path)
            messagebox.showinfo("提示", f"已创建新的单词本: {self.file_path}")

            return wb, sheet, []
        except Exception as e:
            messagebox.showerror("错误", f"创建新单词本时出现错误: {e}")
            return None, None, []

    def create_ui(self):
        """创建用户界面"""
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 顶部框架 - 包含标题和进度
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))

        # 标题
        title_frame = ttk.Frame(top_frame)
        title_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Label(title_frame, text="单词测验程序", font=self.title_font).pack(anchor=tk.W)

        # 右上角进度显示
        progress_frame = ttk.Frame(top_frame)
        progress_frame.pack(side=tk.RIGHT)

        self.progress_label = ttk.Label(
            progress_frame,
            text="0/20",
            font=('Arial', 16, 'bold'),
            foreground="#007ACC"
        )
        self.progress_label.pack(padx=10, pady=5)

        # 模式选择
        mode_frame = ttk.LabelFrame(main_frame, text="测验模式", padding="10")
        mode_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Radiobutton(mode_frame, text="根据单词选释义", variable=self.mode, value="word_to_meaning",
                        command=self.reset_test).pack(anchor=tk.W)
        ttk.Radiobutton(mode_frame, text="根据释义选单词", variable=self.mode, value="meaning_to_word",
                        command=self.reset_test).pack(anchor=tk.W)

        # 问题区域
        self.question_frame = ttk.LabelFrame(main_frame, text="问题", padding="10")
        self.question_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self.word_label = ttk.Label(self.question_frame, text="", font=self.question_font, wraplength=700)
        self.word_label.pack(pady=(0, 10))

        self.meaning_label = ttk.Label(self.question_frame, text="", font=self.question_font, wraplength=700)
        self.meaning_label.pack(pady=(0, 10))

        # 选项区域
        self.options_frame = ttk.LabelFrame(main_frame, text="选项", padding="10")
        self.options_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 结果区域
        self.result_frame = ttk.LabelFrame(main_frame, text="结果", padding="10")
        self.result_frame.pack(fill=tk.X, pady=(0, 10))

        self.result_label = ttk.Label(self.result_frame, text="", font=self.question_font)
        self.result_label.pack()

        # 底部按钮区域
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=(10, 0))

        # 左侧按钮
        left_button_frame = ttk.Frame(bottom_frame)
        left_button_frame.pack(side=tk.LEFT)

        self.start_button = ttk.Button(left_button_frame, text="开始测试", command=self.start_test,
                                       style='Accent.TButton')
        self.start_button.pack(side=tk.LEFT, padx=(0, 10))

        self.check_button = ttk.Button(left_button_frame, text="检查答案", command=self.check_answer,
                                       state=tk.DISABLED)
        self.check_button.pack(side=tk.LEFT, padx=(0, 10))

        self.reset_button = ttk.Button(left_button_frame, text="重置测试", command=self.reset_test)
        self.reset_button.pack(side=tk.LEFT)

        # **新增：右下角导入按钮**
        right_button_frame = ttk.Frame(bottom_frame)
        right_button_frame.pack(side=tk.RIGHT)

        self.import_button = ttk.Button(
            right_button_frame,
            text="导入Excel",
            command=self.import_excel_file,
            style='Accent.TButton'
        )
        self.import_button.pack(side=tk.RIGHT)

        # 初始化进度显示
        self.update_progress()

    def show_welcome(self):
        """显示欢迎信息"""
        self.word_label.config(text="欢迎使用单词测验程序！")
        self.meaning_label.config(text=f"已加载 {len(self.words)} 个单词，请选择测验模式并点击开始测试。")
        self.result_label.config(text="")
        self.clear_options()
        self.update_progress()

    def clear_options(self):
        """清除选项区域"""
        for widget in self.options_frame.winfo_children():
            widget.destroy()

    def update_progress(self):
        """更新进度显示"""
        self.progress_label.config(text=f"{self.question_index}/20")

    def reset_test(self):
        """重置测试"""
        self.correct_count = 0
        self.total_count = 0
        self.wrong_answers = []
        self.test_questions = []
        self.question_index = 0
        self.current_word = None
        self.selected_var.set("")

        if not self.words:
            self.show_welcome()
            return

        if self.mode.get() == "word_to_meaning":
            self.word_label.config(text="")
            self.meaning_label.config(text="")
        else:
            self.word_label.config(text="")
            self.meaning_label.config(text="")

        self.clear_options()
        self.result_label.config(text="测试已重置，请点击开始测试。")
        self.update_progress()
        self.check_button.config(state=tk.DISABLED)

    def select_questions(self):
        """选择20个问题"""
        if not self.words:
            self.result_label.config(text="未成功加载单词，请检查文件。", foreground="red")
            return []

        eligible_words = []
        for word_info in self.words:
            word, meaning, last_correct_time, correct_ratio = word_info
            if correct_ratio < 0.99:
                if last_correct_time and (datetime.now() - last_correct_time) > timedelta(days=7):
                    eligible_words.extend([word_info] * 5)  # 权重提高
                else:
                    eligible_words.append(word_info)

        if not eligible_words:
            self.result_label.config(text="所有单词的正确率都达到了99%以上，无需测试！", foreground="green")
            return []

        # 处理单词量不足20的情况
        if len(eligible_words) < 20:
            result = eligible_words.copy()
            while len(result) < 20:
                result.append(random.choice(eligible_words))
            random.shuffle(result)
            return result

        return random.sample(eligible_words, 20)

    def start_test(self):
        """开始测试"""
        if not self.words:
            self.result_label.config(text="未成功加载单词，请检查文件。", foreground="red")
            return

        if self.question_index == 0:
            self.test_questions = self.select_questions()
            if not self.test_questions:
                return

        if self.question_index >= len(self.test_questions):
            self.result_label.config(text="测试已结束，请查看结果。", foreground="red")
            return

        self.current_word = self.test_questions[self.question_index]
        self.selected_var.set("")
        self.result_label.config(text="")

        if self.mode.get() == "word_to_meaning":
            self.word_label.config(text=f"单词：{self.current_word[0]}")
            self.meaning_label.config(text="")
            self.show_options(self.current_word[1], [word[1] for word in self.words if word[1]])
        else:
            self.word_label.config(text="")
            self.meaning_label.config(text=f"释义：{self.current_word[1]}")
            self.show_options(self.current_word[0], [word[0] for word in self.words if word[0]])

        self.check_button.config(state=tk.NORMAL)
        self.update_progress()

    def show_options(self, correct_answer, all_options):
        """显示选项"""
        self.clear_options()

        # 创建自定义样式并设置字体
        style = ttk.Style()
        style.configure('Option.TRadiobutton', font=self.option_font)

        options = [correct_answer]
        while len(options) < 6:
            random_option = random.choice(all_options)
            if random_option and random_option not in options:
                options.append(random_option)

        random.shuffle(options)

        for i, option in enumerate(options):
            option_frame = ttk.Frame(self.options_frame)
            option_frame.pack(fill=tk.X, pady=5)

            # 创建一个包含变量的唯一名称，用于每个单选按钮
            var_name = f"option_{i}"

            # 创建单选按钮
            radio_btn = ttk.Radiobutton(
                option_frame,
                variable=self.selected_var,
                value=option,
                style='Option.TRadiobutton'
            )
            radio_btn.pack(side=tk.LEFT, padx=(5, 10))

            # 创建标签来显示选项文本，并设置wraplength
            label = ttk.Label(
                option_frame,
                text=option,
                font=self.option_font,
                wraplength=550  # 设置适当的换行宽度
            )
            label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

    def check_answer(self):
        """检查答案"""
        if not self.current_word:
            return

        user_answer = self.selected_var.get()
        if not user_answer:
            self.result_label.config(text="请选择一个选项！", foreground="red")
            return

        self.total_count += 1
        is_correct = False

        if self.mode.get() == "word_to_meaning":
            if user_answer == self.current_word[1]:
                self.result_label.config(text="回答正确！", foreground="green")
                is_correct = True
            else:
                self.result_label.config(text=f"回答错误，正确答案是：{self.current_word[1]}", foreground="red")
                self.wrong_answers.append((self.current_word[0], self.current_word[1], user_answer))
        else:
            if user_answer == self.current_word[0]:
                self.result_label.config(text="回答正确！", foreground="green")
                is_correct = True
            else:
                self.result_label.config(text=f"回答错误，正确答案是：{self.current_word[0]}", foreground="red")
                self.wrong_answers.append((self.current_word[0], self.current_word[1], user_answer))

        if is_correct:
            self.correct_count += 1
            self.update_word_info(self.current_word, True)
        else:
            self.update_word_info(self.current_word, False)

        self.save_excel()

        self.question_index += 1
        self.check_button.config(state=tk.DISABLED)
        self.update_progress()

        if self.total_count >= 20:
            self.show_result()
        else:
            self.root.after(1000, self.start_test)

    def update_word_info(self, word_info, is_correct):
        """更新单词信息"""
        word, meaning, last_correct_time, correct_ratio = word_info

        for i, (w, m, lct, cr) in enumerate(self.words):
            if w == word and m == meaning:
                if is_correct:
                    new_correct_ratio = (correct_ratio * (self.total_count - 1) + 1) / self.total_count
                    new_last_correct_time = datetime.now()
                    self.words[i] = (w, m, new_last_correct_time, new_correct_ratio)
                    self.sheet.cell(row=i + 2, column=4, value=new_last_correct_time)
                    self.sheet.cell(row=i + 2, column=5, value=new_correct_ratio)
                else:
                    new_correct_ratio = correct_ratio * (self.total_count - 1) / self.total_count
                    self.words[i] = (w, m, last_correct_time, new_correct_ratio)
                    self.sheet.cell(row=i + 2, column=5, value=new_correct_ratio)
                break

    def save_excel(self):
        """保存Excel文件"""
        try:
            if self.wb:
                self.wb.save(self.file_path)
        except Exception as e:
            messagebox.showerror("错误", f"保存文件时出现错误: {e}")

    def show_result(self):
        """显示测试结果"""
        result_window = tk.Toplevel(self.root)
        result_window.title("测试结果")
        result_window.geometry("600x600")
        result_window.resizable(True, True)
        result_window.transient(self.root)
        result_window.grab_set()

        if sys.platform.startswith('win'):
            result_font = ('Microsoft YaHei UI', 12)
            title_font = ('Microsoft YaHei UI', 16, 'bold')
            button_font = ('Microsoft YaHei UI', 12)
        else:
            result_font = ('SimHei', 12)
            title_font = ('SimHei', 16, 'bold')
            button_font = ('SimHei', 12)

        ttk.Label(result_window, text="测试结果", font=title_font).pack(pady=20)

        score = self.correct_count / self.total_count * 100
        result_text = f"本次测试共 {self.total_count} 题，你答对了 {self.correct_count} 题，得分：{score:.2f} 分。"
        ttk.Label(result_window, text=result_text, font=result_font).pack(pady=10)

        chart_frame = ttk.Frame(result_window)
        chart_frame.pack(fill=tk.X, padx=50, pady=20)

        total_width = 400
        correct_width = int(total_width * (self.correct_count / self.total_count))
        wrong_width = total_width - correct_width

        chart = tk.Canvas(chart_frame, height=30, width=total_width, bg="white")
        chart.pack(fill=tk.X)

        chart.create_rectangle(0, 0, correct_width, 30, fill="green")
        chart.create_rectangle(correct_width, 0, total_width, 30, fill="red")

        ttk.Label(result_window, text=f"正确率: {self.correct_count}/{self.total_count} ({score:.2f}%)",
                  font=result_font).pack(pady=5)

        button_frame = ttk.Frame(result_window)
        button_frame.pack(fill=tk.X, padx=50, pady=20)

        if self.wrong_answers:
            ttk.Button(
                button_frame,
                text="查看错题解析",
                command=self.show_error_analysis,
                style='Accent.TButton'
            ).pack(side=tk.LEFT, padx=5, pady=10)

        ttk.Button(
            button_frame,
            text="重新开始",
            command=lambda: [result_window.destroy(), self.reset_test(), self.start_test()]
        ).pack(side=tk.LEFT, padx=5, pady=10)

        ttk.Button(
            button_frame,
            text="结束测试",
            command=lambda: [result_window.destroy(), self.reset_test()]
        ).pack(side=tk.LEFT, padx=5, pady=10)

    def show_error_analysis(self):
        """显示错误分析"""
        if not self.wrong_answers:
            messagebox.showinfo("提示", "没有错题！")
            return

        error_window = tk.Toplevel(self.root)
        error_window.title("错题解析")
        error_window.geometry("600x600")
        error_window.resizable(True, True)
        error_window.transient(self.root)
        error_window.grab_set()

        if sys.platform.startswith('win'):
            error_font = ('Microsoft YaHei UI', 12)
            title_font = ('Microsoft YaHei UI', 16, 'bold')
            button_font = ('Microsoft YaHei UI', 12)
        else:
            error_font = ('SimHei', 12)
            title_font = ('SimHei', 16, 'bold')
            button_font = ('SimHei', 12)

        ttk.Label(error_window, text="错题解析", font=title_font).pack(pady=20)

        frame = ttk.Frame(error_window)
        frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        text_widget = tk.Text(frame, wrap=tk.WORD, font=error_font)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(frame, command=text_widget.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        text_widget.config(yscrollcommand=scrollbar.set)

        error_text = "错误答案解释：\n\n"
        for i, (word, meaning, wrong_answer) in enumerate(self.wrong_answers, 1):
            if self.mode.get() == "word_to_meaning":
                error_text += f"{i}. 单词：{word}\n"
                error_text += f"   你的答案：{wrong_answer}\n"
                error_text += f"   正确答案：{meaning}\n\n"
            else:
                error_text += f"{i}. 释义：{meaning}\n"
                error_text += f"   你的答案：{wrong_answer}\n"
                error_text += f"   正确答案：{word}\n\n"

        text_widget.insert(tk.END, error_text)
        text_widget.config(state=tk.DISABLED)

        ttk.Button(
            error_window,
            text="导出为 PDF",
            command=lambda: self.export_to_pdf(error_text),
            style='Accent.TButton'
        ).pack(pady=10)

    def export_to_pdf(self, content):
        """导出为PDF"""
        try:
            pdf = FPDF()
            pdf.add_page()
            # 添加中文字体支持
            try:
                pdf.add_font('SimHei', '', 'SimHei.ttf', uni=True)
                pdf.set_font('SimHei', size=12)
            except:
                # 如果找不到字体文件，使用默认字体
                pdf.set_font("Arial", size=12)
                messagebox.showwarning("警告", "未找到中文字体文件，PDF中的中文可能无法正确显示。")

            lines = content.split("\n")
            for line in lines:
                pdf.multi_cell(0, 10, txt=line, align='L')

            pdf_filename = "单词测验错题解析.pdf"
            pdf.output(pdf_filename)

            messagebox.showinfo("导出成功", f"错题解析已成功导出为 {pdf_filename}")
        except Exception as e:
            messagebox.showerror("导出失败", f"导出 PDF 时出现错误：{e}")

    # **新增：导入Excel文件功能**
    def import_excel_file(self):
        """导入本地Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )

        if not file_path:
            return  # 用户取消选择

        try:
            # 尝试加载选中的文件
            wb = load_workbook(file_path)
            sheet = wb.active

            # 验证文件格式是否正确（检查表头）
            headers = [cell.value for cell in sheet[1]]
            required_headers = ["单词", "发音", "释义", "最后正确时间", "正确率"]

            if not all(header in headers for header in required_headers):
                messagebox.showerror("格式错误", "选择的Excel文件格式不正确，缺少必要的列。")
                return

            # 加载单词数据
            words = []
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if len(row) < 5:
                    continue
                word, meaning, last_correct_time, correct_ratio = row[0], row[2], row[3], row[4] or 0
                if word and meaning:
                    if last_correct_time and not isinstance(last_correct_time, datetime):
                        try:
                            last_correct_time = datetime.strptime(str(last_correct_time), '%Y-%m-%d %H:%M:%S')
                        except:
                            last_correct_time = None
                    words.append((word, meaning, last_correct_time, correct_ratio))

            # 更新应用状态
            self.file_path = Path(file_path)
            self.wb = wb
            self.sheet = sheet
            self.words = words

            # 重置测试并显示欢迎信息
            self.reset_test()
            self.show_welcome()

            messagebox.showinfo("成功", f"已成功导入 {len(words)} 个单词。")

        except Exception as e:
            messagebox.showerror("错误", f"导入文件时出现错误: {e}")


def main():
    root = tk.Tk()
    style = ttk.Style()
    style.configure('Accent.TButton', font=('Arial', 12, 'bold'))
    app = VocabularyTestApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()