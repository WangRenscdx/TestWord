import random
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import messagebox, ttk, filedialog
import os
import sys
from fpdf import FPDF
from openpyxl import load_workbook, Workbook
from pathlib import Path
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]


class VocabularyTestApp:
    def __init__(self, root):
        self.root = root
        self.root.title("智能单词测验系统")
        # 设置窗口大小和位置
        self.screen_width = root.winfo_screenwidth()
        self.screen_height = root.winfo_screenheight()
        self.root.geometry(
            f"{int(self.screen_width * 0.8)}x{int(self.screen_height * 0.8)}+{int(self.screen_width * 0.1)}+{int(self.screen_height * 0.1)}")
        self.root.resizable(True, True)

        # 主题颜色配置 - 修改了强调色为更醒目的橙色
        self.primary_color = "#007ACC"
        self.secondary_color = "#223E5F"
        self.accent_color = "#FF6B35"  # 更醒目的橙色
        self.light_color = "#F5F7FA"
        self.dark_color = "#2D3B45"

        # 字体配置参数
        self.title_font_size = 18
        self.question_font_size = 20
        self.option_font_size = 16
        self.result_font_size = 14
        self.button_font_size = 12

        self.setup_fonts()
        self.mode = tk.StringVar(value="word_to_meaning")
        self.correct_count = 0
        self.total_count = 0
        self.wrong_answers = []
        self.test_questions = []
        self.question_index = 0
        self.selected_var = tk.StringVar()
        self.current_word = None

        # 生词本管理相关变量
        self.word_to_add = tk.StringVar()
        self.pos_to_add = tk.StringVar()  # 新增：词性
        self.meaning_to_add = tk.StringVar()
        self.search_term = tk.StringVar()
        self.current_words_display = []
        self.current_page = 1
        self.words_per_page = 10

        # 历史记录相关变量
        self.history_records = []
        self.history_displayed = False

        # 统计相关变量
        self.stats_displayed = False

        # 计时器相关变量
        self.test_started = False
        self.start_time = None
        self.time_limit = timedelta(minutes=20)
        self.time_remaining = self.time_limit
        self.timer_id = None

        # 文件路径
        self.desktop_path = Path.home() / "Desktop" / "单词本.xlsx"
        self.file_path = self.desktop_path if self.desktop_path.exists() else Path("单词本.xlsx")
        self.history_path = Path("单词测验历史记录.xlsx")

        # 加载单词和历史记录
        self.wb, self.sheet, self.words = self.load_words()
        self.load_history()

        # 创建主UI
        self.create_main_ui()
        if self.words:
            self.show_welcome()
        else:
            self.show_empty_state()

        # 绑定窗口大小变化事件
        self.root.bind("<Configure>", self.on_window_resize)

    def setup_fonts(self):
        """设置支持中文的字体"""
        if sys.platform.startswith('win'):
            self.default_font = ('Microsoft YaHei UI', 10)
            self.title_font = ('Microsoft YaHei UI', self.title_font_size, 'bold')
            self.question_font = ('Microsoft YaHei UI', self.question_font_size)
            self.button_font = ('Microsoft YaHei UI', self.button_font_size)
            self.option_font = ('Microsoft YaHei UI', self.option_font_size)
            self.result_font = ('Microsoft YaHei UI', self.result_font_size)
        else:
            self.default_font = ('SimHei', 10)
            self.title_font = ('SimHei', self.title_font_size, 'bold')
            self.question_font = ('SimHei', self.question_font_size)
            self.button_font = ('SimHei', self.button_font_size)
            self.option_font = ('SimHei', self.option_font_size)
            self.result_font = ('SimHei', self.result_font_size)

    def create_main_ui(self):
        """创建主用户界面"""
        # 创建主框架
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        # 创建顶部导航栏
        self.create_navigation_bar()

        # 创建内容区域
        self.content_frame = ttk.Frame(self.main_frame)
        self.content_frame.pack(fill=tk.BOTH, expand=True, pady=(20, 0))

        # 初始化测试内容区域
        self.create_test_content()

    def create_navigation_bar(self):
        """创建顶部导航栏"""
        nav_frame = ttk.Frame(self.main_frame, height=50)
        nav_frame.pack(fill=tk.X)

        # 配置导航按钮样式
        style = ttk.Style()
        style.configure('Nav.TButton', font=self.button_font, padding=10)
        style.configure('ActiveNav.TButton', font=self.button_font, padding=10, foreground=self.primary_color)

        # 导航按钮
        self.test_button = ttk.Button(nav_frame, text="测验", command=self.show_test_content, style='ActiveNav.TButton')
        self.test_button.pack(side=tk.LEFT, padx=(0, 5))

        self.dictionary_button = ttk.Button(nav_frame, text="生词本", command=self.show_dictionary_content,
                                            style='Nav.TButton')
        self.dictionary_button.pack(side=tk.LEFT, padx=(0, 5))

        self.history_button = ttk.Button(nav_frame, text="历史记录", command=self.show_history_content,
                                         style='Nav.TButton')
        self.history_button.pack(side=tk.LEFT, padx=(0, 5))

        self.stats_button = ttk.Button(nav_frame, text="学习统计", command=self.show_stats_content, style='Nav.TButton')
        self.stats_button.pack(side=tk.LEFT, padx=(0, 5))

        # 右侧系统信息
        self.system_info_frame = ttk.Frame(nav_frame)
        self.system_info_frame.pack(side=tk.RIGHT)

        self.word_count_label = ttk.Label(self.system_info_frame, text=f"单词总数: {len(self.words)}",
                                          font=self.default_font)
        self.word_count_label.pack(side=tk.LEFT, padx=(0, 20))

    def create_test_content(self):
        """创建测验内容区域"""
        # 顶部框架 - 包含标题、进度和计时器
        top_frame = ttk.Frame(self.content_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))

        # 标题
        title_frame = ttk.Frame(top_frame)
        title_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Label(title_frame, text="单词测验", font=self.title_font).pack(anchor=tk.W)

        # 右上角进度和计时器显示
        progress_timer_frame = ttk.Frame(top_frame)
        progress_timer_frame.pack(side=tk.RIGHT)

        self.progress_label = ttk.Label(
            progress_timer_frame,
            text="0/20",
            font=('Arial', 16, 'bold'),
            foreground=self.primary_color
        )
        self.progress_label.pack(side=tk.LEFT, padx=(0, 20))

        # 计时器显示
        self.timer_label = ttk.Label(
            progress_timer_frame,
            text="20:00",
            font=('Arial', 16, 'bold'),
            foreground="green"
        )
        self.timer_label.pack(side=tk.LEFT, padx=(0, 10))

        # 模式选择
        mode_frame = ttk.LabelFrame(self.content_frame, text="测验模式", padding="10")
        mode_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Radiobutton(mode_frame, text="根据单词选释义", variable=self.mode, value="word_to_meaning",
                        command=self.reset_test).pack(anchor=tk.W)
        ttk.Radiobutton(mode_frame, text="根据释义选单词", variable=self.mode, value="meaning_to_word",
                        command=self.reset_test).pack(anchor=tk.W)

        # 问题区域
        self.question_frame = ttk.LabelFrame(self.content_frame, text="问题", padding="10")
        self.question_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self.word_label = ttk.Label(self.question_frame, text="", font=self.question_font, wraplength=700)
        self.word_label.pack(pady=(0, 10))

        self.pos_label = ttk.Label(self.question_frame, text="", font=self.question_font, wraplength=700)  # 新增：词性标签
        self.pos_label.pack(pady=(0, 10))

        self.meaning_label = ttk.Label(self.question_frame, text="", font=self.question_font, wraplength=700)
        self.meaning_label.pack(pady=(0, 10))

        # 选项区域
        self.options_frame = ttk.LabelFrame(self.content_frame, text="选项", padding="10")
        self.options_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 结果区域
        self.result_frame = ttk.Frame(self.content_frame)
        self.result_frame.pack(fill=tk.X, pady=(0, 10))

        # 结果标签
        result_label_frame = ttk.LabelFrame(self.result_frame, text="结果", padding="10")
        result_label_frame.pack(fill=tk.X)

        self.result_label = ttk.Label(result_label_frame, text="", font=self.question_font)
        self.result_label.pack()

        # 底部按钮区域
        bottom_frame = ttk.Frame(self.content_frame)
        bottom_frame.pack(fill=tk.X, pady=(10, 0), side=tk.BOTTOM, anchor=tk.S)

        # 左侧按钮
        left_button_frame = ttk.Frame(bottom_frame)
        left_button_frame.pack(side=tk.LEFT)

        # 配置醒目的开始测试按钮样式
        style = ttk.Style()
        style.configure('Accent.TButton',
                        font=self.button_font,
                        foreground='white',
                        background=self.accent_color,  # 使用更醒目的橙色
                        padding=10,
                        borderwidth=1,
                        relief=tk.RAISED)
        style.map('Accent.TButton',
                  foreground=[('pressed', 'white'), ('active', 'white')],
                  background=[('pressed', '!disabled', '#E65C00'), ('active', '#FF8C66')])  # 更深的橙色变体

        self.start_button = ttk.Button(left_button_frame, text="开始测试", command=self.start_test,
                                       style='Accent.TButton')
        self.start_button.pack(side=tk.LEFT, padx=(0, 10))

        self.reset_button = ttk.Button(left_button_frame, text="重置测试", command=self.reset_test)
        self.reset_button.pack(side=tk.LEFT)

        # 右下角导入按钮
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

    def create_dictionary_content(self):
        """创建生词本内容区域"""
        self.dictionary_frame = ttk.Frame(self.content_frame)
        self.dictionary_frame.pack(fill=tk.BOTH, expand=True)

        # 顶部搜索和添加区域
        top_frame = ttk.Frame(self.dictionary_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))

        # 搜索区域
        search_frame = ttk.LabelFrame(top_frame, text="搜索", padding="10")
        search_frame.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        search_entry = ttk.Entry(search_frame, textvariable=self.search_term, width=30)
        search_entry.pack(side=tk.LEFT, padx=(0, 10))

        search_button = ttk.Button(search_frame, text="搜索", command=self.search_words)
        search_button.pack(side=tk.LEFT)

        # 添加单词区域
        add_frame = ttk.LabelFrame(top_frame, text="添加新单词", padding="10")
        add_frame.pack(side=tk.RIGHT, fill=tk.X, expand=True)

        word_entry_frame = ttk.Frame(add_frame)
        word_entry_frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(word_entry_frame, text="单词:", font=self.default_font).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(word_entry_frame, textvariable=self.word_to_add, width=15).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(word_entry_frame, text="词性:", font=self.default_font).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(word_entry_frame, textvariable=self.pos_to_add, width=10).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(word_entry_frame, text="释义:", font=self.default_font).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(word_entry_frame, textvariable=self.meaning_to_add, width=15).pack(side=tk.LEFT)

        add_button = ttk.Button(add_frame, text="添加", command=self.add_word)
        add_button.pack(side=tk.RIGHT, pady=(0, 5))

        # 单词列表区域
        list_frame = ttk.LabelFrame(self.dictionary_frame, text="单词列表", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 创建表格 - 增加词性列
        columns = ("单词", "词性", "释义")
        self.word_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=15)

        # 设置列宽和标题
        self.word_tree.heading("单词", text="单词")
        self.word_tree.column("单词", width=200, anchor=tk.CENTER)

        self.word_tree.heading("词性", text="词性")
        self.word_tree.column("词性", width=100, anchor=tk.CENTER)

        self.word_tree.heading("释义", text="释义")
        self.word_tree.column("释义", width=300, anchor=tk.CENTER)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.word_tree.yview)
        self.word_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.word_tree.pack(fill=tk.BOTH, expand=True)

        # 绑定双击事件以删除单词
        self.word_tree.bind("<Double-1>", self.on_double_click_word)

        # 分页控制
        pagination_frame = ttk.Frame(self.dictionary_frame)
        pagination_frame.pack(fill=tk.X, pady=(0, 10))

        self.prev_button = ttk.Button(pagination_frame, text="上一页", command=self.prev_page, state=tk.DISABLED)
        self.prev_button.pack(side=tk.LEFT, padx=(0, 10))

        self.page_label = ttk.Label(pagination_frame, text="第 1 页", font=self.default_font)
        self.page_label.pack(side=tk.LEFT)

        self.next_button = ttk.Button(pagination_frame, text="下一页", command=self.next_page)
        self.next_button.pack(side=tk.LEFT, padx=(10, 0))

        # 刷新单词列表
        self.refresh_word_list()

    def create_history_content(self):
        """创建历史记录内容区域"""
        self.history_frame = ttk.Frame(self.content_frame)
        self.history_frame.pack(fill=tk.BOTH, expand=True)

        # 历史记录表格
        columns = ("日期", "正确数", "总数", "正确率")
        self.history_tree = ttk.Treeview(self.history_frame, columns=columns, show="headings", height=15)

        # 设置列宽和标题
        for col in columns:
            self.history_tree.heading(col, text=col)
            self.history_tree.column(col, width=150, anchor=tk.CENTER)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.history_frame, orient=tk.VERTICAL, command=self.history_tree.yview)
        self.history_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.history_tree.pack(fill=tk.BOTH, expand=True)

        # 绑定选择事件
        self.history_tree.bind("<<TreeviewSelect>>", self.on_history_select)

        # 详细结果区域
        self.detail_frame = ttk.LabelFrame(self.history_frame, text="详细结果", padding="10")
        self.detail_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        self.detail_text = tk.Text(self.detail_frame, font=self.default_font, wrap=tk.WORD, height=10)
        self.detail_text.pack(fill=tk.BOTH, expand=True)
        self.detail_text.config(state=tk.DISABLED)

        # 刷新历史记录
        self.refresh_history()

    def create_stats_content(self):
        """创建学习统计内容区域"""
        self.stats_frame = ttk.Frame(self.content_frame)
        self.stats_frame.pack(fill=tk.BOTH, expand=True)

        # 创建图表框架
        chart_frame = ttk.LabelFrame(self.stats_frame, text="学习统计", padding="10")
        chart_frame.pack(fill=tk.BOTH, expand=True)

        # 创建图表
        self.figure, (self.ax1, self.ax2) = plt.subplots(1, 2, figsize=(10, 5))
        self.canvas = FigureCanvasTkAgg(self.figure, master=chart_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        # 更新统计图表
        self.update_stats_charts()

    def show_test_content(self):
        """显示测验内容区域"""
        self.clear_content_frame()
        self.test_button.config(style='ActiveNav.TButton')
        self.dictionary_button.config(style='Nav.TButton')
        self.history_button.config(style='Nav.TButton')
        self.stats_button.config(style='Nav.TButton')

        self.create_test_content()
        if self.words:
            self.show_welcome()
        else:
            self.show_empty_state()

    def show_dictionary_content(self):
        """显示生词本内容区域"""
        self.clear_content_frame()
        self.test_button.config(style='Nav.TButton')
        self.dictionary_button.config(style='ActiveNav.TButton')
        self.history_button.config(style='Nav.TButton')
        self.stats_button.config(style='Nav.TButton')

        self.create_dictionary_content()

    def show_history_content(self):
        """显示历史记录内容区域"""
        self.clear_content_frame()
        self.test_button.config(style='Nav.TButton')
        self.dictionary_button.config(style='Nav.TButton')
        self.history_button.config(style='ActiveNav.TButton')
        self.stats_button.config(style='Nav.TButton')

        self.create_history_content()

    def show_stats_content(self):
        """显示学习统计内容区域"""
        self.clear_content_frame()
        self.test_button.config(style='Nav.TButton')
        self.dictionary_button.config(style='Nav.TButton')
        self.history_button.config(style='Nav.TButton')
        self.stats_button.config(style='ActiveNav.TButton')

        self.create_stats_content()

    def clear_content_frame(self):
        """清空内容区域"""
        for widget in self.content_frame.winfo_children():
            widget.destroy()

    def on_window_resize(self, event):
        """窗口大小变化时动态调整字体"""
        if event.widget == self.root and event.width > 400:
            # 根据窗口宽度动态调整字体大小
            new_title_size = min(24, max(16, int(event.width / 60)))
            new_question_size = min(28, max(18, int(event.width / 50)))
            new_button_size = min(16, max(10, int(event.width / 100)))

            if new_title_size != self.title_font_size:
                self.title_font_size = new_title_size
                self.question_font_size = new_question_size
                self.button_font_size = new_button_size
                self.setup_fonts()

                # 更新测验区域字体
                if hasattr(self, 'word_label'):
                    self.word_label.config(font=self.question_font)
                    self.pos_label.config(font=self.question_font)  # 更新词性标签字体
                    self.meaning_label.config(font=self.question_font)
                    self.result_label.config(font=self.question_font)

                # 更新导航栏字体
                if hasattr(self, 'test_button'):
                    style = ttk.Style()
                    style.configure('Nav.TButton', font=self.button_font, padding=10)
                    style.configure('ActiveNav.TButton', font=self.button_font, padding=10,
                                    foreground=self.primary_color)

    def load_words(self):
        """加载单词数据"""
        try:
            if self.file_path.exists():
                wb = load_workbook(filename=self.file_path)
                sheet = wb.active

                # 检查列数
                max_column = sheet.max_column

                words = []
                for row in sheet.iter_rows(min_row=2, values_only=True):
                    if not row[0]:  # 如果第一列(单词)为空，则跳过
                        continue

                    word = row[0]
                    pos = row[1] if max_column >= 2 else ""  # 词性(如果有)
                    meaning = row[2] if max_column >= 3 else ""  # 词义(如果有)

                    words.append((word, pos, meaning))

                return wb, sheet, words
            else:
                # 创建新的Excel文件
                wb = Workbook()
                sheet = wb.active
                sheet.title = "单词本"
                sheet.append(["单词", "词性", "词义"])  # 更新表头为三列
                wb.save(filename=self.file_path)
                return wb, sheet, []
        except Exception as e:
            messagebox.showerror("错误", f"加载单词文件时出错: {e}")
            return None, None, []

    def load_history(self):
        """加载历史记录"""
        try:
            if self.history_path.exists():
                wb = load_workbook(filename=self.history_path)
                sheet = wb.active
                self.history_records = []
                for row in sheet.iter_rows(min_row=2, values_only=True):
                    if row[0] and row[1] and row[2]:
                        # 解析错误单词
                        wrong_words = []
                        if len(row) > 3 and row[3]:
                            wrong_items = row[3].split(';')
                            for item in wrong_items:
                                if item:
                                    parts = item.split(':')
                                    if len(parts) >= 3:  # 处理包含词性的错误单词
                                        word, pos, meaning = parts[0], parts[1], ':'.join(parts[2:])
                                        wrong_words.append((word, pos, meaning))
                                    elif len(parts) == 2:  # 兼容旧格式
                                        word, meaning = parts
                                        wrong_words.append((word, "", meaning))
                        self.history_records.append({
                            'date': row[0],
                            'correct': row[1],
                            'total': row[2],
                            'wrong_words': wrong_words
                        })
        except Exception as e:
            messagebox.showerror("错误", f"加载历史记录时出错: {e}")

    def save_words(self):
        """保存单词数据到Excel"""
        try:
            if not self.wb:
                self.wb = Workbook()
                self.sheet = self.wb.active
                self.sheet.title = "单词本"
                self.sheet.append(["单词", "词性", "词义"])  # 更新表头为三列

            # 清空现有数据
            self.sheet.delete_rows(2, self.sheet.max_row)

            # 添加新数据
            for word, pos, meaning in self.words:
                self.sheet.append([word, pos, meaning])

            self.wb.save(filename=self.file_path)
            self.word_count_label.config(text=f"单词总数: {len(self.words)}")
            return True
        except Exception as e:
            messagebox.showerror("错误", f"保存单词文件时出错: {e}")
            return False

    def save_history(self):
        """保存历史记录到Excel"""
        try:
            wb = Workbook()
            sheet = wb.active
            sheet.title = "测验历史"
            sheet.append(["日期", "正确数", "总数", "错误单词"])

            for record in self.history_records:
                # 格式化错误单词 - 包含词性
                wrong_words_str = ';'.join([f"{word}:{pos}:{meaning}" for word, pos, meaning in record['wrong_words']])
                sheet.append([record['date'], record['correct'], record['total'], wrong_words_str])

            wb.save(filename=self.history_path)
            return True
        except Exception as e:
            messagebox.showerror("错误", f"保存历史记录时出错: {e}")
            return False

    def select_questions(self):
        """选择测试题目"""
        if len(self.words) < 4:
            messagebox.showinfo("提示", "单词数量不足，请先添加更多单词。")
            return []

        # 随机选择20个单词（如果有足够的单词）
        num_questions = min(20, len(self.words))
        return random.sample(self.words, num_questions)

    def show_welcome(self):
        """显示欢迎信息"""
        self.word_label.config(text="欢迎使用单词测验程序！")
        self.pos_label.config(text="")  # 清空词性标签
        self.meaning_label.config(text="请选择测验模式并点击开始测试。")
        self.clear_options()
        self.result_label.config(text="")

    def show_empty_state(self):
        """显示空状态信息"""
        self.word_label.config(text="单词本为空")
        self.pos_label.config(text="")  # 清空词性标签
        self.meaning_label.config(text="请导入Excel文件或添加新单词。")
        self.clear_options()
        self.result_label.config(text="")

    def show_options(self, correct_answer, all_options):
        """显示选项"""
        self.clear_options()

        # 确保有足够的选项
        options = set(all_options)
        options.discard(correct_answer)  # 移除正确答案，避免重复
        options = list(options)

        # 如果选项不足，添加一些随机选项（可能会重复）
        while len(options) < 3:
            random_option = random.choice(all_options)
            if random_option != correct_answer:
                options.append(random_option)

        # 随机选择3个错误选项
        wrong_options = random.sample(options, 3)

        # 合并正确答案和错误选项，并打乱顺序
        all_display_options = [correct_answer] + wrong_options
        random.shuffle(all_display_options)

        # 创建选项按钮
        for i, option in enumerate(all_display_options):
            is_correct = option == correct_answer
            radio_btn = ttk.Radiobutton(
                self.options_frame,
                text=option,
                variable=self.selected_var,
                value=option,
                style='Option.TRadiobutton',
                command=lambda opt=option, corr=is_correct: self.submit_answer(opt, corr)
            )
            radio_btn.pack(anchor=tk.W, pady=5)

    def clear_options(self):
        """清除所有选项"""
        for widget in self.options_frame.winfo_children():
            widget.destroy()

    def start_test(self):
        """开始测试"""
        if not self.words:
            self.result_label.config(text="未成功加载单词，请检查文件或添加新单词。", foreground="red")
            return

        self.test_questions = self.select_questions()
        if not self.test_questions:
            return

        self.test_started = True
        self.start_time = datetime.now()
        self.time_remaining = self.time_limit
        self.correct_count = 0
        self.total_count = 0
        self.wrong_answers = []
        self.question_index = 0

        self.update_timer()
        self.update_progress()
        self.show_question()

    def reset_test(self):
        """重置测试"""
        self.test_started = False
        self.time_remaining = self.time_limit
        self.correct_count = 0
        self.total_count = 0
        self.wrong_answers = []
        self.question_index = 0

        if self.timer_id:
            self.root.after_cancel(self.timer_id)
            self.timer_id = None

        self.timer_label.config(text="20:00", foreground="green")
        self.update_progress()

        if self.words:
            self.show_welcome()
        else:
            self.show_empty_state()

    def show_question(self):
        """显示当前题目"""
        if self.question_index < len(self.test_questions):
            self.current_word = self.test_questions[self.question_index]

            if self.mode.get() == "word_to_meaning":
                self.word_label.config(text=f"单词：{self.current_word[0]}")
                self.pos_label.config(text=f"词性：{self.current_word[1]}")
                self.meaning_label.config(text="")
                self.show_options(self.current_word[2], [word[2] for word in self.words if word[2]])
            else:
                self.word_label.config(text="")
                self.pos_label.config(text=f"词性：{self.current_word[1]}")
                self.meaning_label.config(text=f"释义：{self.current_word[2]}")
                self.show_options(self.current_word[0], [word[0] for word in self.words if word[0]])

            self.update_progress()
        else:
            self.show_result()

    def submit_answer(self, selected_option, is_correct):
        """提交答案（自动校验并跳转下一题）"""
        if not selected_option:
            return

        self.total_count += 1

        if is_correct:
            self.correct_count += 1
            self.result_label.config(text="回答正确！", foreground="green")
        else:
            correct_answer = self.current_word[2] if self.mode.get() == "word_to_meaning" else self.current_word[0]
            self.result_label.config(text=f"回答错误，正确答案是：{correct_answer}", foreground="red")
            self.wrong_answers.append(self.current_word)  # 保存错误单词

        # 短暂显示结果后继续
        self.root.after(1500, self.next_question)

    def next_question(self):
        """移动到下一题或结束测试"""
        self.question_index += 1

        if self.question_index < len(self.test_questions):
            self.show_question()  # 显示下一题
            self.result_label.config(text="")  # 清空结果提示
        else:
            self.show_result()  # 测试结束，显示汇总结果

        self.update_progress()

    def show_result(self):
        """显示测试结果"""
        if self.timer_id:
            self.root.after_cancel(self.timer_id)
            self.timer_id = None

        self.test_started = False

        # 计算正确率
        accuracy = (self.correct_count / self.total_count) * 100 if self.total_count > 0 else 0

        # 构建结果文本
        result_text = f"测试完成！\n"
        result_text += f"正确数：{self.correct_count}/{self.total_count}\n"
        result_text += f"正确率：{accuracy:.2f}%\n\n"

        if self.wrong_answers:
            result_text += "错误的单词：\n"
            for word, pos, meaning in self.wrong_answers:
                result_text += f"- {word} [{pos}]：{meaning}\n"

        self.word_label.config(text="测试结果")
        self.pos_label.config(text="")  # 清空词性标签
        self.meaning_label.config(text="")
        self.clear_options()
        self.result_label.config(text=result_text)

        # 保存历史记录
        self.history_records.append({
            'date': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'correct': self.correct_count,
            'total': self.total_count,
            'wrong_words': self.wrong_answers
        })
        self.save_history()

    def update_timer(self):
        """更新计时器显示"""
        if not self.test_started:
            return

        elapsed = datetime.now() - self.start_time
        remaining = self.time_limit - elapsed

        if remaining.total_seconds() <= 0:
            self.timer_label.config(text="00:00", foreground="red")
            self.show_result()
            return

        minutes, seconds = divmod(int(remaining.total_seconds()), 60)
        time_str = f"{minutes:02d}:{seconds:02d}"

        # 根据剩余时间改变颜色
        if remaining.total_seconds() <= 60:
            self.timer_label.config(foreground="red")
        elif remaining.total_seconds() <= 300:
            self.timer_label.config(foreground="orange")
        else:
            self.timer_label.config(foreground="green")

        self.timer_label.config(text=time_str)
        self.timer_id = self.root.after(1000, self.update_timer)

    def update_progress(self):
        """更新进度显示"""
        total_questions = len(self.test_questions) if self.test_questions else 0
        self.progress_label.config(text=f"{self.question_index}/{total_questions}")

    def add_word(self):
        """添加新单词"""
        word = self.word_to_add.get().strip()
        pos = self.pos_to_add.get().strip()  # 获取词性
        meaning = self.meaning_to_add.get().strip()

        if not word or not meaning:
            messagebox.showwarning("警告", "单词和释义不能为空！")
            return

        # 检查是否已存在该单词
        for w, p, m in self.words:
            if w.lower() == word.lower():
                messagebox.showinfo("提示", f"单词 '{word}' 已存在！")
                return

        # 添加到单词列表
        self.words.append((word, pos, meaning))
        self.save_words()
        self.refresh_word_list()

        # 清空输入框
        self.word_to_add.set("")
        self.pos_to_add.set("")
        self.meaning_to_add.set("")

        messagebox.showinfo("成功", f"单词 '{word}' 已添加！")

    def search_words(self):
        """搜索单词"""
        term = self.search_term.get().strip().lower()
        if not term:
            self.current_words_display = self.words.copy()
        else:
            self.current_words_display = []
            for word, pos, meaning in self.words:
                if term in word.lower() or term in pos.lower() or term in meaning.lower():
                    self.current_words_display.append((word, pos, meaning))

        self.current_page = 1
        self.refresh_word_list()

    def refresh_word_list(self):
        """刷新单词列表"""
        # 清空当前显示
        for item in self.word_tree.get_children():
            self.word_tree.delete(item)

        # 获取当前页的单词
        start_idx = (self.current_page - 1) * self.words_per_page
        end_idx = min(start_idx + self.words_per_page, len(self.current_words_display))
        current_words = self.current_words_display[start_idx:end_idx]

        # 添加到表格
        for word, pos, meaning in current_words:
            self.word_tree.insert("", tk.END, values=(word, pos, meaning))

        # 更新分页控件
        total_pages = (len(self.current_words_display) + self.words_per_page - 1) // self.words_per_page
        self.page_label.config(text=f"第 {self.current_page} 页，共 {total_pages} 页")

        # 启用/禁用分页按钮
        self.prev_button.config(state=tk.NORMAL if self.current_page > 1 else tk.DISABLED)
        self.next_button.config(state=tk.NORMAL if self.current_page < total_pages else tk.DISABLED)

    def prev_page(self):
        """上一页"""
        if self.current_page > 1:
            self.current_page -= 1
            self.refresh_word_list()

    def next_page(self):
        """下一页"""
        total_pages = (len(self.current_words_display) + self.words_per_page - 1) // self.words_per_page
        if self.current_page < total_pages:
            self.current_page += 1
            self.refresh_word_list()

    def on_double_click_word(self, event):
        """双击删除单词"""
        item = self.word_tree.selection()
        if not item:
            return

        word, pos, meaning = self.word_tree.item(item[0], "values")
        if messagebox.askyesno("确认", f"确定要删除单词 '{word}' 吗？"):
            # 从列表中删除
            self.words = [w for w in self.words if w[0] != word]
            self.save_words()
            self.search_words()  # 刷新显示

    def refresh_history(self):
        """刷新历史记录"""
        # 清空当前显示
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)

        # 添加到表格
        for record in self.history_records:
            accuracy = (record['correct'] / record['total'] * 100) if record['total'] > 0 else 0
            self.history_tree.insert("", tk.END, values=(
                record['date'],
                record['correct'],
                record['total'],
                f"{accuracy:.2f}%"
            ))

    def on_history_select(self, event):
        """历史记录选择事件"""
        item = self.history_tree.selection()
        if not item:
            return

        # 获取选中的记录
        record_idx = int(item[0].replace("I", ""), 16) - 1
        if 0 <= record_idx < len(self.history_records):
            record = self.history_records[record_idx]

            # 构建详细结果
            detail_text = f"日期: {record['date']}\n"
            detail_text += f"正确数: {record['correct']}/{record['total']}\n"
            accuracy = (record['correct'] / record['total'] * 100) if record['total'] > 0 else 0
            detail_text += f"正确率: {accuracy:.2f}%\n\n"

            if record['wrong_words']:
                detail_text += "错误的单词：\n"
                for word, pos, meaning in record['wrong_words']:
                    detail_text += f"- {word} [{pos}]：{meaning}\n"
            else:
                detail_text += "恭喜！没有错误的单词。"

            # 显示详细结果
            self.detail_text.config(state=tk.NORMAL)
            self.detail_text.delete(1.0, tk.END)
            self.detail_text.insert(tk.END, detail_text)
            self.detail_text.config(state=tk.DISABLED)

    def update_stats_charts(self):
        """更新统计图表"""
        # 清除现有图表
        self.ax1.clear()
        self.ax2.clear()

        # 正确率趋势图
        if self.history_records:
            dates = [record['date'] for record in self.history_records]
            accuracies = [(record['correct'] / record['total'] * 100) for record in self.history_records]

            self.ax1.plot(dates, accuracies, marker='o', linestyle='-', color=self.primary_color)
            self.ax1.set_title('正确率趋势')
            self.ax1.set_xlabel('日期')
            self.ax1.set_ylabel('正确率 (%)')
            self.ax1.set_ylim(0, 105)
            self.ax1.tick_params(axis='x', rotation=45)

            # 添加数据标签
            for x, y in zip(dates, accuracies):
                self.ax1.annotate(f'{y:.1f}%', (x, y), textcoords='offset points',
                                  xytext=(0, 5), ha='center')

        # 错误单词分布图
        if self.history_records:
            # 统计每个单词的错误次数
            word_errors = {}
            for record in self.history_records:
                for word, pos, meaning in record['wrong_words']:
                    full_word = f"{word} [{pos}]"
                    if full_word in word_errors:
                        word_errors[full_word] += 1
                    else:
                        word_errors[full_word] = 1

            if word_errors:
                # 取错误最多的10个单词
                top_words = sorted(word_errors.items(), key=lambda x: x[1], reverse=True)[:10]
                words = [word for word, count in top_words]
                counts = [count for word, count in top_words]

                self.ax2.barh(words, counts, color=self.accent_color)
                self.ax2.set_title('错误最多的单词')
                self.ax2.set_xlabel('错误次数')

                # 添加数据标签
                for i, (word, count) in enumerate(top_words):
                    self.ax2.text(count + 0.2, i, str(count), va='center')

        # 调整布局并绘制
        self.figure.tight_layout()
        self.canvas.draw()

    def import_excel_file(self):
        """导入Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )

        if not file_path:
            return

        try:
            wb = load_workbook(filename=file_path)
            sheet = wb.active

            # 检查列数
            max_column = sheet.max_column

            new_words = []
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if not row[0]:  # 如果第一列(单词)为空，则跳过
                    continue

                word = row[0]
                pos = row[1] if max_column >= 2 else ""  # 词性(如果有)
                meaning = row[2] if max_column >= 3 else ""  # 词义(如果有)

                # 检查是否已存在该单词
                exists = False
                for w, p, m in self.words:
                    if w.lower() == word.lower():
                        exists = True
                        break

                if not exists:
                    new_words.append((word, pos, meaning))

            if new_words:
                self.words.extend(new_words)
                self.save_words()
                self.current_words_display = self.words.copy()
                self.refresh_word_list()
                messagebox.showinfo("成功", f"已成功导入 {len(new_words)} 个新单词！")
            else:
                messagebox.showinfo("提示", "没有找到新的单词或文件格式不正确。")

        except Exception as e:
            messagebox.showerror("错误", f"导入文件时出错: {e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = VocabularyTestApp(root)
    root.mainloop()