import random
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import messagebox

from fpdf import FPDF
from openpyxl import load_workbook


def load_words():
    file_path = r'C:\Users\15658\Desktop\单词本.xlsx'
    try:
        wb = load_workbook(file_path)
        sheet = wb.active
        words = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            word = row[0]
            meaning = row[2]
            last_correct_time = row[3]
            correct_ratio = row[4] if row[4] else 0
            # 过滤掉单词或释义为空或 None 的情况
            if word and meaning:
                words.append((word, meaning, last_correct_time, correct_ratio))
        return wb, sheet, words
    except FileNotFoundError:
        messagebox.showerror("错误", f"未找到文件: {file_path}")
        return None, None, []
    except Exception as e:
        messagebox.showerror("错误", f"读取文件时出现错误: {e}")
        return None, None, []


def select_20_questions(words):
    eligible_words = []
    for word_info in words:
        word, meaning, last_correct_time, correct_ratio = word_info
        if correct_ratio < 0.99:
            if last_correct_time and (datetime.now() - last_correct_time) > timedelta(days=7):
                eligible_words.extend([word_info] * 5)
            else:
                eligible_words.append(word_info)
    if len(eligible_words) < 20:
        result_label.config(text="可用单词不足 20 个，无法开始测试。", fg="red")
        return []
    return random.sample(eligible_words, 20)


def start_test(mode, reset=False):
    global current_word, question_index, test_questions, correct_count, total_count, wrong_answers
    if reset:
        correct_count = 0
        total_count = 0
        wrong_answers = []
        question_index = 0
        test_questions = []

    if not words:
        result_label.config(text="未成功加载单词，请检查文件。", fg="red")
        return
    if question_index == 0:
        test_questions = select_20_questions(words)
        if not test_questions:
            return
    if question_index >= len(test_questions):
        result_label.config(text="测试已结束，请查看结果。", fg="red")
        return
    current_word = test_questions[question_index]
    if mode == "word_to_meaning":
        word_label.config(text=f"单词：{current_word[0]}")
        show_options(current_word[1], [word[1] for word in words if word[1]])
    elif mode == "meaning_to_word":
        meaning_label.config(text=f"释义：{current_word[1]}")
        show_options(current_word[0], [word[0] for word in words if word[0]])
    result_label.config(text="")


def show_options(correct_answer, all_options):
    global option_frame
    options = [correct_answer]
    while len(options) < 6:
        random_option = random.choice(all_options)
        if random_option and random_option not in options:
            options.append(random_option)
    random.shuffle(options)
    for widget in option_frame.winfo_children():
        widget.destroy()

    for i, option in enumerate(options):
        option_frame_single = tk.Frame(option_frame, bd=2, relief=tk.GROOVE)
        option_frame_single.pack(pady=5, fill=tk.X)

        tk.Checkbutton(option_frame_single, text=option, variable=selected_var, onvalue=option, offvalue="",
                       font=("Arial", 8), wraplength=200).pack(side=tk.LEFT, padx=5, pady=2)


def check_answer():
    global correct_count, total_count, wrong_answers, question_index
    if not words:
        return
    user_answer = selected_var.get()
    total_count += 1
    is_correct = False
    if mode.get() == "word_to_meaning":
        if user_answer == current_word[1]:
            result_label.config(text="回答正确！", fg="green")
            is_correct = True
        else:
            result_label.config(text=f"回答错误，正确答案是：{current_word[1]}", fg="red")
            wrong_answers.append((current_word[0], current_word[1], user_answer))
    elif mode.get() == "meaning_to_word":
        if user_answer == current_word[0]:
            result_label.config(text="回答正确！", fg="green")
            is_correct = True
        else:
            result_label.config(text=f"回答错误，正确答案是：{current_word[0]}", fg="red")
            wrong_answers.append((current_word[0], current_word[1], user_answer))
    if is_correct:
        correct_count += 1
        update_word_info(current_word, True)
    else:
        update_word_info(current_word, False)
    question_index += 1
    if total_count == 20:
        save_excel()
        show_result()
    else:
        start_test(mode.get())


def update_word_info(word_info, is_correct):
    for i, (word, meaning, last_correct_time, correct_ratio) in enumerate(words):
        if word == word_info[0] and meaning == word_info[1]:
            if is_correct:
                new_correct_ratio = (correct_ratio * (total_count - 1) + 1) / total_count
                words[i] = (word, meaning, datetime.now(), new_correct_ratio)
                sheet.cell(row=i + 2, column=4, value=datetime.now())
                sheet.cell(row=i + 2, column=5, value=new_correct_ratio)
            else:
                new_correct_ratio = correct_ratio * (total_count - 1) / total_count
                words[i] = (word, meaning, last_correct_time, new_correct_ratio)
                sheet.cell(row=i + 2, column=5, value=new_correct_ratio)
            break


def save_excel():
    try:
        wb.save(r'C:\Users\E8002877\Desktop\单词本.xlsx')
    except Exception as e:
        messagebox.showerror("错误", f"保存文件时出现错误: {e}")


def show_result():
    result_window = tk.Toplevel(root)
    result_window.title("测试结果")
    result_window.geometry("600x600")
    score = correct_count / total_count * 100
    result_text = f"本次测试共 {total_count} 题，你答对了 {correct_count} 题，得分：{score:.2f} 分。"
    result_label = tk.Label(result_window, text=result_text, font=("Arial", 12))
    result_label.pack(pady=20)
    error_analysis_button = tk.Button(result_window, text="查看错题解析", command=show_error_analysis)
    error_analysis_button.pack(pady=10)
    restart_button = tk.Button(result_window, text="重新开始", command=lambda: restart_test(result_window))
    restart_button.pack(pady=10)
    end_button = tk.Button(result_window, text="结束测试", command=lambda: end_test(result_window))
    end_button.pack(pady=10)


def show_error_analysis():
    error_window = tk.Toplevel(root)
    error_window.title("错题解析")
    error_window.geometry("600x600")
    text_box = tk.Text(error_window, font=("Arial", 12))
    text_box.pack(pady=20, padx=20, fill=tk.BOTH, expand=True)
    error_text = "错误答案解释：\n"
    for word, meaning, wrong_answer in wrong_answers:
        if mode.get() == "word_to_meaning":
            error_text += f"单词：{word}，你的答案：{wrong_answer}，正确答案：{meaning}\n"
        else:
            error_text += f"释义：{meaning}，你的答案：{wrong_answer}，正确答案：{word}\n"
    text_box.insert(tk.END, error_text)
    export_button = tk.Button(error_window, text="导出为 PDF",
                              command=lambda: export_to_pdf(text_box.get("1.0", tk.END)))
    export_button.pack(pady=10)


def export_to_pdf(content):
    try:
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=12)
        lines = content.split("\n")
        for line in lines:
            pdf.cell(200, 10, txt=line, ln=True, align='L')
        pdf_filename = "error_analysis.pdf"
        pdf.output(pdf_filename)
        messagebox.showinfo("导出成功", f"错题解析已成功导出为 {pdf_filename}")
    except Exception as e:
        messagebox.showerror("导出失败", f"导出 PDF 时出现错误：{e}")


def restart_test(result_window):
    global correct_count, total_count, wrong_answers, question_index, test_questions
    correct_count = 0
    total_count = 0
    wrong_answers = []
    question_index = 0
    test_questions = []
    result_window.destroy()
    start_test(mode.get(), reset=True)


def end_test(result_window):
    global correct_count, total_count, wrong_answers, question_index, test_questions
    correct_count = 0
    total_count = 0
    wrong_answers = []
    question_index = 0
    test_questions = []
    result_window.destroy()
    word_label.config(text="")
    meaning_label.config(text="")
    for widget in option_frame.winfo_children():
        widget.destroy()
    result_label.config(text="测试已结束，可重新开始。", fg="red")


# 创建主窗口
root = tk.Tk()
root.title("单词测验程序")
root.geometry("800x600")

# 在主窗口创建后创建变量
mode = tk.StringVar()
correct_count = 0
total_count = 0
wrong_answers = []
test_questions = []
question_index = 0
selected_var = tk.StringVar()

# 创建界面组件
word_label = tk.Label(root, text="", font=("Arial", 16))
word_label.pack(pady=10)

meaning_label = tk.Label(root, text="", font=("Arial", 16))
meaning_label.pack(pady=10)

option_frame = tk.Frame(root)
option_frame.pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 14))
result_label.pack(pady=10)

button_frame = tk.Frame(root)
button_frame.pack(side=tk.BOTTOM, anchor=tk.S, pady=20)

start_word_to_meaning_button = tk.Button(button_frame, text="根据单词选释义",
                                         command=lambda: [mode.set("word_to_meaning"),
                                                          start_test("word_to_meaning", reset=True)],
                                         font=("Arial", 14))
start_word_to_meaning_button.pack(side=tk.LEFT, padx=5)

start_meaning_to_word_button = tk.Button(button_frame, text="根据释义选单词",
                                         command=lambda: [mode.set("meaning_to_word"),
                                                          start_test("meaning_to_word", reset=True)],
                                         font=("Arial", 14))
start_meaning_to_word_button.pack(side=tk.LEFT, padx=5)

check_button = tk.Button(button_frame, text="检查答案", command=check_answer, font=("Arial", 14))
check_button.pack(side=tk.LEFT, padx=5)

# 加载单词数据
wb, sheet, words = load_words()

try:
    root.mainloop()
except Exception as e:
    messagebox.showerror("错误", f"主循环出错: {e}")
