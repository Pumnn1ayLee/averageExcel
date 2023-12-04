import pandas as pd
import tkinter as tk
from tkinter import filedialog

def calculate_and_save():
    file_path = file_path_entry.get()
    keyword_columns = keyword_entry.get().split(',')
    number_columns = number_entry.get().split(',')
    output_path = output_path_entry.get()

    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)

        # 选择关键词列和数字列
        selected_columns = keyword_columns + number_columns

        # 确保所选的列存在
        missing_columns = [col for col in selected_columns if col not in df.columns]
        if missing_columns:
            result_text.config(text=f"以下列不存在: {', '.join(missing_columns)}")
            return

        # 计算平均值
        result = df.groupby(keyword_columns)[number_columns].mean()

        # 保存结果到文本文件
        result.to_csv(output_path, header=True, index=True, sep='\t')
        result_text.config(text=f"结果已保存到：{output_path}")
    except FileNotFoundError:
        result_text.config(text="文件未找到或文件格式不正确。请检查文件路径和文件内容.")

# 创建窗口
window = tk.Tk()
window.title("Excel 数据平均值计算")

# 创建输入框和标签
file_path_label = tk.Label(window, text="Excel文件路径:")
file_path_label.pack()
file_path_entry = tk.Entry(window)
file_path_entry.pack()

keyword_label = tk.Label(window, text="关键词列名称 (用逗号分隔多个列):")
keyword_label.pack()
keyword_entry = tk.Entry(window)
keyword_entry.pack()

number_label = tk.Label(window, text="数字列名称 (用逗号分隔多个列):")
number_label.pack()
number_entry = tk.Entry(window)
number_entry.pack()

output_path_label = tk.Label(window, text="输出结果路径:")
output_path_label.pack()
output_path_entry = tk.Entry(window)
output_path_entry.pack()

file_button = tk.Button(window, text="选择Excel文件", command=lambda: file_path_entry.insert(0, filedialog.askopenfilename()))
file_button.pack()

output_button = tk.Button(window, text="选择输出结果路径", command=lambda: output_path_entry.insert(0, filedialog.asksaveasfilename(defaultextension=".txt")))
output_button.pack()

calculate_button = tk.Button(window, text="计算平均值并保存", command=calculate_and_save)
calculate_button.pack()

result_text = tk.Label(window, text="")
result_text.pack()

window.mainloop()
