import tkinter as tk
from tkinter import filedialog
import subprocess

def select_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        entry_path.delete(0, tk.END)
        entry_path.insert(0, folder_path)

def process_folder():
    folder_path = entry_path.get()
    if folder_path:
        try:
            subprocess.run(["python", "extract.py", folder_path])
            tk.messagebox.showinfo("Success", "PDF 文件夹处理完成！")
        except Exception as e:
            tk.messagebox.showerror("Error", f"处理出错：{e}")
    else:
        tk.messagebox.showwarning("Warning", "请选择 PDF 文件夹路径！")

# 创建主窗口
root = tk.Tk()
root.title("PDF 文件夹处理工具")

# 添加按钮用于选择文件夹
btn_browse = tk.Button(root, text="选择文件夹", command=select_folder)
btn_browse.pack(pady=10)

# 添加文本框用于显示选择的文件夹路径
entry_path = tk.Entry(root, width=50)
entry_path.pack(pady=5)

# 添加按钮用于触发处理
btn_process = tk.Button(root, text="处理 PDF 文件夹", command=process_folder)
btn_process.pack(pady=5)

# 运行主循环
root.mainloop()
