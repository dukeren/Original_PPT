import os
import win32com.client
import pythoncom
import sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from ttkthemes import ThemedTk
from PIL import Image, ImageTk  # 请确保安装了 Pillow 库

def remove_unused_layouts(ppt_path, output_dir):
    pythoncom.CoInitialize()
    powerpoint = None
    presentation = None

    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = True

        status_text.insert(tk.END, f"正在打开文件: {ppt_path}\n")
        presentation = powerpoint.Presentations.Open(ppt_path)
        status_text.insert(tk.END, "文件已成功打开\n")

        status_text.insert(tk.END, "获取使用的布局...\n")
        used_layouts = set()
        for slide in presentation.Slides:
            layout_index = slide.Layout
            if isinstance(layout_index, int):
                layout = slide.CustomLayout
                used_layouts.add(layout.Name)
            else:
                used_layouts.add(slide.Layout.Name)
        status_text.insert(tk.END, f"使用的布局数量: {len(used_layouts)}\n")

        status_text.insert(tk.END, "处理母版和布局...\n")
        for master in presentation.Designs:
            status_text.insert(tk.END, f"处理母版: {master.Name}\n")
            layouts_to_remove = []
            
            for layout in master.SlideMaster.CustomLayouts:
                if layout.Name not in used_layouts:
                    layouts_to_remove.append(layout)
            
            for layout in layouts_to_remove:
                status_text.insert(tk.END, f"删除未使用的布局: {layout.Name}\n")
                layout.Delete()

        output_filename = "output_" + os.path.basename(ppt_path)
        output_path = os.path.join(output_dir, output_filename)

        status_text.insert(tk.END, f"正在保存文件到: {output_path}\n")
        presentation.SaveAs(output_path)
        status_text.insert(tk.END, "文件已保存\n")

    except Exception as e:
        status_text.insert(tk.END, f"发生错误: {str(e)}\n")
        status_text.insert(tk.END, f"错误类型: {type(e).__name__}\n")
        status_text.insert(tk.END, f"错误发生在: {sys.exc_info()[2].tb_lineno}行\n")
    finally:
        if presentation:
            presentation.Close()
        if powerpoint:
            powerpoint.Quit()
        pythoncom.CoUninitialize()

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx")])
    if file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)

def select_output_dir():
    dir_path = filedialog.askdirectory()
    if dir_path:
        output_dir_entry.delete(0, tk.END)
        output_dir_entry.insert(0, dir_path)

def execute():
    ppt_path = file_entry.get()
    output_dir = output_dir_entry.get()
    if not ppt_path:
        messagebox.showerror("错误", "请选择一个PowerPoint文件")
        return
    if not os.path.exists(ppt_path):
        messagebox.showerror("错误", f"找不到文件: {ppt_path}")
        return
    if not output_dir:
        messagebox.showerror("错误", "请选择输出目录")
        return
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    status_text.delete(1.0, tk.END)
    remove_unused_layouts(ppt_path, output_dir)
    messagebox.showinfo("完成", "操作已完成")

# 创建主窗口
root = ThemedTk(theme="ubuntu")
root.title("PowerPoint布局清理 By 渡客")
root.geometry("670x400")

icon_path = os.path.join(os.path.dirname(__file__), "Image", "logo.ico")
if os.path.exists(icon_path):
    icon_image = Image.open(icon_path)
    icon_photo = ImageTk.PhotoImage(icon_image)
    root.iconphoto(True, icon_photo)

# 创建文件选择部分
file_frame = ttk.Frame(root)
file_frame.pack(pady=10, padx=10, fill=tk.X)

file_label = ttk.Label(file_frame, text="选择文件:")
file_label.pack(side=tk.LEFT)

file_entry = ttk.Entry(file_frame)
file_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 5))

file_button = ttk.Button(file_frame, text="浏览", command=select_file)
file_button.pack(side=tk.LEFT)

# 创建输出目录选择部分
output_dir_frame = ttk.Frame(root)
output_dir_frame.pack(pady=10, padx=10, fill=tk.X)

output_dir_label = ttk.Label(output_dir_frame, text="输出目录:")
output_dir_label.pack(side=tk.LEFT)

output_dir_entry = ttk.Entry(output_dir_frame)
output_dir_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 5))

output_dir_button = ttk.Button(output_dir_frame, text="浏览", command=select_output_dir)
output_dir_button.pack(side=tk.LEFT)

# 设置默认输出目录
default_output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Outfile")
output_dir_entry.insert(0, default_output_dir)

# 创建执行按钮
execute_button = ttk.Button(root, text="执行", command=execute)
execute_button.pack(pady=10)

# 创建状态文本框
status_text = tk.Text(root, height=20, width=70)
status_text.pack(pady=10, padx=10, expand=True, fill=tk.BOTH)

# 添加滚动条
scrollbar = ttk.Scrollbar(status_text)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
status_text.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=status_text.yview)

# 运行主循环
root.mainloop()