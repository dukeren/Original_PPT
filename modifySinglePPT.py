import win32com.client as win32
from pptx import Presentation
import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.ttk as ttk
from ttkthemes import ThemedTk
import win32gui
import os
import tempfile
import shutil
import atexit
from PIL import Image, ImageTk
import pythoncom

class PowerPointHandler:
    def __init__(self):
        self.powerpoint = None
        self.presentation = None

    def open_powerpoint(self):
        if not self.powerpoint:
            pythoncom.CoInitialize()
            self.powerpoint = win32.Dispatch("PowerPoint.Application")

    def open_presentation(self, ppt_path):
        self.open_powerpoint()
        if self.presentation:
            self.presentation.Close()
        self.presentation = self.powerpoint.Presentations.Open(ppt_path, WithWindow=False)

    def close_presentation(self):
        if self.presentation:
            self.presentation.Close()
            self.presentation = None

    def close_powerpoint(self):
        if self.powerpoint:
            self.powerpoint.Quit()
            self.powerpoint = None
        pythoncom.CoUninitialize()

    def __del__(self):
        self.close_presentation()
        self.close_powerpoint()

ppt_handler = PowerPointHandler()

def select_ppt():
    file_path = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx")])
    if file_path:
        ppt_path_var.set(file_path)
        update_slide_list(file_path)

def update_slide_list(ppt_path):
    try:
        prs = Presentation(ppt_path)
        slide_list.delete(0, tk.END)
        for i, slide in enumerate(prs.slides, start=1):
            slide_list.insert(tk.END, f"Slide {i}")
        ppt_handler.open_presentation(ppt_path)
    except Exception as e:
        messagebox.showerror("Error", f"无法打开PPT文件: {e}")

def apply_layout():
    reference_ppt_path = ppt_path_var.get()
    selected_indices = slide_list.curselection()
    if not reference_ppt_path or not selected_indices:
        messagebox.showwarning("警告", "请选择参考PPT和幻灯片")
        return
    reference_slide_index = selected_indices[0]
    
    current_foreground = win32gui.GetForegroundWindow()
    
    apply_reference_layout(reference_ppt_path, reference_slide_index)
    
    win32gui.SetForegroundWindow(current_foreground)

def apply_reference_layout(reference_ppt_path: str, reference_slide_index: int):
    temp_ppt_path = None
    temp_presentation = None
    try:
        powerpoint = win32.Dispatch("PowerPoint.Application")
        current_presentation = powerpoint.ActivePresentation
        if not current_presentation:
            messagebox.showerror("错误", "没有打开的PowerPoint演示文稿。")
            return

        current_slide = powerpoint.ActiveWindow.View.Slide
        if not current_slide:
            messagebox.showerror("错误", "没有选中的幻灯片。")
            return

        print(f"当前选中的幻灯片: 第 {current_slide.SlideIndex} 页")

        temp_dir = tempfile.gettempdir()
        temp_ppt_path = os.path.join(temp_dir, f"temp_reference_{os.getpid()}.pptx")
        
        shutil.copy2(reference_ppt_path, temp_ppt_path)
        temp_presentation = powerpoint.Presentations.Open(temp_ppt_path)
        
        reference_slide = temp_presentation.Slides(reference_slide_index + 1)

        for shape in current_slide.Shapes:
            try:
                matching_shape = None
                for ref_shape in reference_slide.Shapes:
                    if ref_shape.Name == shape.Name:
                        matching_shape = ref_shape
                        break

                if matching_shape:
                    if shape.HasTextFrame:
                        matching_shape.TextFrame.TextRange.Text = shape.TextFrame.TextRange.Text
                else:
                    shape.Copy()
                    reference_slide.Shapes.Paste()

            except Exception as e:
                print(f"处理形状 '{shape.Name}' 时发生错误: {e}")

        reference_slide.Copy()
        new_slide = current_presentation.Slides.Paste(current_slide.SlideIndex + 1)
        
        current_slide.Delete()

        powerpoint.ActiveWindow.ViewType = 1
        powerpoint.ActiveWindow.View.GotoSlide(new_slide.SlideIndex)
        
        messagebox.showinfo("成功", f"已成功根据参考PPT的第{reference_slide_index + 1}页更新了当前PPT的选中页面布局。")

    except Exception as e:
        messagebox.showerror("错误", f"发生错误: {e}")
    
    finally:
        if temp_presentation:
            try:
                temp_presentation.Close()
            except:
                pass
        
        if temp_ppt_path and os.path.exists(temp_ppt_path):
            try:
                os.remove(temp_ppt_path)
            except Exception as e:
                print(f"无法删除临时文件 {temp_ppt_path}: {e}")
                atexit.register(lambda file=temp_ppt_path: os.remove(file) if os.path.exists(file) else None)

def on_slide_select(event):
    selected_indices = slide_list.curselection()
    if selected_indices:
        selected_index = selected_indices[0]
        update_preview(selected_index)

def update_preview(slide_index):
    ppt_path = ppt_path_var.get()
    if not ppt_path:
        return

    try:
        if not ppt_handler.presentation:
            ppt_handler.open_presentation(ppt_path)
        
        if not ppt_handler.presentation:
            raise Exception("无法打开演示文稿")

        slide = ppt_handler.presentation.Slides(slide_index + 1)
        
        temp_image_path = os.path.join(tempfile.gettempdir(), f"slide_preview_{os.getpid()}.png")
        try:
            slide.Export(temp_image_path, "PNG")
        except Exception as export_error:
            print(f"导出幻灯片时出错: {export_error}")
            messagebox.showerror("预览错误", "无法导出幻灯片进行预览")
            return

        with Image.open(temp_image_path) as img:
            img.thumbnail((300, 200))
            photo = ImageTk.PhotoImage(img)
        
        preview_canvas.delete("all")
        preview_canvas.config(width=photo.width(), height=photo.height())
        preview_canvas.create_image(0, 0, anchor=tk.NW, image=photo)
        preview_canvas.image = photo

        os.remove(temp_image_path)
        
    except Exception as e:
        messagebox.showerror("预览错误", f"无法生成预览: {e}")
        print(f"预览错误详情: {e}")

# 主窗口设置
root = ThemedTk(theme="ubuntu")
root.title("PPT单页布局修改 By 渡客")
root.geometry("760x500")

icon_path = os.path.join(os.path.dirname(__file__), "Image", "logo.ico")
if os.path.exists(icon_path):
    icon_image = Image.open(icon_path)
    icon_photo = ImageTk.PhotoImage(icon_image)
    root.iconphoto(True, icon_photo)

style = ttk.Style()
style.theme_use("ubuntu")

left_frame = ttk.Frame(root)
left_frame.pack(side=tk.LEFT, padx=10, pady=10, fill=tk.BOTH, expand=True)

right_frame = ttk.Frame(root)
right_frame.pack(side=tk.RIGHT, padx=10, pady=10, fill=tk.BOTH, expand=True)

ppt_path_var = tk.StringVar()
ttk.Label(left_frame, text="选择参考PPT:").pack(pady=5)
ttk.Entry(left_frame, textvariable=ppt_path_var, width=40).pack(side=tk.TOP, pady=5)
ttk.Button(left_frame, text="浏览", command=select_ppt).pack(pady=5)

ttk.Label(left_frame, text="选择参考幻灯片:").pack(pady=5)
slide_list = tk.Listbox(left_frame, width=40, height=10)
slide_list.pack(pady=5)
slide_list.bind('<<ListboxSelect>>', on_slide_select)

apply_button = ttk.Button(left_frame, text="应用布局", command=apply_layout)
apply_button.pack(pady=10)

ttk.Label(right_frame, text="预览:").pack(pady=5)
preview_canvas = tk.Canvas(right_frame, width=300, height=200, bg='white')
preview_canvas.pack(pady=5)

def on_closing():
    ppt_handler.close_presentation()
    ppt_handler.close_powerpoint()
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_closing)

root.mainloop()