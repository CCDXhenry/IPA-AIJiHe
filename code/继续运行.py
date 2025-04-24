import tkinter as tk
from tkinter import messagebox

def create_continue_dialog(title="异常提示程序遇到异", message="程序遇到异常，需要人工手动程序遇到异常，需要人工手动"):
    root = tk.Tk()
    root.withdraw()  # 先隐藏窗口
    root.title(title)
    root.configure(highlightbackground='red', highlightcolor='red', highlightthickness=2)
    # 创建标题标签（红色）
    title_label = tk.Label(root, text=title, fg='red', font=('Arial', 12, 'bold'))
    title_label.pack(pady=10)
    
    # 创建临时标签计算文本宽度
    temp_label = tk.Label(root, text=message)
    temp_label.pack()
    temp_label.update()
    
    # 根据文本宽度动态设置窗口大小
    text_width = temp_label.winfo_reqwidth()
    window_width = min(max(text_width + 40, 200), 800)  # 最小200，最大800
    window_height = 150
    
    # 销毁临时标签
    temp_label.destroy()
    
    # 添加内容
    label = tk.Label(root, text=message, wraplength=window_width-40)
    label.pack(pady=20)
    
    # 添加按钮
    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)
    
    def on_continue():
        root.destroy()
    
    continue_btn = tk.Button(button_frame, text="继续运行", command=on_continue)
    continue_btn.pack()
    
    # 确保所有部件都已更新
    root.update_idletasks()
    
    # 计算窗口位置
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - root.winfo_width()) // 2
    y = screen_height - root.winfo_height() - 400  # 距离底部100像素
    
    # 设置窗口位置和属性
    root.geometry(f"+{x}+{y}")
    root.attributes("-topmost", True)
    root.deiconify()  # 显示窗口
    
    root.mainloop()

if __name__ == "__main__":
    create_continue_dialog()