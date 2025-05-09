import tkinter as tk
from tkinter import messagebox
from openpyxl import load_workbook
import msoffcrypto
import io
def create_password_dialog(title="密码输入", message="请输入文件密码:", file_path=None):
    root = tk.Tk()
    root.withdraw()  # 先隐藏窗口
    root.title(title)
    root.configure(highlightbackground='red', highlightcolor='red', highlightthickness=2)
    
    # 创建标题标签（红色）
    title_label = tk.Label(root, text=title, fg='red', font=('Arial', 12, 'bold'))
    title_label.pack(pady=10)
    
    # 创建提示信息标签
    message_label = tk.Label(root, text=message)
    message_label.pack()
    
    # 创建密码输入框
    password_var = tk.StringVar()
    password_entry = tk.Entry(root, textvariable=password_var)
    password_entry.pack(pady=10)
    
    # 添加按钮
    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)
    
    def on_submit():
        password = password_var.get()
        if not password:
            messagebox.showerror("错误", "密码不能为空")
            return
        
        if not file_path:
            messagebox.showerror("错误", "未提供文件路径")
            return
            
        try:
            # 使用 msoffcrypto-tool 来处理加密的 Excel 文件
            temp_buffer = io.BytesIO()
            
            with open(file_path, 'rb') as file:
                excel = msoffcrypto.OfficeFile(file)
                try:
                    excel.load_key(password=password)
                    # 解密到临时缓冲区
                    excel.decrypt(temp_buffer)
                    
                    # 将解密后的内容保存回原文件
                    temp_buffer.seek(0)
                    with open(file_path, 'wb') as output:
                        output.write(temp_buffer.read())
                        
                    messagebox.showinfo("成功", "密码已移除，文件已保存")
                    root.destroy()
                except Exception as e:
                    messagebox.showerror("错误", "密码错误")
        except Exception as e:
            messagebox.showerror("错误", f"文件验证失败: {str(e)}")
    
    submit_btn = tk.Button(button_frame, text="确认", command=on_submit)
    submit_btn.pack()
    
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
    file_path = rf"C:\project\python_project\关于新入网号码回访开发需求\20250402.xlsx"
    create_password_dialog(file_path=file_path)