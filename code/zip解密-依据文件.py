import tkinter as tk
from tkinter import messagebox
import zipfile
import os

def create_password_dialog(title="密码输入", message="请输入ZIP文件密码:", file_path=None):
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
    
    # 创建密码输入框（显示明文）
    password_var = tk.StringVar()
    password_entry = tk.Entry(root, textvariable=password_var)  # 移除了show='*'参数
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
         # 第一次解压 - 解压到临时目录
        temp_dir = os.path.join(os.path.dirname(file_path), "temp_extract") # 临时目录
        try:

            os.makedirs(temp_dir, exist_ok=True)
            
            with zipfile.ZipFile(file_path) as zip_file:
                try:
                    zip_file.extractall(temp_dir, pwd=password.encode('utf-8'))
                    
                    # 查找解压后的zip文件进行第二次解压
                    for extracted_file in os.listdir(temp_dir):
                        if extracted_file.endswith('.zip'):
                            inner_zip_path = os.path.join(temp_dir, extracted_file)
                            
                            # 第二次解压
                            with zipfile.ZipFile(inner_zip_path) as inner_zip:
                                inner_zip.extractall(os.path.dirname(file_path))
                            
                            # 获取最终xlsx文件名
                            base_name = os.path.splitext(os.path.basename(file_path))[0]
                            output_file = os.path.join(os.path.dirname(file_path), f"{base_name}.xlsx")
                            print(f"准备将解压文件重命名为: {output_file}")
                            
                            # 只重命名从内部ZIP解压出来的xlsx文件
                            # 获取内部ZIP文件名（正确处理编码）
                            inner_zip_name = os.path.splitext(extracted_file)[0]
                            inner_zip_name = os.path.splitext(inner_zip_name)[0] 
                            
                            target_xlsx = f"{inner_zip_name}.xlsx"
                            print(f"预期解压文件名: {target_xlsx}")
                            
                            if os.path.exists(os.path.join(os.path.dirname(file_path), target_xlsx)):
                                print(f"找到正确的待重命名文件: {target_xlsx}")
                                os.rename(
                                    os.path.join(os.path.dirname(file_path), target_xlsx),
                                    output_file
                                )
                                print(f"文件重命名完成: {target_xlsx} -> {output_file}")
                            else:
                                print(f"错误: 未找到预期的xlsx文件 {target_xlsx}")
                            messagebox.showinfo("成功", f"文件已解压为:\n{output_file}")
                            root.destroy()
                            return
                    
                    messagebox.showerror("错误", "未找到内部ZIP文件")
                except RuntimeError as e:
                    if 'Bad password' in str(e):
                        messagebox.showerror("错误", "密码错误")
                    else:
                        messagebox.showerror("错误", f"解密失败: {str(e)}")
        except Exception as e:
            messagebox.showerror("错误", f"文件验证失败: {str(e)}")
        finally:
            # 清理临时目录
            if os.path.exists(temp_dir):
                for f in os.listdir(temp_dir):
                    os.remove(os.path.join(temp_dir, f))
                os.rmdir(temp_dir)
    
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
    file_path = rf"C:\project\python_project\AI+市场工具&稽核\AI+市场工具&稽核第一期需求\选项目-制表-得出结论\调剂依据文件.xlsx"
    file_path = file_path.replace('.xlsx', '.zip')
    create_password_dialog(file_path=file_path)