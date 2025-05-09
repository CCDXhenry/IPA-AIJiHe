import tkinter as tk
from tkinter import messagebox
import zipfile
import os

def extract_zip(zip_path, password=None):
    """解压zip文件，支持双重压缩"""
    try:
        base_name = os.path.splitext(os.path.basename(zip_path))[0]
        output_dir = os.path.dirname(zip_path)
        
        with zipfile.ZipFile(zip_path) as zip_file:
            # 获取第一个文件名
            first_file = zip_file.namelist()[0]
            # 临时输出文件路径
            temp_output = os.path.join(output_dir, first_file)
            
            # 使用密码解压第一个文件
            with zip_file.open(first_file, pwd=password.encode('utf-8') if password else None) as zf:
                content = zf.read()
                
            # 写入临时文件
            with open(temp_output, 'wb') as f:
                f.write(content)
            
            # 检查解压出的文件扩展名
            file_ext = os.path.splitext(first_file)[1].lower()
            final_output = os.path.join(output_dir, f"{base_name}.xlsx")
            
            # 如果目标文件已存在，先删除
            if os.path.exists(final_output):
                os.remove(final_output)
                
            if file_ext != '.xlsx':
                print("解压出的文件是zip文件，继续解压...")
                # 如果是zip文件，继续解压（无需密码）
                with zipfile.ZipFile(temp_output) as inner_zip:
                    with inner_zip.open(inner_zip.namelist()[0]) as inner_file:
                        with open(final_output, 'wb') as f:
                            f.write(inner_file.read())
                # 删除中间zip文件
                os.remove(temp_output)
                return final_output
            else:
                print(f"解压出的文件扩展名为{file_ext}，直接重命名...")
                # 如果不是zip文件，重命名为xlsx
                final_output = os.path.join(output_dir, f"{base_name}.xlsx")
                if temp_output != final_output:
                    os.rename(temp_output, final_output)
                return final_output
                
    except Exception as e:
        raise Exception(f"解压失败: {str(e)}")

def create_password_dialog(title="密码输入", message="请输入ZIP文件密码:", file_path=None):
    root = tk.Tk()
    root.withdraw()
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
            output_file = extract_zip(file_path, password)
            messagebox.showinfo("成功", f"文件已解压为:\n{output_file}")
            root.destroy()
        except Exception as e:
            if 'Bad password' in str(e):
                messagebox.showerror("错误", "密码错误")
            else:
                messagebox.showerror("错误", str(e))
    
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
    file_path = rf"C:\project\python_project\AI+市场工具&稽核\AI+市场工具&稽核第一期需求\选项目-制表-得出结论\25年5月8日数据\制表取数表2.zip"
    create_password_dialog(file_path=file_path)