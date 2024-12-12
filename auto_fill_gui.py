import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import sys
from auto_fill import read_personal_info, fill_document
import webbrowser

class AutoFillApp:
    def __init__(self, root):
        self.root = root
        self.root.title("文档自动填写工具 v1.0")
        
        # 设置窗口大小和位置
        window_width = 800
        window_height = 600
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        # 设置样式
        style = ttk.Style()
        style.configure('Title.TLabel', font=('微软雅黑', 12, 'bold'))
        style.configure('Info.TLabel', font=('微软雅黑', 10))
        style.configure('Action.TButton', font=('微软雅黑', 10))
        
        # 创建主框架
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题和说明
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(title_frame, text="文档自动填写工具", 
                 style='Title.TLabel').pack(side=tk.LEFT)
        ttk.Button(title_frame, text="使用说明", 
                  command=self.show_help, style='Action.TButton').pack(side=tk.RIGHT)
        
        # 信息文件选择
        info_frame = ttk.LabelFrame(main_frame, text="步骤1：选择个人信息文件", padding=10)
        info_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.info_path = tk.StringVar()
        ttk.Entry(info_frame, textvariable=self.info_path, width=70).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(info_frame, text="浏览", command=self.select_info_file, 
                  style='Action.TButton').pack(side=tk.LEFT)
        
        # 添加信息文件格式说明
        ttk.Label(info_frame, text='（请选择包含个人信息的.docx文件，格式为"键: 值"）', 
                 style='Info.TLabel').pack(side=tk.LEFT, padx=(10, 0))
        
        # 目标文件选择
        target_frame = ttk.LabelFrame(main_frame, text="步骤2：选择需要填写的文件（最多10个）", padding=10)
        target_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 添加文件列表显示
        list_frame = ttk.Frame(target_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建文件列表
        self.file_listbox = tk.Listbox(list_frame, height=6)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox.configure(yscrollcommand=scrollbar.set)
        
        # 文件操作按钮
        button_frame = ttk.Frame(target_frame)
        button_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(button_frame, text="添加文件", 
                  command=self.add_target_files, 
                  style='Action.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="移除选中", 
                  command=self.remove_selected_file, 
                  style='Action.TButton').pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="清空列表", 
                  command=self.clear_file_list, 
                  style='Action.TButton').pack(side=tk.LEFT, padx=5)
        
        # 操作按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=20)
        
        ttk.Button(button_frame, text="开始填写", command=self.process_files,
                  style='Action.TButton', width=20).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="清除日志", command=self.clear_log,
                  style='Action.TButton', width=20).pack(side=tk.LEFT, padx=5)
        
        # 日志显示区域
        log_frame = ttk.LabelFrame(main_frame, text="处理日志", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, font=('微软雅黑', 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                             relief=tk.SUNKEN, padding=(5, 2))
        status_bar.pack(fill=tk.X, pady=(10, 0))
        
        # 显示欢迎信息
        self.show_welcome_message()
        
    def show_welcome_message(self):
        welcome_msg = """欢迎使用文档自动填写工具！

使用步骤：
1. 选择包含个人信息的Word文档（.docx格式）
2. 添加需要填写的目标文档（最多10个）
3. 点击"开始填写"按钮

注意事项：
- 个人信息文档中的内容需要采用"键: 值"的格式
- 确保Word程序未在运行中
- 填写完成的文件会保存在原文件同目录下，文件名前加"filled_"
"""
        self.log(welcome_msg)
        
    def show_help(self):
        help_text = """个人信息文件格式示例：

姓名: 张三
性别: 男
出生日期: 1990年1月1日
工作单位: XX大学
学历: 研究生
专业: 计算机科学
...

注意事项：
1. 每行一个信息项
2. 使用冒号(:)分隔键和值
3. 键值对之间要换行
4. 保存为.docx格式

如需帮助，请联系管理员。"""
        
        help_window = tk.Toplevel(self.root)
        help_window.title("使用说明")
        help_window.geometry("500x400")
        
        text = scrolledtext.ScrolledText(help_window, wrap=tk.WORD, 
                                       font=('微软雅黑', 10), padding=10)
        text.pack(fill=tk.BOTH, expand=True)
        text.insert(tk.END, help_text)
        text.config(state=tk.DISABLED)
        
    def select_info_file(self):
        filename = filedialog.askopenfilename(
            title="选择个人信息文件",
            filetypes=[("Word文档", "*.docx")],
            initialdir=os.path.dirname(self.info_path.get()) if self.info_path.get() else None
        )
        if filename:
            self.info_path.set(filename)
            self.status_var.set(f"已选择信息文件: {os.path.basename(filename)}")
            
    def add_target_files(self):
        """添加目标文件"""
        current_files = list(self.file_listbox.get(0, tk.END))
        if len(current_files) >= 10:
            messagebox.showwarning("警告", "最多只能添加10个文件！")
            return
            
        filenames = filedialog.askopenfilenames(
            title="选择需要填写的文件",
            filetypes=[("Word文档", "*.doc;*.docx")],
            initialdir=os.path.dirname(self.info_path.get()) if self.info_path.get() else None
        )
        
        for filename in filenames:
            if filename not in current_files:
                if len(current_files) + 1 <= 10:
                    self.file_listbox.insert(tk.END, filename)
                    current_files.append(filename)
                else:
                    messagebox.showwarning("警告", "已达到最大文件数限制！")
                    break
        
        self.status_var.set(f"已选择 {len(current_files)} 个文件")
    
    def remove_selected_file(self):
        """移除选中的文件"""
        selection = self.file_listbox.curselection()
        if selection:
            self.file_listbox.delete(selection)
            self.status_var.set(f"已选择 {self.file_listbox.size()} 个文件")
    
    def clear_file_list(self):
        """清空文件列表"""
        self.file_listbox.delete(0, tk.END)
        self.status_var.set("已清空文件列表")
    
    def clear_log(self):
        self.log_text.delete(1.0, tk.END)
        self.show_welcome_message()
            
    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()
        
    def process_files(self):
        """处理文件"""
        info_file = self.info_path.get()
        target_files = list(self.file_listbox.get(0, tk.END))
        
        if not info_file:
            messagebox.showerror("错误", "请选择个人信息文件！")
            return
            
        if not target_files:
            messagebox.showerror("错误", "请选择需要填写的文件！")
            return
            
        try:
            self.status_var.set("正在处理...")
            self.log("\n" + "="*50)
            self.log("开始处理...")
            
            self.log("正在读取个人信息...")
            info_dict = read_personal_info(info_file)
            if not info_dict:
                self.log("读取个人信息失败！")
                self.status_var.set("处理失败")
                return
            
            # 显示处理进度
            total = len(target_files)
            for i, target_file in enumerate(target_files, 1):
                self.log(f"\n正在处理 ({i}/{total}): {os.path.basename(target_file)}")
                try:
                    fill_document(target_file, info_dict)
                    self.log(f"完成: {os.path.basename(target_file)}")
                except Exception as e:
                    self.log(f"处理失败: {str(e)}")
            
            self.log("\n所有文件处理完成！")
            self.status_var.set("处理完成")
            messagebox.showinfo("完成", f"已完成 {total} 个文件的处理！")
            
        except Exception as e:
            self.log(f"发生错误: {str(e)}")
            self.status_var.set("处理出错")
            messagebox.showerror("错误", f"处理过程中发生错误：{str(e)}")

def main():
    root = tk.Tk()
    app = AutoFillApp(root)
    root.mainloop()

if __name__ == "__main__":
    main() 
