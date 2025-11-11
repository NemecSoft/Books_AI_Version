import tkinter as tk
from tkinter import ttk

# 简单的GUI测试程序
class SimpleGUIApp:
    def __init__(self, root):
        self.root = root
        self.root.title("GUI测试程序")
        self.root.geometry("400x300")
        
        # 设置中文字体
        self.style = ttk.Style()
        try:
            self.style.configure("TButton", font=('SimHei', 10))
            self.style.configure("TLabel", font=('SimHei', 10))
            print("成功设置中文字体")
        except Exception as e:
            print(f"设置字体时出错: {e}")
        
        # 创建简单的界面元素
        self.create_widgets()
    
    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题标签
        title_label = ttk.Label(main_frame, text="GUI测试成功", font=('SimHei', 16))
        title_label.pack(pady=20)
        
        # 状态标签
        status_label = ttk.Label(main_frame, text="窗口已正常显示")
        status_label.pack(pady=10)
        
        # 退出按钮
        exit_btn = ttk.Button(main_frame, text="退出", command=self.root.destroy)
        exit_btn.pack(pady=20)

if __name__ == "__main__":
    print("正在创建GUI窗口...")
    root = tk.Tk()
    print("Tk实例已创建")
    app = SimpleGUIApp(root)
    print("App实例已创建，启动主循环")
    root.mainloop()
    print("GUI已关闭")