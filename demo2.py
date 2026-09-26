# 导入Python内置的Tkinter库（无需安装）
import tkinter as tk

# 定义主窗口函数
def create_main_window():
    # 1. 创建主窗口对象（程序核心窗口）
    root = tk.Tk()
    # 设置窗口标题
    root.title("简易关闭按钮示例")
    # 设置窗口大小（宽x高），可根据需要调整
    root.geometry("300x150")
    
    # 2. 定义关闭函数（点击按钮时执行）
    def close_window():
        # 直接销毁主窗口，退出程序
        root.destroy()
    
    # 3. 创建关闭按钮
    # button参数说明：
    # - text：按钮显示的文字
    # - command：点击按钮触发的函数（这里绑定close_window）
    # - font：字体和大小（提升显示效果）
    # - bg：按钮背景色，fg：文字颜色（可选美化）
    close_btn = tk.Button(
        root,
        text="关闭窗口",
        command=close_window,
        font=("微软雅黑", 12),
        bg="#ff4444",
        fg="white",
        width=10  # 按钮宽度
    )
    
    # 4. 放置按钮（居中显示，上下留空）
    close_btn.pack(pady=50)  # pady=50：上下各留50像素间距
    
    # 5. 启动窗口主循环（保持窗口显示，等待用户操作）
    root.mainloop()

# 程序入口（IDLE运行时执行）
if __name__ == "__main__":
    create_main_window()
