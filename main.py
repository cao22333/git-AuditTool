import tkinter as tk
from ui.main_ui import AuditToolApp # 从ui.main_ui导入主应用类

def main():
    """
    应用程序主入口点。
    初始化 Tkinter 根窗口并启动 AuditToolApp。
    """
    root = tk.Tk()
    app = AuditToolApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()