import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
import os
import re
from datetime import datetime
from docx import Document
from docx.shared import Inches
import threading


root=tk.Tk()

root.title("liuCc")
root.geometry("1000x200")
root.resizable(True,True)
root.attributes("-alpha",1)
label = tk.Label(root, text="请输入内容：", font=("Arial", 12))
label.pack(pady=5)
entry = tk.Entry(root, width=10, font=("Arial", 12))
entry.pack(pady=50)

root.mainloop()#必须放置于最后
