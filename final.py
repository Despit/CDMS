import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
import os
import re
from datetime import datetime
from docx import Document

# -------------------------- 全局配置 --------------------------
# 字段定义（按你的需求排序，标注必填项）
FIELDS = [
    ("证书编号", "required"), ("送检班组", "required"), ("样品名称", "required"),
    ("生产厂商", "optional"), ("型号", "optional"), ("样品编号", "optional"),
    ("批准人", "optional"), ("核验员", "optional"), ("检定员", "optional"),
    ("校验日期", "optional"), ("有效日期", "optional"), ("安装位置", "optional"),
    ("设备精度", "optional"), ("量程范围", "optional"), ("输出范围", "optional"),
    ("外观检查", "optional"), ("绝缘电阻", "optional"), ("输出纹波含量", "optional"),
    ("输出值1", "optional"), ("输出值2", "optional"), ("输出值3", "optional"),
    ("输出值4", "optional"), ("输出值5", "optional"), ("输出值6", "optional"),
    ("输出值7", "optional"), ("输出值8", "optional"), ("输出值9", "optional"),
    ("输出值10", "optional"), ("输出值11", "optional"), ("误差1", "optional"),
    ("误差2", "optional"), ("误差3", "optional"), ("误差4", "optional"),
    ("误差5", "optional"), ("误差6", "optional"), ("误差7", "optional"),
    ("误差8", "optional"), ("误差9", "optional"), ("误差10", "optional"),
    ("误差11", "optional"), ("最大基本误差", "optional")
]
# 文件路径
EXCEL_PATH = os.path.join(os.path.expanduser("~"), "Desktop", "设备校验数据.xlsx")
TEMPLATE_DEFAULT_PATH = os.path.join(os.path.expanduser("~"), "Desktop", "template.docx")
EXPORT_DIR = os.path.join(os.path.expanduser("~"), "Desktop", "导出Word文档")

# -------------------------- Excel 核心操作 --------------------------
def init_excel():
    """初始化Excel，创建表头"""
    if not os.path.exists(EXCEL_PATH):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "设备校验数据"
        # 写入表头
        headers = [field[0] for field in FIELDS]
        ws.append(headers)
        # 表头样式
        for cell in ws[1]:
            cell.font = openpyxl.styles.Font(bold=True)
            cell.alignment = openpyxl.styles.Alignment(horizontal="center")
        wb.save(EXCEL_PATH)
    return openpyxl.load_workbook(EXCEL_PATH)

def validate_input(input_data):
    """校验必填项"""
    for field, typ in FIELDS:
        if typ == "required" and not input_data[field].strip():
            messagebox.warning("提示", f"{field}为必填项！")
            return False
    return True

def save_data_to_excel(data, row_num=None):
    """保存数据到Excel（新增/修改）"""
    try:
        wb = init_excel()
        ws = wb.active
        # 组装数据行
        row_data = [data[field[0]] for field in FIELDS]
        if row_num is None:
            # 新增数据：追加到最后一行
            ws.append(row_data)
            row_num = ws.max_row
        else:
            # 修改数据：覆盖指定行
            for col, value in enumerate(row_data, 1):
                ws.cell(row=row_num, column=col, value=value)
        # 单元格居中
        for cell in ws[row_num]:
            cell.alignment = openpyxl.styles.Alignment(horizontal="center")
        wb.save(EXCEL_PATH)
        wb.close()
        return True
    except Exception as e:
        messagebox.showerror("错误", f"保存失败：{str(e)}")
        return False

def read_all_data():
    """读取所有数据（返回：表头+数据行）"""
    try:
        wb = init_excel()
        ws = wb.active
        max_row = ws.max_row
        max_col = ws.max_column
        if max_row == 1:
            return [[field[0] for field in FIELDS]]  # 仅返回表头
        # 读取所有行
        data = []
        for row in range(1, max_row + 1):
            row_data = [ws.cell(row=row, column=col).value or "" for col in range(1, max_col + 1)]
            data.append(row_data)
        wb.close()
        return data
    except Exception as e:
        messagebox.showerror("错误", f"读取数据失败：{str(e)}")
        return []

def delete_data(row_num):
    """删除指定行数据"""
    if messagebox.askyesno("确认", "是否删除选中的数据？"):
        try:
            wb = init_excel()
            ws = wb.active
            ws.delete_rows(row_num)
            # 重新排序序号（若需要）
            wb.save(EXCEL_PATH)
            wb.close()
            messagebox.showinfo("成功", "数据已删除！")
            return True
        except Exception as e:
            messagebox.showerror("错误", f"删除失败：{str(e)}")
            return False

# -------------------------- Word 导出操作 --------------------------
def init_export_dir():
    """初始化导出文件夹"""
    if not os.path.exists(EXPORT_DIR):
        os.makedirs(EXPORT_DIR)

def replace_placeholder(doc, data_dict):
    """替换Word模板中的占位符"""
    # 替换段落占位符
    for para in doc.paragraphs:
        for run in para.runs:
            for field in FIELDS:
                placeholder = f"{{{{{field[0]}}}}}"
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, data_dict[field[0]] or "")
    # 替换文本框占位符（可选）
    for shape in doc.element.body.iter():
        if shape.tag.endswith('textbox'):
            for para in shape:
                if hasattr(para, 'runs'):
                    for run in para.runs:
                        for field in FIELDS:
                            placeholder = f"{{{{{field[0]}}}}}"
                            if placeholder in run.text:
                                run.text = run.text.replace(placeholder, data_dict[field[0]] or "")

def export_selected_data(selected_rows):
    """导出选中的数据到Word"""
    # 选择模板文件
    template_path = filedialog.askopenfilename(
        title="选择Word模板",
        filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")],
        initialfile=TEMPLATE_DEFAULT_PATH
    )
    if not template_path:
        return
    # 校验模板是否存在
    if not os.path.exists(template_path):
        messagebox.showerror("错误", "模板文件不存在！")
        return
    # 初始化导出目录
    init_export_dir()
    # 读取Excel数据
    all_data = read_all_data()
    if len(all_data) <= 1:
        messagebox.showinfo("提示", "暂无数据可导出！")
        return
    # 表头映射
    headers = all_data[0]
    success_count = 0
    # 批量导出选中行
    for row_idx in selected_rows:
        if row_idx < 2:  # 跳过表头
            continue
        row_data = all_data[row_idx - 1]  # Excel行号=表格行号+1
        # 转为字典
        data_dict = dict(zip(headers, row_data))
        # 打开模板
        doc = Document(template_path)
        # 替换占位符
        replace_placeholder(doc, data_dict)
        # 生成文件名
        file_name = f"设备校验报告_{data_dict['证书编号']}_{datetime.now().strftime('%Y%m%d%H%M%S')}.docx"
        export_path = os.path.join(EXPORT_DIR, file_name)
        # 保存文档
        try:
            doc.save(export_path)
            success_count += 1
        except Exception as e:
            messagebox.warning(f"导出失败（证书编号：{data_dict['证书编号']}）：{str(e)}")
    # 导出完成提示
    messagebox.showinfo("成功", f"共导出{success_count}个文档！\n保存路径：{EXPORT_DIR}")

# -------------------------- GUI 界面 --------------------------
class DeviceCheckApp:
    def __init__(self, root):
        self.root = root
        self.root.title("设备校验数据收集系统")
        self.root.geometry("1200x700")
        self.root.minsize(1000, 600)

        # 变量初始化
        self.current_edit_row = None  # 当前编辑的行号
        self.selected_items = []      # 选中的行ID

        # 初始化Excel
        init_excel()

        # 布局分栏：左侧录入区，右侧数据展示区
        self.paned = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 1. 左侧录入面板
        self.input_frame = ttk.LabelFrame(self.paned, text="数据录入")
        self.paned.add(self.input_frame, weight=1)
        self.create_input_form()

        # 2. 右侧数据面板
        self.data_frame = ttk.LabelFrame(self.paned, text="历史数据")
        self.paned.add(self.data_frame, weight=2)
        self.create_data_table()

    def create_input_form(self):
        """创建数据录入表单"""
        # 滚动容器（适配多字段）
        canvas = tk.Canvas(self.input_frame)
        scroll_y = ttk.Scrollbar(self.input_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scroll_y.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        # 生成输入框
        self.input_vars = {}
        for idx, (field, typ) in enumerate(FIELDS):
            # 标签
            lbl = ttk.Label(scrollable_frame, text=f"{field}{'*' if typ=='required' else ''}：")
            lbl.grid(row=idx, column=0, padx=10, pady=5, sticky="e")
            # 输入框
            var = tk.StringVar()
            entry = ttk.Entry(scrollable_frame, textvariable=var, width=30)
            entry.grid(row=idx, column=1, padx=10, pady=5, sticky="w")
            self.input_vars[field] = var

        # 操作按钮
        btn_frame = ttk.Frame(scrollable_frame)
        btn_frame.grid(row=len(FIELDS), column=0, columnspan=2, pady=20)
        # 提交按钮
        ttk.Button(btn_frame, text="提交数据", command=self.submit_data).grid(row=0, column=0, padx=10)
        # 清空按钮
        ttk.Button(btn_frame, text="清空表单", command=self.clear_form).grid(row=0, column=1, padx=10)
        # 取消编辑按钮
        ttk.Button(btn_frame, text="取消编辑", command=self.cancel_edit).grid(row=0, column=2, padx=10)

    def create_data_table(self):
        """创建历史数据表格"""
        # 按钮栏
        btn_frame = ttk.Frame(self.data_frame)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        # 修改按钮
        ttk.Button(btn_frame, text="修改选中", command=self.edit_selected).pack(side=tk.LEFT, padx=5)
        # 删除按钮
        ttk.Button(btn_frame, text="删除选中", command=self.delete_selected).pack(side=tk.LEFT, padx=5)
        # 导出按钮
        ttk.Button(btn_frame, text="导出选中到Word", command=self.export_selected).pack(side=tk.LEFT, padx=5)
        # 刷新按钮
        ttk.Button(btn_frame, text="刷新数据", command=self.refresh_table).pack(side=tk.LEFT, padx=5)

        # 表格（带复选框）
        self.tree = ttk.Treeview(self.data_frame, columns=[f"col{i}" for i in range(len(FIELDS))], show="headings", selectmode="extended")
        # 设置表头
        for idx, (field, _) in enumerate(FIELDS):
            self.tree.heading(f"col{idx}", text=field)
            self.tree.column(f"col{idx}", width=80, minwidth=50)
        # 滚动条
        scroll_y = ttk.Scrollbar(self.data_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scroll_x = ttk.Scrollbar(self.data_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        # 布局
        self.tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=5)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        # 绑定选中事件
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        # 初始加载数据
        self.refresh_table()

    def get_input_data(self):
        """获取表单输入数据"""
        data = {}
        for field, _ in FIELDS:
            data[field] = self.input_vars[field].get().strip()
        return data

    def fill_form(self, data):
        """填充表单（编辑时）"""
        self.clear_form()
        for field, _ in FIELDS:
            self.input_vars[field].set(data[field] or "")

    def submit_data(self):
        """提交/修改数据"""
        input_data = self.get_input_data()
        if not validate_input(input_data):
            return
        # 保存数据
        if self.current_edit_row is None:
            # 新增数据
            if save_data_to_excel(input_data):
                messagebox.showinfo("成功", "数据新增成功！")
        else:
            # 修改数据
            if save_data_to_excel(input_data, self.current_edit_row):
                messagebox.showinfo("成功", "数据修改成功！")
                self.current_edit_row = None
        # 清空表单+刷新表格
        self.clear_form()
        self.refresh_table()

    def clear_form(self):
        """清空表单"""
        for var in self.input_vars.values():
            var.set("")

    def cancel_edit(self):
        """取消编辑"""
        self.current_edit_row = None
        self.clear_form()

    def refresh_table(self):
        """刷新数据表格"""
        # 清空现有数据
        for item in self.tree.get_children():
            self.tree.delete(item)
        # 加载新数据
        all_data = read_all_data()
        for row in all_data[1:]:  # 跳过表头
            self.tree.insert("", tk.END, values=row)

    def on_select(self, event):
        """获取选中的行"""
        self.selected_items = self.tree.selection()

    def edit_selected(self):
        """编辑选中的行"""
        if len(self.selected_items) != 1:
            messagebox.warning("提示", "请仅选中一行数据进行修改！")
            return
        # 获取选中行数据
        selected_item = self.selected_items[0]
        row_data = self.tree.item(selected_item)["values"]
        # 转为字典
        data_dict = dict(zip([f[0] for f in FIELDS], row_data))
        # 填充表单
        self.fill_form(data_dict)
        # 获取Excel行号（表格行号=Excel行号-1）
        self.current_edit_row = self.tree.index(selected_item) + 2  # +1跳过表头，+1因为Excel行从1开始

    def delete_selected(self):
        """删除选中的行"""
        if not self.selected_items:
            messagebox.warning("提示", "请选中要删除的数据！")
            return
        # 批量删除
        for item in self.selected_items:
            row_idx = self.tree.index(item) + 2  # Excel行号
            delete_data(row_idx)
        # 刷新表格
        self.refresh_table()

    def export_selected(self):
        """导出选中的数据"""
        if not self.selected_items:
            messagebox.warning("提示", "请选中要导出的数据！")
            return
        # 获取选中行的Excel行号
        selected_rows = [self.tree.index(item) + 2 for item in self.selected_items]
        # 导出Word
        export_selected_data(selected_rows)

# -------------------------- 程序入口 --------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = DeviceCheckApp(root)
    root.mainloop()
