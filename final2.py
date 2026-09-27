import shutil  # 用于复制文件
import tempfile  # 用于创建临时文件（可选）
import pyglet #字体导入

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
import os
import re
from datetime import datetime
from docx import Document
from docx.shared import Inches
import threading

# -------------------------- 全局配置 --------------------------
# 字段定义（按需求排序，标注必填项）
pyglet.font.add_file("./SarasaUiSC-Regular.ttf")  # 更纱黑体 UI SC Regular
pyglet.font.load("./SarasaUiSC-Regular.ttf")
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
EXCEL_PATH = os.path.join(os.path.expanduser("~"), "gitLib", "CalibrationDataManagementSystem", "设备校验数据.xlsx")
TEMPLATE_DEFAULT_PATH = os.path.join(os.path.expanduser("~"), "gitLib", "CalibrationDataManagementSystem", "template.docx")
EXPORT_DIR = os.path.join(os.path.expanduser("~"), "gitLib", "CalibrationDataManagementSystem", "导出Word文档")

# -------------------------- 自定义组件 --------------------------
class PlaceholderEntry(ttk.Entry):
    """带占位符的Entry组件（兼容原生Tkinter）"""
    def __init__(self, master=None, placeholder="", **kwargs):
        super().__init__(master, **kwargs)
        self.placeholder = placeholder
        self.var = kwargs.get("textvariable", tk.StringVar())
        
        # 绑定事件：获取焦点/失去焦点/内容变化
        self.bind("<FocusIn>", self._on_focus_in)
        self.bind("<FocusOut>", self._on_focus_out)
        self.var.trace("w", self._on_var_change)
        
        # 初始化占位符
        self._show_placeholder()

    def _show_placeholder(self):
        """显示占位符"""
        if not self.var.get():
            # 设置占位符样式（灰色）
            self.config(style="Placeholder.TEntry")
            self.insert(0, self.placeholder)

    def _hide_placeholder(self):
        """隐藏占位符"""
        if self.get() == self.placeholder:
            self.delete(0, tk.END)
            # 恢复正常样式
            self.config(style="TEntry")

    def _on_focus_in(self, event):
        """获取焦点时隐藏占位符"""
        self._hide_placeholder()

    def _on_focus_out(self, event):
        """失去焦点时显示占位符（无内容时）"""
        self._show_placeholder()

    def _on_var_change(self, *args):
        """内容变化时同步占位符状态"""
        if self.var.get():
            self._hide_placeholder()
        else:
            self._show_placeholder()

# -------------------------- 工具函数 --------------------------
def safe_open_excel():
    """安全打开Excel（避免文件占用）"""
    try:
        return openpyxl.load_workbook(EXCEL_PATH, read_only=False, data_only=True)
    except PermissionError:
        messagebox.showerror("错误", "Excel文件已被打开，请关闭后重试！")
        return None

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
        header_font = openpyxl.styles.Font(bold=True)
        header_align = openpyxl.styles.Alignment(horizontal="center")
        for cell in ws[1]:
            cell.font = header_font
            cell.alignment = header_align
        # 调整列宽（自适应）
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 20)  # 最大宽度20
            ws.column_dimensions[col_letter].width = adjusted_width
        wb.save(EXCEL_PATH)
        wb.close()
    return safe_open_excel()

def validate_input(input_data):
    """校验必填项"""
    for field, typ in FIELDS:
        if typ == "required" and not input_data[field].strip():
            messagebox.showwarning("提示", f"{field}为必填项！")
            return False
    # 日期格式校验（可选）
    date_fields = ["校验日期", "有效日期"]
    date_pattern = re.compile(r"\d{4}年\d{2}月\d{2}日")
    for field in date_fields:
        val = input_data[field].strip()
        if val and not date_pattern.match(val):
            messagebox.showwarning("提示", f"{field}格式错误（请输入：YYYY年MM月DD日）！")
            return False
    return True

def save_data_to_excel(data, row_num=None):
    """保存数据到Excel（新增/修改）"""
    wb = init_excel()
    if not wb:
        return False
    try:
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
        cell_align = openpyxl.styles.Alignment(horizontal="center")
        for cell in ws[row_num]:
            cell.alignment = cell_align
        wb.save(EXCEL_PATH)
        return True
    except Exception as e:
        messagebox.showerror("错误", f"保存失败：{str(e)}")
        return False
    finally:
        if wb:
            wb.close()

def read_all_data():
    """读取所有数据（返回：表头+数据行）"""
    wb = safe_open_excel()
    if not wb:
        return []
    try:
        ws = wb.active
        max_row = ws.max_row
        max_col = ws.max_column
        if max_row == 1:
            return [[field[0] for field in FIELDS]]  # 仅返回表头
        # 批量读取数据（提升性能）
        data = []
        for row in ws.iter_rows(min_row=1, max_row=max_row, max_col=max_col, values_only=True):
            row_data = [str(cell) if cell is not None else "" for cell in row]
            data.append(row_data)
        return data
    except Exception as e:
        messagebox.showerror("错误", f"读取数据失败：{str(e)}")
        return []
    finally:
        wb.close()

def delete_data(row_num):
    """删除指定行数据"""
    if not messagebox.askyesno("确认", "是否删除选中的数据？"):
        return False
    wb = safe_open_excel()
    if not wb:
        return False
    try:
        ws = wb.active
        ws.delete_rows(row_num)
        wb.save(EXCEL_PATH)
        messagebox.showinfo("成功", "数据已删除！")
        return True
    except Exception as e:
        messagebox.showerror("错误", f"删除失败：{str(e)}")
        return False
    finally:
        if wb:
            wb.close()

# -------------------------- Word 导出操作 --------------------------
def init_export_dir():
    """初始化导出文件夹"""
    if not os.path.exists(EXPORT_DIR):
        os.makedirs(EXPORT_DIR)

def replace_placeholder(doc, data_dict):
    """高效替换占位符（保留所有格式：逐Run替换+兼容页眉/页脚/文本框）"""
    # 预编译占位符正则
    placeholder_pattern = re.compile(r"\{\{(.+?)\}\}")
    
    # 定义通用替换函数（逐Run处理，保留样式）
    def replace_in_runs(runs):
        for run in runs:
            if placeholder_pattern.search(run.text):
                # 逐占位符替换，保留Run的样式（字体/颜色/大小等）
                new_text = placeholder_pattern.sub(lambda m: data_dict.get(m.group(1), ""), run.text)
                run.text = new_text

    # 1. 替换正文段落（逐Run处理）
    for para in doc.paragraphs:
        replace_in_runs(para.runs)

    # 2. 替换表格单元格（逐Run处理）
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    replace_in_runs(para.runs)

    # 3. 替换页眉/页脚（保留页眉页脚格式）
    for section in doc.sections:
        # 页眉
        for header in section.header.paragraphs:
            replace_in_runs(header.runs)
        # 页脚
        for footer in section.footer.paragraphs:
            replace_in_runs(footer.runs)

    # 4. 替换文本框内内容（保留文本框样式）
    for shape in doc.element.body.iter():
        if shape.tag.endswith('textbox'):
            for para in shape:
                if hasattr(para, 'runs'):
                    replace_in_runs(para.runs)
                    
def export_single_data(row_data, headers, template_path, export_path):
    """导出单条数据（保留格式+不修改原始模板）"""
    try:
        # 1. 创建临时模板副本（避免修改原始模板）
        temp_template = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        shutil.copy2(template_path, temp_template.name)  # 保留文件元数据和格式
        temp_template.close()
        
        # 2. 打开临时副本（只读模式避免格式损坏）
        doc = Document(temp_template.name)
        data_dict = dict(zip(headers, row_data))
        
        # 3. 替换占位符（保留所有格式）
        replace_placeholder(doc, data_dict)
        
        # 4. 保存到目标路径（禁用格式自动修正）
        doc.save(export_path)
        
        # 5. 清理临时文件
        os.unlink(temp_template.name)
        return True, data_dict["证书编号"]
    except Exception as e:
        # 异常时清理临时文件
        if 'temp_template' in locals() and os.path.exists(temp_template.name):
            os.unlink(temp_template.name)
        return False, f"{data_dict.get('证书编号', '未知')}：{str(e)}"

def export_selected_data(selected_rows):
    """导出选中的数据到Word（多线程提升效率）"""
    # 选择模板文件
    template_path = filedialog.askopenfilename(
        title="选择Word模板",
        filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")],
        initialfile=TEMPLATE_DEFAULT_PATH
    )
    if not template_path:
        return
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
    # 批量导出（多线程）
    success_count = 0
    fail_info = []
    threads = []
    
    def export_worker():
        nonlocal success_count, fail_info
        for row_idx in selected_rows:
            if row_idx < 2:
                continue
            row_data = all_data[row_idx - 1]
            cert_no = row_data[0]
            file_name = f"设备校验报告_{cert_no}_{datetime.now().strftime('%Y%m%d%H%M%S')}.docx"
            export_path = os.path.join(EXPORT_DIR, file_name)
            # 导出单条数据
            res, msg = export_single_data(row_data, headers, template_path, export_path)
            if res:
                success_count += 1
            else:
                fail_info.append(msg)
    
    # 启动导出线程（避免界面卡顿）
    export_thread = threading.Thread(target=export_worker)
    export_thread.daemon = True
    export_thread.start()
    # 等待线程完成
    export_thread.join()
    
    # 导出完成提示
    tip_msg = f"共导出{success_count}个文档！\n保存路径：{EXPORT_DIR}"
    if fail_info:
        tip_msg += f"\n失败项：\n{chr(10).join(fail_info)}"
    messagebox.showinfo("导出完成", tip_msg)

# -------------------------- GUI 界面 --------------------------
class DeviceCheckApp:
    def __init__(self, root):
        self.root = root
        self.root.title("设备校验数据收集系统")
        self.root.geometry("1500x800")
        self.root.minsize(1000, 600)

        # 变量初始化
        self.current_edit_row = None  # 当前编辑的行号
        self.selected_items = []      # 选中的行ID

        # 初始化Excel
        init_excel()

        # 样式配置
        self.style = ttk.Style()
        self.style.configure("Treeview.Heading", font=("SimHei", 9, "bold"))
        self.style.configure("Treeview", font=("SimHei", 9), rowheight=25)
        # 新增：占位符样式（灰色文字）
        self.style.configure("Placeholder.TEntry", foreground="#999999")

        # 布局分栏：左侧录入区，右侧数据展示区
        self.paned = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 1. 左侧录入面板
        self.input_frame = ttk.LabelFrame(self.paned, text="数据录入", padding=10)
        self.paned.add(self.input_frame, weight=1)
        self.create_input_form()

        # 2. 右侧数据面板
        self.data_frame = ttk.LabelFrame(self.paned, text="历史数据", padding=10)
        self.paned.add(self.data_frame, weight=3)
        self.create_data_table()

    def create_input_form(self):
        """创建数据录入表单（优化布局）"""
        # 滚动容器（适配多字段）
        canvas = tk.Canvas(self.input_frame, bg="black")
        scroll_y = ttk.Scrollbar(self.input_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scroll_y.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        # 生成输入框（分两列布局，减少滚动）
        self.input_vars = {}
        field_count = len(FIELDS)
        col1_fields = FIELDS[:field_count//2]
        col2_fields = FIELDS[field_count//2:]
        
        # 第一列
        for idx, (field, typ) in enumerate(col1_fields):
            lbl = ttk.Label(scrollable_frame, text=f"{field}{'*' if typ=='required' else ''}：")
            lbl.grid(row=idx, column=0, padx=10, pady=5, sticky="e")
            var = tk.StringVar()
            entry = ttk.Entry(scrollable_frame, textvariable=var, width=25)
            entry.grid(row=idx, column=1, padx=10, pady=5, sticky="w")
            self.input_vars[field] = var
        
        # 第二列
        for idx, (field, typ) in enumerate(col2_fields):
            lbl = ttk.Label(scrollable_frame, text=f"{field}{'*' if typ=='required' else ''}：")
            lbl.grid(row=idx, column=2, padx=10, pady=5, sticky="e")
            var = tk.StringVar()
            entry = ttk.Entry(scrollable_frame, textvariable=var, width=25)
            entry.grid(row=idx, column=3, padx=10, pady=5, sticky="w")
            self.input_vars[field] = var

        # 操作按钮（居中）
        btn_frame = ttk.Frame(scrollable_frame)
        max_row = max(len(col1_fields), len(col2_fields))
        btn_frame.grid(row=max_row, column=0, columnspan=4, pady=20)
        # 提交按钮
        ttk.Button(btn_frame, text="提交数据", command=self.submit_data, width=12).grid(row=0, column=0, padx=8)
        # 清空按钮
        ttk.Button(btn_frame, text="清空表单", command=self.clear_form, width=12).grid(row=0, column=1, padx=8)
        # 取消编辑按钮
        ttk.Button(btn_frame, text="取消编辑", command=self.cancel_edit, width=12).grid(row=0, column=2, padx=8)

    def create_data_table(self):
        """创建历史数据表格（优化交互）"""
        # 按钮栏（带分隔）
        btn_frame = ttk.Frame(self.data_frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)
        # 修改按钮
        ttk.Button(btn_frame, text="修改选中", command=self.edit_selected, width=12).pack(side=tk.LEFT, padx=5)
        # 删除按钮
        ttk.Button(btn_frame, text="删除选中", command=self.delete_selected, width=12).pack(side=tk.LEFT, padx=5)
        # 导出按钮
        ttk.Button(btn_frame, text="导出选中到Word", command=self.export_selected, width=15).pack(side=tk.LEFT, padx=5)
        # 刷新按钮
        ttk.Button(btn_frame, text="刷新数据", command=self.refresh_table, width=12).pack(side=tk.LEFT, padx=5)
        # 搜索框（右侧）
        search_var = tk.StringVar()
        search_var.trace("w", lambda *args: self.filter_table(search_var.get()))
        # 替换为自定义占位符输入框
        search_entry = PlaceholderEntry(btn_frame, textvariable=search_var, width=20, placeholder="搜索证书编号/样品名称")
        search_entry.pack(side=tk.RIGHT, padx=5)
        ttk.Label(btn_frame, text="搜索：").pack(side=tk.RIGHT)

        # 表格（带复选框，优化列宽）
        self.tree = ttk.Treeview(self.data_frame, columns=[f"col{i}" for i in range(len(FIELDS))], show="headings", selectmode="extended")
        # 设置表头和列宽
        for idx, (field, _) in enumerate(FIELDS):
            self.tree.heading(f"col{idx}", text=field)
            # 关键字段加宽，其他字段适配
            if field in ["证书编号", "样品名称", "生产厂商"]:
                self.tree.column(f"col{idx}", width=120, minwidth=80)
            else:
                self.tree.column(f"col{idx}", width=70, minwidth=50)
        # 滚动条
        scroll_y = ttk.Scrollbar(self.data_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scroll_x = ttk.Scrollbar(self.data_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        # 布局
        self.tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=5)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        # 绑定事件
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        # 初始加载数据
        self.refresh_table()
        self.all_table_data = read_all_data()  # 缓存全量数据

    def filter_table(self, keyword):
        """表格数据筛选"""
        if not keyword:
            self.refresh_table()
            return
        # 筛选包含关键词的行（证书编号/样品名称）
        filtered_data = [self.all_table_data[0]]  # 保留表头
        for row in self.all_table_data[1:]:
            if keyword.lower() in row[0].lower() or keyword.lower() in row[2].lower():
                filtered_data.append(row)
        # 更新表格
        for item in self.tree.get_children():
            self.tree.delete(item)
        for row in filtered_data[1:]:
            self.tree.insert("", tk.END, values=row)

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
        self.all_table_data = read_all_data()
        for row in self.all_table_data[1:]:
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
        # 获取Excel行号
        self.current_edit_row = self.tree.index(selected_item) + 2

    def delete_selected(self):
        """删除选中的行"""
        if not self.selected_items:
            messagebox.warning("提示", "请选中要删除的数据！")
            return
        # 批量删除
        for item in self.selected_items:
            row_idx = self.tree.index(item) + 2
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
        # 导出Word（异步执行，避免界面卡顿）
        export_thread = threading.Thread(target=export_selected_data, args=(selected_rows,))
        export_thread.daemon = True
        export_thread.start()

# -------------------------- 程序入口 --------------------------
if __name__ == "__main__":
    # 解决Tkinter中文显示问题
    root = tk.Tk()
    root.option_add("*Font", "Sarasa-UI-SC 9")
    app = DeviceCheckApp(root)
    root.mainloop()
