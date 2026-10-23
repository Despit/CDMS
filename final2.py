import shutil  # 用于复制文件
import tempfile  # 用于创建临时文件（可选）

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, messagebox, filedialog
import openpyxl
import os
import re
from datetime import date, datetime
from concurrent.futures import ThreadPoolExecutor
from docx import Document
from docx.shared import Inches

import mysql.connector
from mysql.connector import Error
# -------------------------- 全局配置 --------------------------
# 字段定义（按需求排序，标注必填项）

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

PAGE_SIZE = 15
DATABASE_LIST_FIELDS = ("证书编号", "送检班组", "样品名称", "样品编号")

            
# 文件路径
EXCEL_PATH = os.path.join(os.path.expanduser("~"), "gitLib", "CalibrationDataManagementSystem", "设备校验数据.xlsx")
TEMPLATE_DEFAULT_PATH = os.path.join(os.path.expanduser("~"), "gitLib", "CalibrationDataManagementSystem", "template.docx")
EXPORT_DIR = os.path.join(os.path.expanduser("~"), "gitLib", "CalibrationDataManagementSystem", "导出Word文档")

# -------------------------- 自定义组件 --------------------------
class PlaceholderEntry(ttk.Entry):
    """带占位符的Entry组件（兼容原生Tkinter）"""
    def __init__(self, master=None, placeholder="", **kwargs):
        self.var = kwargs.get("textvariable")
        if self.var is None:
            self.var = tk.StringVar(master=master)
            kwargs["textvariable"] = self.var
        super().__init__(master, **kwargs)
        self.placeholder = placeholder
        self._placeholder_visible = False
        self._has_focus = False
        
        # 绑定事件：获取焦点/失去焦点/内容变化
        self.bind("<FocusIn>", self._on_focus_in)
        self.bind("<FocusOut>", self._on_focus_out)
        self.var.trace_add("write", self._on_var_change)
        
        # 初始化占位符
        self._show_placeholder()

    def _show_placeholder(self):
        """显示占位符"""
        if not self.var.get() and not self._has_focus and not self._placeholder_visible:
            # 占位文字不写入绑定变量，避免触发搜索等业务逻辑。
            self.configure(textvariable="")
            self.insert(0, self.placeholder)
            self.config(style="Placeholder.TEntry")
            self._placeholder_visible = True

    def _hide_placeholder(self):
        """隐藏占位符"""
        if self._placeholder_visible:
            self.delete(0, tk.END)
            self.configure(textvariable=self.var)
            self.config(style="TEntry")
            self._placeholder_visible = False

    def _on_focus_in(self, event):
        """获取焦点时隐藏占位符"""
        self._has_focus = True
        self._hide_placeholder()

    def _on_focus_out(self, event):
        """失去焦点时显示占位符（无内容时）"""
        self._has_focus = False
        self._show_placeholder()

    def _on_var_change(self, *args):
        """内容变化时同步占位符状态"""
        if self.var.get():
            self._hide_placeholder()
        elif not self._has_focus:
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

def validate_input(input_data, require_sample_no=False):
    """校验必填项"""
    for field, typ in FIELDS:
        if typ == "required" and not input_data[field].strip():
            messagebox.showwarning("提示", f"{field}为必填项！")
            return False
    if require_sample_no and not input_data["样品编号"].strip():
        messagebox.showwarning("提示", "样品编号为必填项！")
        return False
    # 日期格式校验（可选）
    # date_fields = ["校验日期", "有效日期"]
    # date_pattern = re.compile(r"\d{4}年\d{2}月\d{2}日")
    # for field in date_fields:
    #     val = input_data[field].strip()
    #     if val and not date_pattern.match(val):
    #         messagebox.showwarning("提示", f"{field}格式错误（请输入：YYYY年MM月DD日）！")
    #         return False
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

def connect_to_database(user=None, password=None):
    """使用环境变量配置连接 MySQL。"""
    user = user or os.getenv("MYSQL_USER", "")
    password = password or os.getenv("MYSQL_PASSWORD", "")
    if not user or not password:
        print("MySQL connection failed: set MYSQL_USER and MYSQL_PASSWORD")
        return None
    try:
        connection = mysql.connector.connect(
            host=os.getenv("MYSQL_HOST", "139.155.96.177"),
            port=int(os.getenv("MYSQL_PORT", "3306")),
            user=user,
            password=password,
            database=os.getenv("MYSQL_DATABASE", "CMDSystem"),
        )
        if connection.is_connected():
            return connection
        connection.close()
    except (Error, ValueError) as error:
        print(f"MySQL connection failed: {error}")
    return None


def execute_command(connection, command, parameters=None):
    """执行 SQL；写操作成功后提交，调用方负责关闭游标。"""
    cursor = None
    try:
        cursor = connection.cursor(buffered=True)
        cursor.execute(command, parameters)
        if not cursor.with_rows:
            connection.commit()
        return cursor
    except Error as error:
        print(f"Failed to execute command: {error}")
        if cursor:
            cursor.close()
        return None


def UpdateTablefromDatabase(limit=PAGE_SIZE, offset=0, keyword="", user=None, password=None):
    """分页读取列表字段和隐藏主键字段，额外读取一条判断是否有下一页。"""
    headers = [field[0] for field in FIELDS]
    data = [headers]
    connection = connect_to_database(user, password)
    if not connection:
        return data, False

    cursor = None
    columns = ", ".join(f"`{field}`" for field in DATABASE_LIST_FIELDS)
    conditions = ["`证书编号` IS NOT NULL"]
    parameters = []
    if keyword:
        conditions.append("(`证书编号` LIKE %s OR `样品名称` LIKE %s)")
        pattern = f"%{keyword}%"
        parameters.extend([pattern, pattern])
    parameters.extend([limit + 1, offset])
    try:
        cursor = connection.cursor(buffered=True)
        cursor.execute(
            f"SELECT {columns} FROM `设备校验数据` "
            f"WHERE {' AND '.join(conditions)} "
            "ORDER BY `证书编号`, `样品编号` LIMIT %s OFFSET %s",
            parameters,
        )
        rows = cursor.fetchall()
        has_more = len(rows) > limit
        field_indexes = {field: headers.index(field) for field in DATABASE_LIST_FIELDS}
        for row in rows[:limit]:
            full_row = [""] * len(headers)
            for field, value in zip(DATABASE_LIST_FIELDS, row):
                full_row[field_indexes[field]] = str(value) if value is not None else ""
            data.append(full_row)
    except Error as error:
        print(f"Failed to read MySQL data: {error}")
        has_more = False
    finally:
        if cursor:
            cursor.close()
        connection.close()
    return data, has_more


def get_data_from_database(primary_key, user=None, password=None):
    """按复合主键读取一条完整记录，用于编辑和完整报告导出。"""
    headers = [field[0] for field in FIELDS]
    columns = ", ".join(f"`{header}`" for header in headers)
    connection = connect_to_database(user, password)
    if not connection:
        return None, "无法连接数据库，请重新登录。"

    cursor = None
    try:
        cursor = connection.cursor(buffered=True)
        cursor.execute(
            f"SELECT {columns} FROM `设备校验数据` "
            "WHERE `证书编号` = %s AND `样品编号` = %s LIMIT 1",
            primary_key,
        )
        row = cursor.fetchone()
        if row is None:
            return None, "未找到对应的数据库记录。"
        return [str(value) if value is not None else "" for value in row], ""
    except Error as error:
        return None, f"读取完整记录失败：{error}"
    finally:
        if cursor:
            cursor.close()
        connection.close()


def save_data_to_database(data, original_primary_key=None, user=None, password=None):
    """新增或更新记录，并校验复合主键（证书编号、样品编号）冲突。"""
    headers = [field[0] for field in FIELDS]
    values = [data[header] for header in headers]
    cert_no = data["证书编号"]
    sample_no = data["样品编号"]
    if not cert_no or not sample_no:
        return False, "证书编号和样品编号均为必填项！"

    connection = connect_to_database(user, password)
    if not connection:
        return False, "无法连接数据库，请重新登录。"

    cursor = None
    try:
        cursor = connection.cursor()
        if original_primary_key is not None:
            cursor.execute(
                "SELECT 1 FROM `设备校验数据` "
                "WHERE `证书编号` = %s AND `样品编号` = %s LIMIT 1",
                original_primary_key,
            )
            if cursor.fetchone() is None:
                return False, "未找到要修改的数据库记录。"

        conflict_query = (
            "SELECT 1 FROM `设备校验数据` "
            "WHERE `证书编号` = %s AND `样品编号` = %s"
        )
        conflict_parameters = [cert_no, sample_no]
        if original_primary_key is not None:
            conflict_query += " AND NOT (`证书编号` = %s AND `样品编号` = %s)"
            conflict_parameters.extend(original_primary_key)
        conflict_query += " LIMIT 1"
        cursor.execute(conflict_query, conflict_parameters)
        if cursor.fetchone() is not None:
            return False, f"主键冲突：证书编号“{cert_no}”与样品编号“{sample_no}”的组合已存在。"

        if original_primary_key is None:
            columns = ", ".join(f"`{header}`" for header in headers)
            placeholders = ", ".join(["%s"] * len(headers))
            cursor.execute(
                f"INSERT INTO `设备校验数据` ({columns}) VALUES ({placeholders})",
                values,
            )
        else:
            assignments = ", ".join(f"`{header}` = %s" for header in headers)
            cursor.execute(
                f"UPDATE `设备校验数据` SET {assignments} "
                "WHERE `证书编号` = %s AND `样品编号` = %s",
                values + list(original_primary_key),
            )
        connection.commit()
        return True, ""
    except Error as error:
        connection.rollback()
        return False, f"数据库保存失败：{error}"
    finally:
        if cursor:
            cursor.close()
        connection.close()


def delete_data_from_database(cert_no, sample_no, user=None, password=None):
    """按复合主键删除数据库记录。"""
    connection = connect_to_database(user, password)
    if not connection:
        return False, "无法连接数据库，请重新登录。"

    cursor = None
    try:
        cursor = connection.cursor()
        cursor.execute(
            "DELETE FROM `设备校验数据` "
            "WHERE `证书编号` = %s AND `样品编号` = %s",
            (cert_no, sample_no),
        )
        if cursor.rowcount == 0:
            return False, "未找到要删除的数据库记录。"
        connection.commit()
        return True, ""
    except Error as error:
        connection.rollback()
        return False, f"数据库删除失败：{error}"
    finally:
        if cursor:
            cursor.close()
        connection.close()




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
        while runs:
            full_text = "".join(run.text for run in runs)
            matches = list(placeholder_pattern.finditer(full_text))
            if not matches:
                return

            match = matches[-1]
            replacement = str(data_dict.get(match.group(1), ""))

            def locate_run(offset, end_offset=False):
                position = 0
                for index, run in enumerate(runs):
                    next_position = position + len(run.text)
                    if offset < next_position or (end_offset and offset == next_position):
                        return index, offset - position
                    position = next_position
                return len(runs) - 1, len(runs[-1].text)

            start_index, start_offset = locate_run(match.start())
            end_index, end_offset = locate_run(match.end(), end_offset=True)
            if start_index == end_index:
                text = runs[start_index].text
                runs[start_index].text = (
                    text[:start_offset] + replacement + text[end_offset:]
                )
                continue

            start_text = runs[start_index].text
            end_text = runs[end_index].text
            runs[start_index].text = start_text[:start_offset] + replacement
            for index in range(start_index + 1, end_index):
                runs[index].text = ""
            runs[end_index].text = end_text[end_offset:]

    # 1. 替换正文段落（逐Run处理）
    for para in doc.paragraphs:
        replace_in_runs(para.runs)

    # 2. 替换表格单元格（逐Run处理）
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    replace_in_runs(para.runs)

    # 替换页眉和页脚
    for section in doc.sections:
        for header in section.header.paragraphs:
            replace_in_runs(header.runs)
        for footer in section.footer.paragraphs:
            replace_in_runs(footer.runs)

    # 替换文本框内内容
    for shape in doc.element.body.iter():
        if shape.tag.endswith("textbox"):
            for para in shape:
                if hasattr(para, "runs"):
                    replace_in_runs(para.runs)


def format_word_date(value):
    """将数据库或 Excel 日期格式化为 Word 使用的中文日期。"""
    if value is None or value == "":
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y年%m月%d日")

    text = str(value).strip()
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        try:
            parsed = date.fromisoformat(text)
        except ValueError:
            return text
    return parsed.strftime("%Y年%m月%d日")
                    
def export_single_data(row_data, headers, template_path, export_path):
    """导出单条数据（保留格式+不修改原始模板）"""
    try:
        # 1. 创建临时模板副本（避免修改原始模板）
        temp_template = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        shutil.copy2(template_path, temp_template.name)  # 保留文件元数据和格式
        temp_template.close()
        
        # 2. 打开临时副本（只读模式避免格式损坏）
        doc = Document(temp_template.name)
        export_row = list(row_data)
        for field in ("校验日期", "有效日期"):
            if field in headers:
                index = headers.index(field)
                export_row[index] = format_word_date(export_row[index])
        data_dict = dict(zip(headers, export_row))
        
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

def export_selected_data(selected_data, headers, template_path):
    """将传入的选中记录导出到 Word，不依赖数据来自 Excel 还是数据库。"""
    init_export_dir()
    success_count = 0
    fail_info = []
    for row_data in selected_data:
        cert_no = row_data[0]
        file_name = f"设备校验报告_{cert_no}_{datetime.now().strftime('%Y%m%d%H%M%S%f')}.docx"
        export_path = os.path.join(EXPORT_DIR, file_name)
        result, message = export_single_data(row_data, headers, template_path, export_path)
        if result:
            success_count += 1
        else:
            fail_info.append(message)
    return success_count, fail_info

# -------------------------- GUI 界面 --------------------------
class DeviceCheckApp:
    def __init__(self, root):
        self.root = root
        tkFont=tkfont.Font(family="更纱黑体 UI SC", size=9)
        root.option_add("*Font", tkFont)
        self.root.title("设备校验数据收集系统")
        self.root.geometry("1920x1080")

        # 变量初始化
        self.current_edit_row = None  # 当前编辑的行号
        self.current_edit_cert_no = None
        self.current_edit_primary_key = None
        self.selected_items = []      # 选中的行ID
        self.localMode = os.getenv("LOCAL_MODE", "0").strip().lower() in {"1", "true", "yes"}
        self.db_user = None
        self.db_password = None
        self.db_logged_in = self.localMode
        self.db_executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="database")
        self.export_executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="word-export")
        self._db_busy = False
        self._pending_refresh = False
        self._db_poll_id = None
        self._export_busy = False
        self._export_poll_id = None
        self._closing = False
        self.page_size = PAGE_SIZE
        self.current_page = 0
        self.has_next_page = False
        self.search_keyword = ""
        self._search_after_id = None

        if self.localMode:
            init_excel()

        # 样式配置
        self.style = ttk.Style()
        self.style.configure("Treeview.Heading", font=("更纱黑体 UI SC", 9, "bold"))
        self.style.configure("Treeview", font=("更纱黑体 UI SC", 9), rowheight=25)
        # 新增：占位符样式（灰色文字）
        self.style.configure("Placeholder.TEntry", foreground="#999999")

        # 布局分栏：左侧录入区，右侧数据展示区
        self.paned = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 1. 左侧录入面板
        self.input_frame = ttk.LabelFrame(self.paned, text="数据录入", padding=10,width=10)
        self.paned.add(self.input_frame, weight=1)
        self.input_frame.configure(width=600)
        self.create_input_form()

        # 2. 右侧数据面板
        self.data_frame = ttk.LabelFrame(self.paned, text="历史数据", padding=10,width=10)
        self.paned.add(self.data_frame, weight=3)
        self.create_data_table()

        self.root.minsize(1100, 700)
        self._pane_constraint_after_id = None
        self.paned.bind("<ButtonRelease-1>", self._schedule_pane_constraint, add="+")
        self.root.bind("<Configure>", self._schedule_pane_constraint, add="+")
        self.root.after_idle(self._constrain_input_panel)
        self.root.protocol("WM_DELETE_WINDOW", self.close)

        # 初始加载数据
        self.refresh_table()

    def create_input_form(self):
        """创建数据录入表单（优化布局）"""
        form_frame = ttk.Frame(self.input_frame)
        form_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        canvas = tk.Canvas(
            form_frame,
            highlightthickness=0,
            background=self.root.cget("background"),
        )
        scroll_y = ttk.Scrollbar(form_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.bind(
            "<Configure>",
            lambda event: canvas.itemconfigure(canvas_window, width=event.width),
        )
        canvas.configure(yscrollcommand=scroll_y.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

        login_frame = ttk.LabelFrame(self.input_frame, text="数据库登录", padding=8)
        login_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=2, pady=(6, 0))
        self.db_username_var = tk.StringVar()
        self.db_password_var = tk.StringVar()
        ttk.Label(login_frame, text="用户名").grid(row=0, column=0, padx=(0, 4), pady=3)
        self.db_username_entry = ttk.Entry(
            login_frame, textvariable=self.db_username_var, width=16
        )
        self.db_username_entry.grid(row=0, column=1, padx=(0, 8), pady=3, sticky="ew")
        ttk.Label(login_frame, text="密码").grid(row=0, column=2, padx=(0, 4), pady=3)
        self.db_password_entry = ttk.Entry(
            login_frame, textvariable=self.db_password_var, width=16, show="*"
        )
        self.db_password_entry.grid(row=0, column=3, padx=(0, 8), pady=3, sticky="ew")
        self.db_login_button = ttk.Button(
            login_frame, text="登录", command=self.login_database
        )
        self.db_login_button.grid(row=0, column=4, padx=(0, 8), pady=3)
        self.db_status_label = ttk.Label(login_frame, text="未登录")
        self.db_status_label.grid(row=1, column=0, columnspan=5, sticky="w", pady=(3, 0))
        login_frame.columnconfigure(1, weight=1)
        login_frame.columnconfigure(3, weight=1)
        if self.localMode:
            self.db_username_entry.configure(state=tk.DISABLED)
            self.db_password_entry.configure(state=tk.DISABLED)
            self.db_login_button.configure(state=tk.DISABLED)
            self.db_status_label.configure(text="本地 Excel 模式，无需登录")

        # 生成输入框（分两列布局，减少滚动）
        self.input_vars = {}
        field_count = len(FIELDS)
        col1_fields = FIELDS[:field_count//2]
        col2_fields = FIELDS[field_count//2:]
        
        # 第一列
        scrollable_frame.columnconfigure(1, weight=1)
        scrollable_frame.columnconfigure(3, weight=1)
        for idx, (field, typ) in enumerate(col1_fields):
            required = typ == "required" or (field == "样品编号" and not self.localMode)
            lbl = ttk.Label(scrollable_frame, text=f"{field}{'*' if required else ''}：")
            lbl.grid(row=idx, column=0, padx=10, pady=5, sticky="e")
            var = tk.StringVar()
            entry = ttk.Entry(scrollable_frame, textvariable=var, width=25)
            entry.grid(row=idx, column=1, padx=10, pady=5, sticky="ew")
            self.input_vars[field] = var
        
        # 第二列
        for idx, (field, typ) in enumerate(col2_fields):
            required = typ == "required" or (field == "样品编号" and not self.localMode)
            lbl = ttk.Label(scrollable_frame, text=f"{field}{'*' if required else ''}：")
            lbl.grid(row=idx, column=2, padx=10, pady=5, sticky="e")
            var = tk.StringVar()
            entry = ttk.Entry(scrollable_frame, textvariable=var, width=25)
            entry.grid(row=idx, column=3, padx=10, pady=5, sticky="ew")
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

    def login_database(self):
        """验证并保存当前数据库登录会话。"""
        if self.localMode:
            return
        if self._db_busy:
            return
        user = self.db_username_var.get().strip()
        password = self.db_password_var.get()
        if not user or not password.strip():
            self.db_status_label.configure(text="请输入用户名和密码")
            messagebox.showwarning("登录信息缺失", "请输入数据库用户名和密码。")
            return
        self.db_login_button.configure(state=tk.DISABLED)

        def authenticate():
            connection = connect_to_database(user, password)
            if not connection:
                return False
            connection.close()
            return True

        self._start_database_task(authenticate, self._finish_database_login)

    def _finish_database_login(self, authenticated, error):
        self.db_login_button.configure(state=tk.NORMAL)
        if error or not authenticated:
            self.db_status_label.configure(text="登录失败，请检查连接信息")
            messagebox.showerror(
                "数据库登录失败",
                str(error) if error else "无法连接数据库，请检查用户名、密码和网络。",
            )
            return

        self.db_user = self.db_username_var.get().strip()
        self.db_password = self.db_password_var.get()
        self.db_logged_in = True
        self.db_login_button.configure(text="退出登录", command=self.logout_database)
        self.db_username_entry.configure(state=tk.DISABLED)
        self.db_password_entry.configure(state=tk.DISABLED)
        self.db_status_label.configure(text=f"已登录：{self.db_user}")
        self.current_page = 0
        self.refresh_table()

    def _start_database_task(self, task, on_complete):
        if self._db_busy or self._closing:
            return False
        self._db_busy = True
        self.db_status_label.configure(text="数据库处理中…")
        future = self.db_executor.submit(task)
        self._poll_database_task(future, on_complete)
        return True

    def _poll_database_task(self, future, on_complete):
        if self._closing:
            return
        if not future.done():
            self._db_poll_id = self.root.after(
                50, lambda: self._poll_database_task(future, on_complete)
            )
            return

        self._db_poll_id = None
        self._db_busy = False
        try:
            result = future.result()
            error = None
        except Exception as exception:
            result = None
            error = exception
        pending_refresh = self._pending_refresh
        self._pending_refresh = False
        on_complete(result, error)

        if (pending_refresh or self._pending_refresh) and not self._db_busy and not self._closing:
            self._pending_refresh = False
            self.root.after_idle(self.refresh_table)

    def _start_export_task(self, selected_data, headers, template_path):
        if self._export_busy or self._closing:
            messagebox.showinfo("提示", "Word 导出任务正在进行，请稍候。")
            return
        self._export_busy = True
        self.export_button.configure(state=tk.DISABLED)
        future = self.export_executor.submit(
            export_selected_data, selected_data, headers, template_path
        )
        self._poll_export_task(future)

    def _poll_export_task(self, future):
        if self._closing:
            return
        if not future.done():
            self._export_poll_id = self.root.after(
                50, lambda: self._poll_export_task(future)
            )
            return

        self._export_poll_id = None
        self._export_busy = False
        self.export_button.configure(state=tk.NORMAL)
        try:
            success_count, fail_info = future.result()
            tip_msg = f"共导出{success_count}个文档！\n保存路径：{EXPORT_DIR}"
            if fail_info:
                tip_msg += f"\n失败项：\n{chr(10).join(fail_info)}"
            messagebox.showinfo("导出完成", tip_msg)
        except Exception as error:
            messagebox.showerror("导出失败", str(error))

    def close(self):
        self._closing = True
        for callback_id in (
            self._db_poll_id,
            self._export_poll_id,
            self._pane_constraint_after_id,
            self._search_after_id,
        ):
            if callback_id is not None:
                self.root.after_cancel(callback_id)
        self.db_executor.shutdown(wait=False, cancel_futures=True)
        self.export_executor.shutdown(wait=False, cancel_futures=True)
        self.root.destroy()

    def logout_database(self):
        """清除数据库会话并移除当前数据库数据。"""
        if self._db_busy:
            messagebox.showinfo("提示", "数据库正在处理中，请稍候再退出登录。")
            return
        self.db_user = None
        self.db_password = None
        self.db_logged_in = False
        self.db_password_var.set("")
        self.db_username_entry.configure(state=tk.NORMAL)
        self.db_password_entry.configure(state=tk.NORMAL)
        self.db_login_button.configure(text="登录", command=self.login_database)
        self.db_status_label.configure(text="未登录")
        self.refresh_table()

    def _schedule_pane_constraint(self, event=None):
        if self._pane_constraint_after_id is None:
            self._pane_constraint_after_id = self.root.after_idle(
                self._constrain_input_panel
            )

    def _constrain_input_panel(self):
        self._pane_constraint_after_id = None
        width = self.paned.winfo_width()
        if width <= 1:
            return
        minimum_left = min(600, max(420, width - 420))
        if self.paned.sashpos(0) < minimum_left:
            self.paned.sashpos(0, minimum_left)

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
        self.export_button = ttk.Button(
            btn_frame, text="导出选中到Word", command=self.export_selected, width=15
        )
        self.export_button.pack(side=tk.LEFT, padx=5)
        # 刷新按钮
        ttk.Button(btn_frame, text="刷新数据", command=self.refresh_table, width=12).pack(side=tk.LEFT, padx=5)
        # 搜索框（右侧）
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._schedule_filter)
        # 替换为自定义占位符输入框
        search_entry = PlaceholderEntry(btn_frame, textvariable=self.search_var, width=20, placeholder="搜索证书编号/样品名称")
        search_entry.pack(side=tk.RIGHT, padx=5)
        ttk.Label(btn_frame, text="搜索：").pack(side=tk.RIGHT, padx=5)

        # 表格（带复选框，优化列宽）
        self.tree = ttk.Treeview(self.data_frame, columns=[f"col{i}" for i in range(len(FIELDS))], show="headings", selectmode="extended")
        self.tree_font = tkfont.Font(root=self.root, font=self.style.lookup("Treeview", "font") or "TkDefaultFont")
        self._tree_row_data = {}
        self._column_widths_before_drag = None
        # 设置表头和列宽
        for idx, (field, _) in enumerate(FIELDS):
            self.tree.heading(f"col{idx}", text=field)
            # 关键字段加宽，其他字段适配
            if field in ["证书编号", "样品名称", "生产厂商"]:
                self.tree.column(f"col{idx}", width=210, minwidth=140, stretch=False)
            else:
                self.tree.column(f"col{idx}", width=120, minwidth=80, stretch=False)
        if not self.localMode:
            self.tree.configure(
                displaycolumns=[
                    f"col{index}"
                    for index, (field, _) in enumerate(FIELDS)
                    if field in DATABASE_LIST_FIELDS[:3]
                ]
            )
        # 滚动条
        scroll_y = ttk.Scrollbar(self.data_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scroll_x = ttk.Scrollbar(self.data_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        # 布局
        self.tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=5)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        page_frame = ttk.Frame(self.data_frame)
        page_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=(0, 5))
        self.previous_page_button = ttk.Button(
            page_frame, text="上一页", command=self.previous_page
        )
        self.previous_page_button.pack(side=tk.LEFT)
        self.page_status = ttk.Label(page_frame, text="第 1 页")
        self.page_status.pack(side=tk.LEFT, padx=10)
        self.next_page_button = ttk.Button(
            page_frame, text="下一页", command=self.next_page
        )
        self.next_page_button.pack(side=tk.LEFT)

        # 绑定事件
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<ButtonPress-1>", self._remember_column_widths, add="+")
        self.tree.bind("<ButtonRelease-1>", self._on_tree_button_release, add="+")
        # # 初始加载数据
        # refresh_table()

    def _wrap_tree_cell(self, value, width):
        """按像素宽度折行，支持中英文及单元格内的显式换行。"""
        text = "" if value is None else str(value)
        available_width = max(width - 12, 1)
        wrapped_lines = []
        for paragraph in text.split("\n"):
            line = ""
            for character in paragraph:
                if line and self.tree_font.measure(line + character) > available_width:
                    wrapped_lines.append(line)
                    line = character
                else:
                    line += character
            wrapped_lines.append(line)
        return "\n".join(wrapped_lines)

    def _insert_table_row(self, row, excel_row=None):
        """插入折行后的显示值，同时保留原始值供编辑使用。"""
        values = [
            self._wrap_tree_cell(value, self.tree.column(f"col{index}", "width"))
            for index, value in enumerate(row)
        ]
        item = self.tree.insert("", tk.END, values=values)
        self._tree_row_data[item] = list(row)
        if excel_row is not None:
            self._tree_row_positions[item] = excel_row

    def _refresh_tree_wrapping(self):
        """根据当前列宽更新单元格折行和统一行高。"""
        max_lines = 1
        for item, row in self._tree_row_data.items():
            values = [
                self._wrap_tree_cell(value, self.tree.column(f"col{index}", "width"))
                for index, value in enumerate(row)
            ]
            max_lines = max(max_lines, *(value.count("\n") + 1 for value in values))
            self.tree.item(item, values=values)
        row_height = max(25, max_lines * self.tree_font.metrics("linespace") + 6)
        self.style.configure("Treeview", rowheight=row_height)

    def _remember_column_widths(self, event):
        if self.tree.identify_region(event.x, event.y) == "separator":
            self._column_widths_before_drag = tuple(
                self.tree.column(f"col{index}", "width") for index in range(len(FIELDS))
            )
        else:
            self._column_widths_before_drag = None

    def _on_tree_button_release(self, event):
        if self._column_widths_before_drag is None:
            return
        current_widths = tuple(
            self.tree.column(f"col{index}", "width") for index in range(len(FIELDS))
        )
        if current_widths != self._column_widths_before_drag:
            self.root.after_idle(self._refresh_tree_wrapping)
        self._column_widths_before_drag = None

    def _schedule_filter(self, *args):
        if self._search_after_id is not None:
            self.root.after_cancel(self._search_after_id)
        self._search_after_id = self.root.after(
            300, lambda: self.filter_table(self.search_var.get())
        )

    def filter_table(self, keyword):
        """筛选数据并从第一页开始显示。"""
        self._search_after_id = None
        self.search_keyword = keyword.strip()
        self.current_page = 0
        self.refresh_table()

    def previous_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.refresh_table()

    def next_page(self):
        if self.has_next_page:
            self.current_page += 1
            self.refresh_table()

    def _update_pagination_controls(self, total_pages=None):
        self.previous_page_button.configure(
            state=tk.NORMAL if self.current_page > 0 else tk.DISABLED
        )
        self.next_page_button.configure(
            state=tk.NORMAL if self.has_next_page else tk.DISABLED
        )
        if total_pages is None:
            status = f"第 {self.current_page + 1} 页"
        else:
            status = f"第 {self.current_page + 1} / {total_pages} 页"
        self.page_status.configure(text=status)

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
        if not self.localMode and not self.db_logged_in:
            messagebox.showwarning("提示", "请先登录数据库！")
            return
        if not self.localMode and self._db_busy:
            messagebox.showinfo("提示", "数据库正在处理中，请稍候再提交。")
            return
        input_data = self.get_input_data()
        if not validate_input(input_data, require_sample_no=not self.localMode):
            return
        if self.localMode:
            saved = save_data_to_excel(input_data, self.current_edit_row)
        else:
            user = self.db_user
            password = self.db_password
            original_primary_key = self.current_edit_primary_key
            self._start_database_task(
                lambda: save_data_to_database(
                    input_data,
                    original_primary_key,
                    user=user,
                    password=password,
                ),
                self._finish_database_write,
            )
            return
        if saved:
            action = "修改" if self.current_edit_row is not None else "新增"
            messagebox.showinfo("成功", f"数据{action}成功！")
            self.current_edit_row = None
            self.current_edit_cert_no = None
            self.current_edit_primary_key = None
            self.current_page = 0
            self.clear_form()
            self.refresh_table()

    def _finish_database_write(self, result, error):
        if error:
            messagebox.showerror("数据库错误", str(error))
            self.db_status_label.configure(text="数据库操作失败")
            return
        saved, error_message = result
        if not saved:
            messagebox.showwarning("数据库操作未完成", error_message)
            self.db_status_label.configure(text=f"已登录：{self.db_user}")
            return

        action = "修改" if self.current_edit_primary_key is not None else "新增"
        messagebox.showinfo("成功", f"数据{action}成功！")
        self.current_edit_row = None
        self.current_edit_cert_no = None
        self.current_edit_primary_key = None
        self.current_page = 0
        self.clear_form()
        self.refresh_table()

    def clear_form(self):
        """清空表单"""
        for var in self.input_vars.values():
            var.set("")

    def cancel_edit(self):
        """取消编辑"""
        self.current_edit_row = None
        self.current_edit_cert_no = None
        self.current_edit_primary_key = None
        self.clear_form()

    def refresh_table(self):
        """刷新数据表格"""
        if not self.localMode and self.db_logged_in:
            if self._db_busy:
                self._pending_refresh = True
                return

            limit = self.page_size
            offset = self.current_page * self.page_size
            keyword = self.search_keyword
            user = self.db_user
            password = self.db_password

            def load_page():
                return UpdateTablefromDatabase(
                    limit=limit,
                    offset=offset,
                    keyword=keyword,
                    user=user,
                    password=password,
                )

            self._start_database_task(load_page, self._finish_database_refresh)
            return

        # 清空现有数据
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._tree_row_data.clear()
        self._tree_row_positions = {}
        # 加载新数据
        if self.localMode:
            self.all_table_data = read_all_data()
            if not self.all_table_data:
                self.all_table_data = [[field[0] for field in FIELDS]]
            rows = [
                (excel_row, row)
                for excel_row, row in enumerate(self.all_table_data[1:], start=2)
                if not self.search_keyword
                or self.search_keyword.lower() in row[0].lower()
                or self.search_keyword.lower() in row[2].lower()
            ]
            total_pages = max(1, (len(rows) + self.page_size - 1) // self.page_size)
            self.current_page = min(self.current_page, total_pages - 1)
            start = self.current_page * self.page_size
            page_rows = rows[start:start + self.page_size]
            self.has_next_page = self.current_page + 1 < total_pages
            for excel_row, row in page_rows:
                self._insert_table_row(row, excel_row)
            self._update_pagination_controls(total_pages)
        else:
            self.all_table_data = [[field[0] for field in FIELDS]]
            self.has_next_page = False
            self._update_pagination_controls()
        self._refresh_tree_wrapping()

    def _finish_database_refresh(self, result, error):
        if error:
            self.db_status_label.configure(text="数据库查询失败")
            messagebox.showerror("数据库错误", f"查询失败：{error}")
            return

        self.all_table_data, self.has_next_page = result
        if len(self.all_table_data) == 1 and self.current_page > 0:
            self.current_page -= 1
            self._pending_refresh = True
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._tree_row_data.clear()
        self._tree_row_positions = {}
        for row in self.all_table_data[1:]:
            self._insert_table_row(row)
        self._update_pagination_controls()
        self._refresh_tree_wrapping()
        self.db_status_label.configure(text=f"已登录：{self.db_user}")

    def on_select(self, event):
        """获取选中的行"""
        self.selected_items = self.tree.selection()

    def edit_selected(self):
        """编辑选中的行"""
        if len(self.selected_items) != 1:
            messagebox.showwarning("提示", "请仅选中一行数据进行修改！")
            return
        # 获取选中行数据
        selected_item = self.selected_items[0]
        row_data = self._tree_row_data.get(selected_item, self.tree.item(selected_item)["values"])
        if self.localMode:
            data_dict = dict(zip([field[0] for field in FIELDS], row_data))
            self.fill_form(data_dict)
            self.current_edit_row = self._tree_row_positions[selected_item]
            self.current_edit_cert_no = None
            self.current_edit_primary_key = None
            return

        if self._db_busy:
            messagebox.showinfo("提示", "数据库正在处理中，请稍候再修改。")
            return
        primary_key = (row_data[0], row_data[5])
        user = self.db_user
        password = self.db_password
        self._start_database_task(
            lambda: get_data_from_database(
                primary_key, user=user, password=password
            ),
            lambda result, error: self._finish_database_record_load(
                selected_item, primary_key, result, error
            ),
        )

    def _finish_database_record_load(self, item, primary_key, result, error):
        if error:
            messagebox.showerror("数据库错误", str(error))
            return
        row, error_message = result
        if row is None:
            messagebox.showwarning("查询失败", error_message)
            return
        self._tree_row_data[item] = row
        data_dict = dict(zip([field[0] for field in FIELDS], row))
        self.fill_form(data_dict)
        self.current_edit_row = self.tree.index(item) + 2
        self.current_edit_cert_no = data_dict["证书编号"]
        self.current_edit_primary_key = primary_key

    def delete_selected(self):
        """删除选中的行"""
        if not self.selected_items:
            messagebox.showwarning("提示", "请选中要删除的数据！")
            return
        if self.localMode:
            row_numbers = sorted(
                (self._tree_row_positions[item] for item in self.selected_items),
                reverse=True,
            )
            for row_num in row_numbers:
                delete_data(row_num)
        elif messagebox.askyesno("确认", "是否删除选中的数据？"):
            primary_keys = [
                (self._tree_row_data[item][0], self._tree_row_data[item][5])
                for item in self.selected_items
            ]
            user = self.db_user
            password = self.db_password

            def delete_records():
                results = [
                    delete_data_from_database(
                        cert_no, sample_no, user=user, password=password
                    )
                    for cert_no, sample_no in primary_keys
                ]
                return results

            self._start_database_task(delete_records, self._finish_database_delete)
            return
        # 刷新表格
        self.refresh_table()

    def _finish_database_delete(self, results, error):
        if error:
            messagebox.showerror("数据库错误", str(error))
            self.db_status_label.configure(text="数据库删除失败")
            return
        failures = [message for success, message in results if not success]
        if failures:
            messagebox.showwarning("删除结果", "\n".join(failures))
        else:
            messagebox.showinfo("成功", "选中数据已删除！")
        self.current_page = 0
        self.refresh_table()

    def export_selected(self):
        """导出选中的数据"""
        if not self.selected_items:
            messagebox.showwarning("提示", "请选中要导出的数据！")
            return
        template_path = filedialog.askopenfilename(
            title="选择Word模板",
            filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")],
            initialfile=TEMPLATE_DEFAULT_PATH,
        )
        if not template_path:
            return
        if not os.path.exists(template_path):
            messagebox.showerror("错误", "模板文件不存在！")
            return

        headers = [field[0] for field in FIELDS]
        if self.localMode:
            selected_data = [self._tree_row_data[item] for item in self.selected_items]
            self._start_export_task(selected_data, headers, template_path)
            return

        if self._db_busy:
            messagebox.showinfo("提示", "数据库正在处理中，请稍候再导出。")
            return
        primary_keys = [
            (self._tree_row_data[item][0], self._tree_row_data[item][5])
            for item in self.selected_items
        ]
        user = self.db_user
        password = self.db_password

        def load_export_rows():
            records = []
            errors = []
            for primary_key in primary_keys:
                row, error = get_data_from_database(
                    primary_key, user=user, password=password
                )
                if row is None:
                    errors.append(error)
                else:
                    records.append(row)
            return records, errors

        self._start_database_task(
            load_export_rows,
            lambda result, error: self._finish_database_export_load(
                result, error, headers, template_path
            ),
        )

    def _finish_database_export_load(self, result, error, headers, template_path):
        if error:
            messagebox.showerror("数据库错误", f"读取导出数据失败：{error}")
            return
        records, errors = result
        if errors:
            messagebox.showwarning("部分记录未读取", "\n".join(errors))
        if records:
            self._start_export_task(records, headers, template_path)

# -------------------------- 程序入口 --------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = DeviceCheckApp(root)
    root.mainloop()
