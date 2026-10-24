import shutil  # 用于复制文件
import tempfile  # 用于创建临时文件（可选）

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, messagebox, filedialog
from calendar import monthrange
import openpyxl
import os
import re
import statistics
import threading
import time
from datetime import date, datetime, timedelta
from concurrent.futures import ThreadPoolExecutor
from docx import Document
from docx.shared import Inches

import mysql.connector
from mysql.connector import Error
from mysql.connector.pooling import MySQLConnectionPool
# -------------------------- 全局配置 --------------------------
# 字段定义（按需求排序，标注必填项）

FIELDS = [
    ("样品状态", "optional"), ("证书编号", "required"), ("送检班组", "required"), ("样品名称", "required"),
    ("生产厂商", "optional"), ("型号", "optional"), ("样品编号", "required"),
    ("批准人", "optional"), ("核验员", "optional"), ("检定员", "optional"),
    ("投运日期", "optional"), ("校验日期", "required"),
    ("有效日期", "optional"), ("安装位置", "optional"),
    ("设备精度", "optional"), ("量程范围", "optional"), ("输出范围", "optional"),
    ("外观检查", "optional"), ("绝缘电阻", "optional"), ("输出纹波含量", "optional"),
    ("输出值1", "optional"), ("输出值2", "optional"), ("输出值3", "optional"),
    ("输出值4", "optional"), ("输出值5", "optional"), ("输出值6", "optional"),
    ("输出值7", "optional"), ("输出值8", "optional"), ("输出值9", "optional"),
    ("输出值10", "optional"), ("输出值11", "optional"), ("误差1", "optional"),
    ("误差2", "optional"), ("误差3", "optional"), ("误差4", "optional"),
    ("误差5", "optional"), ("误差6", "optional"), ("误差7", "optional"),
    ("误差8", "optional"), ("误差9", "optional"), ("误差10", "optional"),
    ("误差11", "optional"), ("最大基本误差", "optional"),
    ("生命周期", "optional"), ("检定周期", "optional")
]
FIELD_INDEX = {field: index for index, (field, _) in enumerate(FIELDS)}

PAGE_SIZE = 15
CERTIFICATE_TYPES = ("检定", "校准")
READ_ONLY_FIELDS = {
    "样品状态", "证书编号", "有效日期", "最大基本误差",
    *(f"误差{index}" for index in range(1, 12)),
}
DATABASE_LIST_FIELDS = (
    "样品状态", "证书编号", "送检班组", "样品名称", "样品编号", "投运日期",
    "校验日期", "有效日期", "生命周期", "检定周期"
)
DATABASE_LIST_CERTIFICATE_RESULT_INDEX = 1 + DATABASE_LIST_FIELDS.index("证书编号")
DATABASE_DISPLAY_FIELDS = (
    "证书编号", "送检班组", "样品名称", "投运日期", "校验日期",
    "有效日期", "生命周期", "检定周期"
)
UNQUALIFIED_ROW_TAG = "sample-unqualified"
DATE_HIGHLIGHT_COLORS = {
    "投运日期": "#F43E06",
    "校验日期": "#EF6F48",
}
REFERENCE_PERCENTAGES = (0, 20, 40, 60, 80, 100, 80, 60, 40, 20, 0)
RANGE_PATTERN = re.compile(
    r"^\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+))\s*-\s*"
    r"([+-]?(?:\d+(?:\.\d*)?|\.\d+))\s*[^\d]*$"
)
MEASUREMENT_PATTERN = re.compile(r"^\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+))\s*[^\d]*$")
PRECISION_PATTERN = re.compile(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)")
_database_pool = None
_database_pool_config = None
_database_pool_lock = threading.Lock()


def _parse_database_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        return date.fromisoformat(text[:10])
    except ValueError:
        for date_format in ("%Y年%m月%d日", "%Y/%m/%d", "%Y.%m.%d"):
            try:
                return datetime.strptime(text, date_format).date()
            except ValueError:
                continue
    return None


def generate_certificate_number(certificate_type, sample_number, calibration_date):
    parsed_date = _parse_database_date(calibration_date)
    if certificate_type not in CERTIFICATE_TYPES or not sample_number or parsed_date is None:
        return ""
    return f"{certificate_type}{sample_number}-{parsed_date.year}"


def generate_reference_values(range_text):
    if range_text is None:
        return []
    match = RANGE_PATTERN.fullmatch(str(range_text))
    if match is None:
        return []
    start_value, end_value = (float(value) for value in match.groups())
    span = end_value - start_value
    return [
        round(start_value + span * percentage / 100, 2)
        for percentage in REFERENCE_PERCENTAGES
    ]


def calculate_maximum_allowed_error(precision_text):
    precision_match = PRECISION_PATTERN.search(str(precision_text or ""))
    if precision_match is None:
        return ""
    return f"{precision_match.group()}%"


def calculate_sample_status(maximum_basic_error, maximum_allowed_error):
    basic_error = _parse_measurement(maximum_basic_error)
    allowed_error = _parse_measurement(maximum_allowed_error)
    if basic_error is None or allowed_error is None:
        return ""
    return "不合格" if basic_error > allowed_error else "合格"


def _parse_measurement(value):
    if value is None:
        return None
    match = MEASUREMENT_PATTERN.fullmatch(str(value))
    return float(match.group(1)) if match else None


def calculate_reference_data(input_range, output_range, measured_values):
    reference_inputs = generate_reference_values(input_range)
    reference_outputs = generate_reference_values(output_range)
    output_range_match = RANGE_PATTERN.fullmatch(str(output_range or ""))
    output_span = None
    if output_range_match:
        range_start, range_end = (float(value) for value in output_range_match.groups())
        if range_end != range_start:
            output_span = range_end - range_start

    errors = []
    numeric_errors = []
    for index in range(11):
        measured = _parse_measurement(measured_values[index])
        if (
            measured is None
            or output_span is None
            or len(reference_outputs) != 11
        ):
            errors.append("")
            continue
        error_value = (measured - reference_outputs[index]) / output_span * 100
        errors.append(f"{error_value:.4f}")
        numeric_errors.append(error_value)

    max_error = f"{max(abs(error) for error in numeric_errors):.4f}" if numeric_errors else ""
    return {
        "reference_inputs": reference_inputs,
        "reference_outputs": reference_outputs,
        "errors": errors,
        "max_error": max_error,
    }


def build_word_data_dict(row_data, headers):
    data_dict = dict(zip(headers, row_data))
    derived = calculate_reference_data(
        data_dict.get("量程范围"),
        data_dict.get("输出范围"),
        [data_dict.get(f"输出值{index}") for index in range(1, 12)],
    )
    for index, value in enumerate(derived["reference_inputs"], start=1):
        data_dict[f"输入值{index}"] = f"{value:g}"
    for index, value in enumerate(derived["reference_outputs"], start=1):
        data_dict[f"标准值{index}"] = f"{value:g}"
    for index, value in enumerate(derived["errors"], start=1):
        data_dict[f"误差{index}"] = value
    data_dict["最大基本误差"] = derived["max_error"]
    data_dict["最大允许误差"] = calculate_maximum_allowed_error(
        data_dict.get("设备精度")
    )
    return data_dict


def _period_in_months(value):
    if value is None or not str(value).strip():
        return None
    match = re.fullmatch(r"(\d+(?:\.\d+)?)\s*(年|个月|月)?", str(value).strip())
    if not match:
        return None
    amount = float(match.group(1))
    months = round(amount if match.group(2) in {"个月", "月"} else amount * 12)
    return months if months > 0 else None


def _date_after_period(start_date, period):
    months_to_add = _period_in_months(period)
    if start_date is None or months_to_add is None:
        return None
    month_index = start_date.year * 12 + start_date.month - 1 + months_to_add
    year, month_index = divmod(month_index, 12)
    month = month_index + 1
    day = min(start_date.day, monthrange(year, month)[1])
    return date(year, month, day)


def database_row_date_highlights(row, today=None):
    today = today or date.today()
    indexes = {field: index for index, (field, _) in enumerate(FIELDS)}
    highlights = {}
    service_life_end = _date_after_period(
        _parse_database_date(row[indexes["投运日期"]]),
        row[indexes["生命周期"]],
    )
    if service_life_end is not None and service_life_end < today:
        highlights["投运日期"] = DATE_HIGHLIGHT_COLORS["投运日期"]

    verification_end = _date_after_period(
        _parse_database_date(row[indexes["校验日期"]]),
        row[indexes["检定周期"]],
    )
    if verification_end is not None and verification_end < today:
        highlights["校验日期"] = DATE_HIGHLIGHT_COLORS["校验日期"]
    return highlights

            
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
    """Delete one Excel row after the caller has confirmed the operation."""
    wb = safe_open_excel()
    if not wb:
        return False
    try:
        ws = wb.active
        ws.delete_rows(row_num)
        wb.save(EXCEL_PATH)
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
        config = {
            "host": os.getenv("MYSQL_HOST", "139.155.96.177"),
            "port": int(os.getenv("MYSQL_PORT", "3306")),
            "user": user,
            "password": password,
            "database": os.getenv("MYSQL_DATABASE", "CMDSystem"),
        }
        global _database_pool, _database_pool_config
        with _database_pool_lock:
            if _database_pool is None:
                _database_pool = MySQLConnectionPool(
                    pool_name="calibration_data_pool",
                    pool_size=1,
                    pool_reset_session=False,
                    **config,
                )
                _database_pool_config = config.copy()
            elif _database_pool_config != config:
                _database_pool.set_config(**config)
                _database_pool_config = config.copy()
            connection = _database_pool.get_connection()
        return connection
    except (Error, ValueError) as error:
        print(f"MySQL connection failed: {error}")
    return None


def close_database_pool():
    """Close idle pooled connections, for example when the user logs out."""
    global _database_pool, _database_pool_config
    with _database_pool_lock:
        pool = _database_pool
        _database_pool = None
        _database_pool_config = None
    if pool is not None:
        pool._remove_connections()


def release_database_connection(connection):
    """Return a pooled connection without a full session reset round-trip."""
    try:
        if connection.in_transaction:
            connection.rollback()
    finally:
        connection.close()


def build_database_list_query(keyword, limit, offset):
    columns = ", ".join(f"`{field}`" for field in DATABASE_LIST_FIELDS)
    conditions = ["`证书编号` IS NOT NULL"]
    filter_parameters = []
    if keyword:
        conditions.append(
            "(`证书编号` LIKE %s OR `样品名称` LIKE %s OR `样品状态` = %s)"
        )
        pattern = f"%{keyword}%"
        filter_parameters.extend([pattern, pattern, keyword])
    where_clause = f"WHERE {' AND '.join(conditions)}"
    page_columns = ", ".join(f"page.`{field}`" for field in DATABASE_LIST_FIELDS)
    query = (
        "SELECT totals.total_count, " + page_columns + " "
        "FROM ("
        f"SELECT COUNT(*) AS total_count FROM `设备校验数据` {where_clause}"
        ") AS totals "
        "LEFT JOIN ("
        f"SELECT {columns} FROM `设备校验数据` {where_clause} "
        "ORDER BY `证书编号`, `样品编号` LIMIT %s OFFSET %s"
        ") AS page ON 1 = 1 "
        "ORDER BY page.`证书编号`, page.`样品编号`"
    )
    return query, filter_parameters + filter_parameters + [limit + 1, offset]


def measure_database_list_refresh(user, password, keyword="", limit=PAGE_SIZE, offset=0):
    started = time.perf_counter()
    connection = connect_to_database(user, password)
    if not connection:
        raise RuntimeError("无法从连接池获取数据库连接，请检查当前登录状态和连接配置。")

    cursor = None
    try:
        cursor = connection.cursor(buffered=True)
        query, parameters = build_database_list_query(keyword, limit, offset)
        cursor.execute(query, parameters)
        rows = cursor.fetchall()
        total_count = rows[0][0] if rows else 0
        page_rows = sum(
            1
            for row in rows
            if row[DATABASE_LIST_CERTIFICATE_RESULT_INDEX] is not None
        )
    finally:
        if cursor:
            cursor.close()
        release_database_connection(connection)

    return {
        "elapsed_ms": (time.perf_counter() - started) * 1000,
        "total_count": total_count,
        "page_rows": page_rows,
    }


def benchmark_database_performance(user, password, keyword="", limit=PAGE_SIZE, offset=0):
    """Measure connection setup and the read queries used by the database list."""
    config = {
        "host": os.getenv("MYSQL_HOST", "139.155.96.177"),
        "port": int(os.getenv("MYSQL_PORT", "3306")),
        "user": user,
        "password": password,
        "database": os.getenv("MYSQL_DATABASE", "CMDSystem"),
    }
    fresh_connection_times = []
    for _ in range(3):
        started = time.perf_counter()
        connection = mysql.connector.connect(**config)
        fresh_connection_times.append((time.perf_counter() - started) * 1000)
        connection.close()

    pool_checkout_times = []
    for _ in range(3):
        started = time.perf_counter()
        connection = connect_to_database(user, password)
        pool_checkout_times.append((time.perf_counter() - started) * 1000)
        if not connection:
            raise RuntimeError("无法从连接池获取数据库连接，请检查当前登录状态和连接配置。")
        release_database_connection(connection)

    pool_cycle_times = []
    for _ in range(3):
        started = time.perf_counter()
        connection = connect_to_database(user, password)
        if not connection:
            raise RuntimeError("无法从连接池获取数据库连接，请检查当前登录状态和连接配置。")
        release_database_connection(connection)
        pool_cycle_times.append((time.perf_counter() - started) * 1000)

    connection = connect_to_database(user, password)
    if not connection:
        raise RuntimeError("无法从连接池获取数据库连接，请检查当前登录状态和连接配置。")

    cursor = None
    try:
        cursor = connection.cursor(buffered=True)
        select_one_times = []
        for _ in range(3):
            started = time.perf_counter()
            cursor.execute("SELECT 1")
            cursor.fetchone()
            select_one_times.append((time.perf_counter() - started) * 1000)

        conditions = ["`证书编号` IS NOT NULL"]
        parameters = []
        if keyword:
            conditions.append(
                "(`证书编号` LIKE %s OR `样品名称` LIKE %s OR `样品状态` = %s)"
            )
            pattern = f"%{keyword}%"
            parameters.extend([pattern, pattern, keyword])
        where_clause = f"WHERE {' AND '.join(conditions)}"

        started = time.perf_counter()
        cursor.execute(
            f"SELECT COUNT(*) FROM `设备校验数据` {where_clause}",
            parameters,
        )
        total_count = cursor.fetchone()[0]
        count_ms = (time.perf_counter() - started) * 1000

        columns = ", ".join(f"`{field}`" for field in DATABASE_LIST_FIELDS)
        started = time.perf_counter()
        cursor.execute(
            f"SELECT {columns} FROM `设备校验数据` {where_clause} "
            "ORDER BY `证书编号`, `样品编号` LIMIT %s OFFSET %s",
            parameters + [limit + 1, offset],
        )
        page_rows = cursor.fetchall()
        page_ms = (time.perf_counter() - started) * 1000

        combined_query, combined_parameters = build_database_list_query(
            keyword, limit, offset
        )
        started = time.perf_counter()
        cursor.execute(combined_query, combined_parameters)
        combined_rows = cursor.fetchall()
        combined_page_ms = (time.perf_counter() - started) * 1000
        metrics = {
            "fresh_connection_ms": statistics.median(fresh_connection_times),
            "pool_checkout_ms": statistics.median(pool_checkout_times),
            "pool_cycle_ms": statistics.median(pool_cycle_times),
            "select_one_ms": statistics.median(select_one_times),
            "count_ms": count_ms,
            "total_count": total_count,
            "page_ms": page_ms,
            "page_rows": min(len(page_rows), limit),
            "combined_page_ms": combined_page_ms,
            "combined_page_rows": sum(
                1
                for row in combined_rows
                if row[DATABASE_LIST_CERTIFICATE_RESULT_INDEX] is not None
            ),
        }
    finally:
        if cursor:
            cursor.close()
        release_database_connection(connection)

    refresh = measure_database_list_refresh(
        user, password, keyword=keyword, limit=limit, offset=offset
    )
    metrics["refresh_cycle_ms"] = refresh["elapsed_ms"]
    metrics["refresh_total_count"] = refresh["total_count"]
    metrics["refresh_page_rows"] = refresh["page_rows"]
    return metrics


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
        return data, False, 1

    cursor = None
    query, parameters = build_database_list_query(keyword, limit, offset)
    try:
        cursor = connection.cursor(buffered=True)
        cursor.execute(query, parameters)
        rows = cursor.fetchall()
        total_count = rows[0][0] if rows else 0
        total_pages = max(1, (total_count + limit - 1) // limit)
        page_rows = [
            row
            for row in rows
            if row[DATABASE_LIST_CERTIFICATE_RESULT_INDEX] is not None
        ]
        has_more = offset + limit < total_count
        field_indexes = {field: headers.index(field) for field in DATABASE_LIST_FIELDS}
        for row in page_rows[:limit]:
            full_row = [""] * len(headers)
            for field, value in zip(DATABASE_LIST_FIELDS, row[1:]):
                full_row[field_indexes[field]] = str(value) if value is not None else ""
            data.append(full_row)
    finally:
        if cursor:
            cursor.close()
        release_database_connection(connection)
    return data, has_more, total_pages


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
        release_database_connection(connection)


def get_data_from_database_many(primary_keys, user=None, password=None):
    """Read several complete records using one pooled connection and cursor."""
    if not primary_keys:
        return [], []
    connection = connect_to_database(user, password)
    if not connection:
        return [], ["无法连接数据库，请重新登录。"] * len(primary_keys)

    headers = [field[0] for field in FIELDS]
    columns = ", ".join(f"`{header}`" for header in headers)
    records = []
    errors = []
    cursor = None
    try:
        cursor = connection.cursor(buffered=True)
        for primary_key in primary_keys:
            try:
                cursor.execute(
                    f"SELECT {columns} FROM `设备校验数据` "
                    "WHERE `证书编号` = %s AND `样品编号` = %s LIMIT 1",
                    primary_key,
                )
                row = cursor.fetchone()
                if row is None:
                    errors.append(f"未找到记录：{primary_key[0]} / {primary_key[1]}")
                else:
                    records.append(
                        [str(value) if value is not None else "" for value in row]
                    )
            except Error as error:
                errors.append(f"读取记录失败：{error}")
    finally:
        if cursor:
            cursor.close()
        release_database_connection(connection)
    return records, errors


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
        release_database_connection(connection)


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
        release_database_connection(connection)


def delete_data_from_database_many(primary_keys, user=None, password=None):
    """Delete several records in one transaction using one connection."""
    if not primary_keys:
        return []
    connection = connect_to_database(user, password)
    if not connection:
        return [(False, "无法连接数据库，请重新登录。") for _ in primary_keys]

    results = []
    cursor = None
    try:
        cursor = connection.cursor()
        for cert_no, sample_no in primary_keys:
            cursor.execute(
                "DELETE FROM `设备校验数据` "
                "WHERE `证书编号` = %s AND `样品编号` = %s",
                (cert_no, sample_no),
            )
            if cursor.rowcount == 0:
                results.append((False, f"未找到记录：{cert_no} / {sample_no}"))
            else:
                results.append((True, ""))
        connection.commit()
        return results
    except Error as error:
        connection.rollback()
        return [(False, f"批量删除失败，事务已回滚：{error}") for _ in primary_keys]
    finally:
        if cursor:
            cursor.close()
        release_database_connection(connection)




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
        data_dict = build_word_data_dict(export_row, headers)
        
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
        cert_no = row_data[FIELD_INDEX["证书编号"]]
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
        self._last_operated_key = None
        self._last_operated_row_number = None
        self.localMode = os.getenv("LOCAL_MODE", "0").strip().lower() in {"1", "true", "yes"}
        self.db_user = None
        self.db_password = None
        self.db_logged_in = self.localMode
        self.db_executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="database")
        self.export_executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="word-export")
        self._db_busy = False
        self._db_future = None
        self._pending_refresh = False
        self._db_poll_id = None
        self._export_busy = False
        self._export_poll_id = None
        self._closing = False
        self._date_highlight_after_id = None
        self.page_size = PAGE_SIZE
        self.current_page = 0
        self.total_pages = 1
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
        self._schedule_date_highlight_refresh()

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
        self.db_benchmark_button = ttk.Button(
            login_frame,
            text="测试耗时",
            command=self.benchmark_database,
            state=tk.DISABLED,
        )
        self.db_benchmark_button.grid(row=0, column=5, padx=(0, 8), pady=3)
        self.db_status_label = ttk.Label(login_frame, text="未登录")
        self.db_status_label.grid(row=1, column=0, columnspan=6, sticky="w", pady=(3, 0))
        login_frame.columnconfigure(1, weight=1)
        login_frame.columnconfigure(3, weight=1)
        if self.localMode:
            self.db_username_entry.configure(state=tk.DISABLED)
            self.db_password_entry.configure(state=tk.DISABLED)
            self.db_login_button.configure(state=tk.DISABLED)
            self.db_benchmark_button.configure(state=tk.DISABLED)
            self.db_status_label.configure(text="本地 Excel 模式，无需登录")

        # 生成输入框（分两列布局，减少滚动）
        self.input_vars = {}
        self.sample_status_var = tk.StringVar(master=self.root, value="")
        self.input_vars["样品状态"] = self.sample_status_var
        self.reference_input_values = []
        self.reference_output_values = []
        self.maximum_allowed_error = ""
        self._loading_form = False
        self.certificate_type_var = tk.StringVar(
            master=self.root, value=CERTIFICATE_TYPES[0]
        )
        ttk.Label(scrollable_frame, text="证书类型：").grid(
            row=0, column=0, padx=10, pady=5, sticky="e"
        )
        ttk.Combobox(
            scrollable_frame,
            textvariable=self.certificate_type_var,
            values=CERTIFICATE_TYPES,
            state="readonly",
            width=22,
        ).grid(row=0, column=1, padx=10, pady=5, sticky="ew")
        ttk.Label(scrollable_frame, text="样品状态：").grid(
            row=0, column=2, padx=10, pady=5, sticky="e"
        )
        self.sample_status_entry = ttk.Entry(
            scrollable_frame,
            textvariable=self.sample_status_var,
            width=25,
            state="readonly",
        )
        self.sample_status_entry.grid(
            row=0, column=3, padx=10, pady=5, sticky="ew"
        )
        form_fields = [field for field in FIELDS if field[0] != "样品状态"]
        field_count = len(form_fields)
        col1_fields = form_fields[:field_count//2]
        col2_fields = form_fields[field_count//2:]
        
        # 第一列
        scrollable_frame.columnconfigure(1, weight=1)
        scrollable_frame.columnconfigure(3, weight=1)
        for idx, (field, typ) in enumerate(col1_fields):
            required = typ == "required" or (field == "样品编号" and not self.localMode)
            lbl = ttk.Label(scrollable_frame, text=f"{field}{'*' if required else ''}：")
            lbl.grid(row=idx + 1, column=0, padx=10, pady=5, sticky="e")
            var = tk.StringVar()
            entry = ttk.Entry(
                scrollable_frame,
                textvariable=var,
                width=25,
                state="readonly" if field in READ_ONLY_FIELDS else tk.NORMAL,
            )
            entry.grid(row=idx + 1, column=1, padx=10, pady=5, sticky="ew")
            self.input_vars[field] = var
        
        # 第二列
        for idx, (field, typ) in enumerate(col2_fields):
            required = typ == "required" or (field == "样品编号" and not self.localMode)
            lbl = ttk.Label(scrollable_frame, text=f"{field}{'*' if required else ''}：")
            lbl.grid(row=idx + 1, column=2, padx=10, pady=5, sticky="e")
            var = tk.StringVar()
            entry = ttk.Entry(
                scrollable_frame,
                textvariable=var,
                width=25,
                state="readonly" if field in READ_ONLY_FIELDS else tk.NORMAL,
            )
            entry.grid(row=idx + 1, column=3, padx=10, pady=5, sticky="ew")
            self.input_vars[field] = var

        self.certificate_type_var.trace_add(
            "write", self._update_derived_form_values
        )
        for field in ("样品编号", "校验日期", "检定周期"):
            self.input_vars[field].trace_add(
                "write", self._update_derived_form_values
            )
        for field in ("量程范围", "输出范围"):
            self.input_vars[field].trace_add("write", self._update_reference_values)
        self.input_vars["设备精度"].trace_add(
            "write", self._update_maximum_allowed_error
        )
        for index in range(1, 12):
            self.input_vars[f"输出值{index}"].trace_add(
                "write", self._update_basic_errors
            )

        # 操作按钮（居中）
        btn_frame = ttk.Frame(scrollable_frame)
        max_row = max(len(col1_fields), len(col2_fields))
        btn_frame.grid(row=max_row + 1, column=0, columnspan=4, pady=20)
        # 提交按钮
        ttk.Button(btn_frame, text="提交数据", command=self.submit_data, width=12).grid(row=0, column=0, padx=8)
        # 清空按钮
        ttk.Button(btn_frame, text="清空表单", command=self.clear_form, width=12).grid(row=0, column=1, padx=8)
        # 取消编辑按钮
        ttk.Button(btn_frame, text="取消编辑", command=self.cancel_edit, width=12).grid(row=0, column=2, padx=8)

    def _update_derived_form_values(self, *args):
        if self._loading_form:
            return
        self.input_vars["证书编号"].set(
            generate_certificate_number(
                self.certificate_type_var.get(),
                self.input_vars["样品编号"].get().strip(),
                self.input_vars["校验日期"].get().strip(),
            )
        )
        valid_until = _date_after_period(
            _parse_database_date(self.input_vars["校验日期"].get()),
            self.input_vars["检定周期"].get(),
        )
        self.input_vars["有效日期"].set(
            valid_until.isoformat() if valid_until is not None else ""
        )

    def _update_reference_values(self, *args):
        if self._loading_form:
            return
        derived = self._calculate_reference_data()
        self.reference_input_values = derived["reference_inputs"]
        self.reference_output_values = derived["reference_outputs"]
        self._set_basic_error_values(derived)

    def _update_maximum_allowed_error(self, *args):
        if not self._loading_form:
            self.maximum_allowed_error = calculate_maximum_allowed_error(
                self.input_vars["设备精度"].get()
            )
            self._update_sample_status()

    def _calculate_reference_data(self):
        return calculate_reference_data(
            self.input_vars["量程范围"].get(),
            self.input_vars["输出范围"].get(),
            [self.input_vars[f"输出值{index}"].get() for index in range(1, 12)],
        )

    def _update_basic_errors(self, *args):
        if not self._loading_form:
            self._set_basic_error_values(self._calculate_reference_data())

    def _set_basic_error_values(self, derived):
        for index, value in enumerate(derived["errors"], start=1):
            self.input_vars[f"误差{index}"].set(value)
        self.input_vars["最大基本误差"].set(derived["max_error"])
        self._update_sample_status()

    def _update_sample_status(self):
        self.sample_status_var.set(
            calculate_sample_status(
                self.input_vars["最大基本误差"].get(),
                self.maximum_allowed_error,
            )
        )

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
            release_database_connection(connection)
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
        self.db_benchmark_button.configure(state=tk.NORMAL)
        self.db_username_entry.configure(state=tk.DISABLED)
        self.db_password_entry.configure(state=tk.DISABLED)
        self.db_status_label.configure(text=f"已登录：{self.db_user}")
        self.current_page = 0
        self.refresh_table()

    def benchmark_database(self):
        if not self.db_logged_in or self.localMode:
            return
        if self._db_busy:
            messagebox.showinfo("提示", "数据库正在处理中，请稍候再测试。")
            return

        self.db_benchmark_button.configure(state=tk.DISABLED)
        started = time.perf_counter()
        if not self._start_database_task(
            lambda: benchmark_database_performance(
                self.db_user,
                self.db_password,
                keyword=self.search_keyword,
                limit=self.page_size,
                offset=self.current_page * self.page_size,
            ),
            self._finish_database_benchmark,
        ):
            self.db_benchmark_button.configure(state=tk.NORMAL)
            return
        self._benchmark_started = started

    def _finish_database_benchmark(self, result, error):
        self.db_benchmark_button.configure(
            state=tk.NORMAL if self.db_logged_in and not self._closing else tk.DISABLED
        )
        if error:
            messagebox.showerror("耗时测试失败", str(error))
            return

        total_ms = (time.perf_counter() - self._benchmark_started) * 1000
        messagebox.showinfo(
            "数据库耗时测试",
            f"物理新建连接（3 次中位数）：{result['fresh_connection_ms']:.1f} ms\n"
            f"连接池借出（3 次中位数）：{result['pool_checkout_ms']:.1f} ms\n"
            f"连接池借出+归还（3 次中位数）：{result['pool_cycle_ms']:.1f} ms\n"
            f"同连接 SELECT 1（3 次中位数）：{result['select_one_ms']:.1f} ms\n"
            f"COUNT(*)：{result['count_ms']:.1f} ms（匹配 {result['total_count']} 条）\n"
            f"分页查询：{result['page_ms']:.1f} ms（取回 {result['page_rows']} 条）\n"
            f"合并计数+分页：{result['combined_page_ms']:.1f} ms "
            f"（取回 {result['combined_page_rows']} 条）\n"
            f"实际刷新周期（借出+合并查询+归还）：{result['refresh_cycle_ms']:.1f} ms "
            f"（匹配 {result['refresh_total_count']} 条，取回 {result['refresh_page_rows']} 条）\n"
            f"整套对照测试耗时：{total_ms:.1f} ms",
        )

    def _start_database_task(self, task, on_complete):
        if self._db_busy or self._closing:
            return False
        self._db_busy = True
        self.db_status_label.configure(text="数据库处理中…")
        future = self.db_executor.submit(task)
        self._db_future = future
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
        if self._db_future is future:
            self._db_future = None
        try:
            result = future.result()
            error = None
        except Exception as exception:
            result = None
            error = exception
        status = f"已登录：{self.db_user}" if self.db_logged_in else "未登录"
        self.db_status_label.configure(text=status)
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
            self._date_highlight_after_id,
        ):
            if callback_id is not None:
                self.root.after_cancel(callback_id)
        self.db_executor.shutdown(wait=False, cancel_futures=True)
        database_future = self._db_future
        if database_future is not None and not database_future.done():
            database_future.add_done_callback(lambda future: close_database_pool())
        else:
            close_database_pool()
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
        close_database_pool()
        self.db_benchmark_button.configure(state=tk.DISABLED)
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
        search_entry = PlaceholderEntry(
            btn_frame,
            textvariable=self.search_var,
            width=24,
            placeholder="搜索证书编号/样品名称/样品状态",
        )
        search_entry.pack(side=tk.RIGHT, padx=5)
        ttk.Label(btn_frame, text="搜索：").pack(side=tk.RIGHT, padx=5)

        # 表格（带复选框，优化列宽）
        table_frame = ttk.Frame(self.data_frame)
        table_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.operation_marker_canvas = tk.Canvas(
            table_frame,
            width=12,
            highlightthickness=0,
            borderwidth=0,
            background=self.root.cget("background"),
        )
        self.operation_marker_canvas.pack(side=tk.LEFT, fill=tk.Y)
        self.tree = ttk.Treeview(
            table_frame,
            columns=[f"col{i}" for i in range(len(FIELDS))],
            show="headings",
            selectmode="extended",
        )
        self.tree.tag_configure(
            UNQUALIFIED_ROW_TAG,
            background="#D92121",
            foreground="#FFFFFF",
        )
        self.tree_font = tkfont.Font(root=self.root, font=self.style.lookup("Treeview", "font") or "TkDefaultFont")
        self._tree_row_data = {}
        self._column_widths_before_drag = None
        self._date_cell_overlays = {}
        # 设置表头和列宽
        for idx, (field, _) in enumerate(FIELDS):
            self.tree.heading(f"col{idx}", text=field)
            # 关键字段加宽，其他字段适配
            if field in ["证书编号", "样品名称", "生产厂商"]:
                self.tree.column(
                    f"col{idx}", width=210, minwidth=140, stretch=False, anchor=tk.CENTER
                )
            else:
                self.tree.column(
                    f"col{idx}", width=120, minwidth=80, stretch=False, anchor=tk.CENTER
                )
        if not self.localMode:
            self.tree.configure(
                displaycolumns=[
                    f"col{index}"
                    for index, (field, _) in enumerate(FIELDS)
                    if field in DATABASE_DISPLAY_FIELDS
                ]
            )
        # 滚动条
        scroll_y = ttk.Scrollbar(self.data_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scroll_x = ttk.Scrollbar(self.data_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self._tree_scrollbar = scroll_y
        self._tree_xscrollbar = scroll_x
        self.tree.configure(
            yscrollcommand=self._on_tree_scroll,
            xscrollcommand=self._on_tree_xscroll,
        )
        # 布局
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
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
        ttk.Label(page_frame, text="跳转至").pack(side=tk.LEFT)
        self.page_number_var = tk.StringVar(value="1")
        self.page_number_entry = ttk.Entry(
            page_frame, textvariable=self.page_number_var, width=5, justify=tk.CENTER
        )
        self.page_number_entry.pack(side=tk.LEFT, padx=(4, 0))
        self.page_number_entry.bind("<Return>", lambda event: self.jump_to_page())
        ttk.Label(page_frame, text="页").pack(side=tk.LEFT, padx=(2, 4))
        ttk.Button(page_frame, text="跳转", command=self.jump_to_page).pack(side=tk.LEFT)
        self.next_page_button = ttk.Button(
            page_frame, text="下一页", command=self.next_page
        )
        self.next_page_button.pack(side=tk.LEFT)

        # 绑定事件
        self.tree.bind("<<TreeviewSelect>>", self.on_select)
        self.tree.bind("<ButtonPress-1>", self._remember_column_widths, add="+")
        self.tree.bind("<ButtonRelease-1>", self._on_tree_button_release, add="+")
        self.tree.bind(
            "<Configure>",
            lambda event: (self._draw_operation_marker(), self._position_date_highlights()),
            add="+",
        )
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
        """Insert values in full column order while retaining raw data separately."""
        displayed_indexes = self._displayed_tree_column_indexes()
        values = [""] * len(FIELDS)
        for index in displayed_indexes:
            values[index] = row[index]
        row_number = (
            excel_row
            if excel_row is not None
            else self.current_page * self.page_size + len(self.tree.get_children()) + 1
        )
        tags = (
            (UNQUALIFIED_ROW_TAG,)
            if not self.localMode and row[FIELD_INDEX["样品状态"]] == "不合格"
            else ()
        )
        item = self.tree.insert("", tk.END, text="", values=values, tags=tags)
        self._tree_row_data[item] = list(row)
        if excel_row is not None:
            self._tree_row_positions[item] = excel_row

    def _date_highlight_tags(self, row):
        if self.localMode:
            return {}
        return database_row_date_highlights(row)

    def _refresh_date_highlights(self):
        self._date_highlight_after_id = None
        self._position_date_highlights()
        self._schedule_date_highlight_refresh()

    def _schedule_date_highlight_refresh(self):
        if self._closing:
            return
        tomorrow = date.today() + timedelta(days=1)
        next_midnight = datetime.combine(tomorrow, datetime.min.time())
        delay_ms = max(1, int((next_midnight - datetime.now()).total_seconds() * 1000))
        self._date_highlight_after_id = self.root.after(
            delay_ms, self._refresh_date_highlights
        )

    @staticmethod
    def _row_identity(row):
        return (
            str(row[FIELD_INDEX["证书编号"]]),
            str(row[FIELD_INDEX["样品编号"]]),
        )

    def _row_is_last_operated(self, row, row_number):
        return (
            self._row_identity(row) == self._last_operated_key
            or row_number == self._last_operated_row_number
        )

    def _mark_last_operated_item(self, item):
        row = self._tree_row_data.get(item)
        if row is None:
            return
        self._last_operated_key = self._row_identity(row)
        self._last_operated_row_number = None
        self._draw_operation_marker()

    def _set_last_operated_key(self, key):
        self._last_operated_key = key
        self._last_operated_row_number = None
        self._draw_operation_marker()

    def _set_last_operated_row_number(self, row_number):
        self._last_operated_key = None
        self._last_operated_row_number = row_number if row_number and row_number > 0 else None
        self._draw_operation_marker()

    def _on_tree_scroll(self, first, last):
        self._tree_scrollbar.set(first, last)
        self._draw_operation_marker()
        self._position_date_highlights()

    def _on_tree_xscroll(self, first, last):
        self._tree_xscrollbar.set(first, last)
        self._position_date_highlights()

    def _position_date_highlights(self):
        if self.localMode or not hasattr(self, "tree"):
            return

        active_overlays = set()
        for item, row in self._tree_row_data.items():
            if row[FIELD_INDEX["样品状态"]] == "不合格":
                continue
            for field, color in self._date_highlight_tags(row).items():
                column_index = next(
                    index for index, (name, _) in enumerate(FIELDS) if name == field
                )
                column_id = f"col{column_index}"
                bounds = self.tree.bbox(item, column_id)
                key = (item, field)
                overlay = self._date_cell_overlays.get(key)
                if not bounds:
                    if overlay is not None:
                        overlay.place_forget()
                    continue

                x, y, width, height = bounds
                if overlay is None:
                    overlay = tk.Label(
                        self.tree,
                        anchor="center",
                        borderwidth=0,
                        padx=2,
                        pady=0,
                        font=self.tree_font,
                    )
                    overlay.bind(
                        "<ButtonPress-1>",
                        lambda event, row_item=item: self._forward_date_cell_click(
                            event, row_item
                        ),
                    )
                    self._date_cell_overlays[key] = overlay
                overlay.configure(text=str(row[column_index] or ""), background=color)
                overlay.place(x=x, y=y, width=width, height=height)
                active_overlays.add(key)

        visible_items = set(self.tree.get_children())
        for key, overlay in tuple(self._date_cell_overlays.items()):
            if key[0] not in visible_items:
                overlay.destroy()
                del self._date_cell_overlays[key]
            elif key not in active_overlays:
                overlay.place_forget()

    def _forward_date_cell_click(self, event, item):
        x = event.x_root - self.tree.winfo_rootx()
        y = event.y_root - self.tree.winfo_rooty()
        self.tree.event_generate(
            "<ButtonPress-1>", x=x, y=y, state=event.state
        )
        self.tree.event_generate(
            "<ButtonRelease-1>", x=x, y=y, state=event.state
        )
        return "break"

    def _draw_operation_marker(self):
        canvas = getattr(self, "operation_marker_canvas", None)
        if canvas is None:
            return
        canvas.delete("all")
        for item, row in self._tree_row_data.items():
            row_number = self._tree_row_positions.get(item)
            if row_number is None:
                row_number = self.current_page * self.page_size + self.tree.index(item) + 1
            if not self._row_is_last_operated(row, row_number):
                continue
            first_column = f"col{self._displayed_tree_column_indexes()[0]}"
            bounds = self.tree.bbox(item, first_column)
            if not bounds:
                continue
            _, row_top, _, row_height = bounds
            marker_height = min(16, row_height)
            top = row_top + (row_height - marker_height) // 2
            canvas.create_rectangle(
                3,
                top,
                9,
                top + marker_height,
                fill="#72F1AE",
                outline="",
            )

    def _displayed_tree_column_ids(self):
        indexes = self._displayed_tree_column_indexes()
        return [f"col{index}" for index in indexes]

    def _displayed_tree_column_indexes(self):
        if self.localMode:
            return list(range(len(FIELDS)))
        return [
            index
            for index, (field, _) in enumerate(FIELDS)
            if field in DATABASE_DISPLAY_FIELDS
        ]

    def _refresh_tree_wrapping(self):
        """根据当前列宽更新单元格折行和统一行高。"""
        max_lines = 1
        displayed_indexes = self._displayed_tree_column_indexes()

        for item, row in self._tree_row_data.items():
            values = [""] * len(FIELDS)
            for index in displayed_indexes:
                wrapped_value = self._wrap_tree_cell(
                    row[index], self.tree.column(f"col{index}", "width")
                )
                values[index] = wrapped_value
                max_lines = max(max_lines, wrapped_value.count("\n") + 1)
            self.tree.item(item, values=values)
        row_height = max(25, max_lines * self.tree_font.metrics("linespace") + 6)
        self.style.configure("Treeview", rowheight=row_height)
        self._draw_operation_marker()
        self._position_date_highlights()

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

    def jump_to_page(self):
        page_text = self.page_number_var.get().strip()
        if not page_text.isdecimal() or not 1 <= int(page_text) <= self.total_pages:
            messagebox.showwarning(
                "页码无效", f"请输入 1 到 {self.total_pages} 之间的页码。"
            )
            self.page_number_var.set(str(self.current_page + 1))
            return

        target_page = int(page_text) - 1
        if target_page != self.current_page:
            self.current_page = target_page
            self.refresh_table()

    def _update_pagination_controls(self, total_pages=None):
        if total_pages is not None:
            self.total_pages = max(1, total_pages)
        self.previous_page_button.configure(
            state=tk.NORMAL if self.current_page > 0 else tk.DISABLED
        )
        self.next_page_button.configure(
            state=tk.NORMAL if self.has_next_page else tk.DISABLED
        )
        self.page_status.configure(
            text=f"第 {self.current_page + 1} 页 / 共 {self.total_pages} 页"
        )
        self.page_number_var.set(str(self.current_page + 1))

    def get_input_data(self):
        """获取表单输入数据"""
        data = {}
        for field, _ in FIELDS:
            data[field] = self.input_vars[field].get().strip()
        return data

    def fill_form(self, data):
        """填充表单（编辑时）"""
        certificate_no = str(data.get("证书编号") or "")
        certificate_type = next(
            (
                value
                for value in CERTIFICATE_TYPES
                if certificate_no.startswith(value)
            ),
            self.certificate_type_var.get(),
        )
        self._loading_form = True
        try:
            self.certificate_type_var.set(certificate_type)
            for field, _ in FIELDS:
                self.input_vars[field].set(data.get(field) or "")
        finally:
            self._loading_form = False
        valid_until = _date_after_period(
            _parse_database_date(self.input_vars["校验日期"].get()),
            self.input_vars["检定周期"].get(),
        )
        self.input_vars["有效日期"].set(
            valid_until.isoformat() if valid_until is not None else ""
        )
        self.maximum_allowed_error = calculate_maximum_allowed_error(
            self.input_vars["设备精度"].get()
        )
        self._update_reference_values()

    def submit_data(self):
        """提交/修改数据"""
        if not self.localMode and not self.db_logged_in:
            messagebox.showwarning("提示", "请先登录数据库！")
            return
        if not self.localMode and self._db_busy:
            messagebox.showinfo("提示", "数据库正在处理中，请稍候再提交。")
            return
        for field in ("样品编号", "校验日期"):
            if not self.input_vars[field].get().strip():
                messagebox.showwarning("提示", f"{field}为必填项！")
                return
        input_data = self.get_input_data()
        is_editing = (
            self.current_edit_row is not None
            if self.localMode
            else self.current_edit_primary_key is not None
        )
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
                lambda result, error: self._finish_database_write(
                    result,
                    error,
                    (input_data["证书编号"], input_data["样品编号"]),
                    is_editing,
                ),
            )
            return
        if saved:
            action = "修改" if self.current_edit_row is not None else "新增"
            self._set_last_operated_key(
                (input_data["证书编号"], input_data["样品编号"])
            )
            messagebox.showinfo("成功", f"数据{action}成功！")
            self.current_edit_row = None
            self.current_edit_cert_no = None
            self.current_edit_primary_key = None
            if not is_editing:
                self.current_page = 0
            self.clear_form()
            self.refresh_table()

    def _finish_database_write(self, result, error, operated_key=None, is_editing=False):
        if error:
            messagebox.showerror("数据库错误", str(error))
            self.db_status_label.configure(text="数据库操作失败")
            return
        saved, error_message = result
        if not saved:
            messagebox.showwarning("数据库操作未完成", error_message)
            self.db_status_label.configure(text=f"已登录：{self.db_user}")
            return

        action = "修改" if is_editing else "新增"
        if operated_key is not None:
            self._set_last_operated_key(operated_key)
        messagebox.showinfo("成功", f"数据{action}成功！")
        self.current_edit_row = None
        self.current_edit_cert_no = None
        self.current_edit_primary_key = None
        if not is_editing:
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
            self.total_pages = total_pages
            start = self.current_page * self.page_size
            page_rows = rows[start:start + self.page_size]
            self.has_next_page = self.current_page + 1 < total_pages
            for excel_row, row in page_rows:
                self._insert_table_row(row, excel_row)
            self._update_pagination_controls(total_pages)
        else:
            self.all_table_data = [[field[0] for field in FIELDS]]
            self.total_pages = 1
            self.has_next_page = False
            self._update_pagination_controls(self.total_pages)
        self._refresh_tree_wrapping()

    def _finish_database_refresh(self, result, error):
        if error:
            self.db_status_label.configure(text="数据库查询失败")
            messagebox.showerror("数据库错误", f"查询失败：{error}")
            return

        self.all_table_data, self.has_next_page, total_pages = result
        self.total_pages = max(1, total_pages)
        if self.current_page >= self.total_pages:
            self.current_page = self.total_pages - 1
            self._pending_refresh = True
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._tree_row_data.clear()
        self._tree_row_positions = {}
        for row in self.all_table_data[1:]:
            self._insert_table_row(row)
        self._update_pagination_controls(self.total_pages)
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
        self._mark_last_operated_item(selected_item)
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
        primary_key = (
            row_data[FIELD_INDEX["证书编号"]],
            row_data[FIELD_INDEX["样品编号"]],
        )
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
        self._mark_last_operated_item(item)

    def delete_selected(self):
        """删除选中的行"""
        if not self.selected_items:
            messagebox.showwarning("提示", "请选中要删除的数据！")
            return
        if not messagebox.askyesno("确认", "是否删除选中的数据？", parent=self.root):
            return

        last_selected_item = self.selected_items[-1]
        self._mark_last_operated_item(last_selected_item)
        last_selected_data = self._tree_row_data[last_selected_item]
        last_selected_row_number = (
            self._tree_row_positions.get(last_selected_item)
            if self.localMode
            else self.current_page * self.page_size + self.tree.index(last_selected_item) + 1
        )
        if self.localMode:
            row_numbers = sorted(
                (self._tree_row_positions[item] for item in self.selected_items),
                reverse=True,
            )
            deleted_rows = []
            deleted_count = 0
            for row_num in row_numbers:
                if delete_data(row_num):
                    deleted_count += 1
                    deleted_rows.append(row_num)
            if deleted_rows:
                self._set_last_operated_row_number(min(deleted_rows) - 1)
                messagebox.showinfo(
                    "成功", f"已删除 {deleted_count} 条数据！", parent=self.root
                )
            self.refresh_table()
            return
        else:
            primary_keys = [
                (
                    self._tree_row_data[item][FIELD_INDEX["证书编号"]],
                    self._tree_row_data[item][FIELD_INDEX["样品编号"]],
                )
                for item in self.selected_items
            ]
            row_numbers = [
                self.current_page * self.page_size + self.tree.index(item) + 1
                for item in self.selected_items
            ]
            user = self.db_user
            password = self.db_password

            def delete_records():
                return delete_data_from_database_many(
                    primary_keys, user=user, password=password
                )

            self._start_database_task(
                delete_records,
                lambda results, error: self._finish_database_delete(
                    results,
                    error,
                    row_numbers,
                ),
            )
            return

    def _finish_database_delete(self, results, error, operated_row_numbers=None):
        if error:
            messagebox.showerror("数据库错误", str(error))
            self.db_status_label.configure(text="数据库删除失败")
            return
        failures = [message for success, message in results if not success]
        if failures:
            messagebox.showwarning("删除结果", "\n".join(failures))
        else:
            messagebox.showinfo("成功", "选中数据已删除！")
        successful_rows = [
            row_number
            for (success, _), row_number in zip(results, operated_row_numbers or [])
            if success
        ]
        if successful_rows:
            previous_row_number = min(successful_rows) - 1
            self._set_last_operated_row_number(previous_row_number)
            self.current_page = (
                (previous_row_number - 1) // self.page_size
                if previous_row_number > 0
                else 0
            )
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
            (
                self._tree_row_data[item][FIELD_INDEX["证书编号"]],
                self._tree_row_data[item][FIELD_INDEX["样品编号"]],
            )
            for item in self.selected_items
        ]
        user = self.db_user
        password = self.db_password

        def load_export_rows():
            return get_data_from_database_many(
                primary_keys, user=user, password=password
            )

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
