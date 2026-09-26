from docx import Document
import openpyxl
import os
from datetime import datetime

# 核心配置（按需修改路径）
TEMPLATE_PATH = os.path.join(os.path.expanduser("~"), "Desktop", "template.docx")  # 模板路径
EXCEL_PATH = os.path.join(os.path.expanduser("~"), "Desktop", "数据收集表.xlsx")    # 数据来源
OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "Desktop", "填充后的文档")       # 输出文件夹

def init_output_dir():
    """初始化输出文件夹（不存在则创建）"""
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

def read_excel_data():
    """读取Excel数据，返回字典列表（每个字典对应一行数据）"""
    data_list = []
    # 检查Excel文件是否存在
    if not os.path.exists(EXCEL_PATH):
        print(f"❌ 未找到Excel文件：{EXCEL_PATH}")
        return data_list
    
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active
    max_row = ws.max_row
    
    # 无数据（仅表头）
    if max_row <= 1:
        print("⚠️ Excel中暂无收集的数据！")
        wb.close()
        return data_list
    
    # 读取表头（作为字典的键）
    headers = [ws.cell(row=1, column=col).value for col in range(1, ws.max_column + 1)]
    # 读取每一行数据，转为字典
    for row_num in range(2, max_row + 1):
        row_data = {}
        for col_num in range(1, ws.max_column + 1):
            # 空值替换为「未填写」
            cell_value = ws.cell(row=row_num, column=col_num).value or "未填写"
            row_data[headers[col_num - 1]] = str(cell_value)
        data_list.append(row_data)
    
    wb.close()
    return data_list

def replace_placeholder_in_paragraph(paragraph, data):
    """替换单个段落中的所有占位符"""
    for run in paragraph.runs:
        # 遍历占位符并替换
        for placeholder, value in data.items():
            placeholder_tag = f"{{{{{placeholder}}}}}"  # 拼接成{{字段名}}格式
            if placeholder_tag in run.text:
                run.text = run.text.replace(placeholder_tag, value)

def fill_word_template(data):
    """填充单个数据到Word模板，生成独立文档"""
    # 检查模板是否存在
    if not os.path.exists(TEMPLATE_PATH):
        print(f"❌ 未找到Word模板：{TEMPLATE_PATH}")
        return False
    
    # 打开模板
    doc = Document(TEMPLATE_PATH)
    
    # 1. 替换普通段落中的占位符（核心逻辑）
    for paragraph in doc.paragraphs:
        replace_placeholder_in_paragraph(paragraph, data)
    
    # 2. 替换文本框中的占位符（可选，若模板有文本框需保留）
    for shape in doc.element.body.iter():
        if shape.tag.endswith('textbox'):
            for paragraph in shape:
                if hasattr(paragraph, 'runs'):
                    for run in paragraph.runs:
                        for placeholder, value in data.items():
                            placeholder_tag = f"{{{{{placeholder}}}}}"
                            if placeholder_tag in run.text:
                                run.text = run.text.replace(placeholder_tag, value)
    
    # 3. 生成唯一的输出文件名（按姓名+时间命名）
    file_name = f"数据收集单_{data['姓名']}_{datetime.now().strftime('%Y%m%d%H%M%S')}.docx"
    output_path = os.path.join(OUTPUT_DIR, file_name)
    
    # 4. 保存填充后的文档
    try:
        doc.save(output_path)
        print(f"✅ 生成文档：{output_path}")
        return True
    except Exception as e:
        print(f"❌ 保存失败（姓名：{data['姓名']}）：{str(e)}")
        return False

if __name__ == "__main__":
    # 初始化输出文件夹
    init_output_dir()
    
    # 读取Excel数据
    data_list = read_excel_data()
    if not data_list:
        input("按回车键退出...")
        exit()
    
    # 批量填充模板（每条数据生成一个文档）
    success_count = 0
    for data in data_list:
        if fill_word_template(data):
            success_count += 1
    
    # 输出统计结果
    print(f"\n📊 批量生成完成！总计{len(data_list)}条数据，成功生成{success_count}个文档")
    print(f"📁 所有文档已保存到：{OUTPUT_DIR}")
    input("按回车键退出...")
