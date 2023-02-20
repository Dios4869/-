import os
import openpyxl
from docx import Document
from docx.shared import Inches

# 定义模板目录路径
TEMPLATES_DIR = "模版/"

# 定义输出目录路径
OUTPUT_DIR = "输出/"

# 定义模板文件名和对应的输出文件名
TEMPLATE_FILENAMES = {
    "文件A.docx": "文件A.docx",
    "文件B.xlsx": "文件B.xlsx",
    "文件C.docx": "文件C.docx"
}

# 定义一个函数，根据给定的数据填充模板
def fill_template(template_filename, output_filename, data):
    # 根据模板文件的扩展名确定文件类型
    _, ext = os.path.splitext(template_filename)

    # 使用相应的库打开模板文件
    if ext == ".docx":
        document = Document(os.path.join(TEMPLATES_DIR, template_filename))
    elif ext == ".xlsx":
        workbook = openpyxl.load_workbook(os.path.join(TEMPLATES_DIR, template_filename))
        worksheet = workbook.active

    # 遍历数据字典，将占位符替换为对应的值
    for key, value in data.items():
        if ext == ".docx":
            # 处理Word文档中的占位符
            for paragraph in document.paragraphs:
                for run in paragraph.runs:
                    if key in run.text:
                        # 保存原始格式
                        original_format = run.font._element
                        # 替换文本并设置格式
                        new_text = run.text.replace(key, value)
                        run.text = new_text
                        run.font._element = original_format
            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            for run in paragraph.runs:
                                if key in run.text:
                                    # 保存原始格式
                                    original_format = run.font._element
                                    # 替换文本并设置格式
                                    new_text = run.text.replace(key, value)
                                    run.text = new_text
                                    run.font._element = original_format
        elif ext == ".xlsx":
            # 处理Excel表格中的占位符
            for row in worksheet.iter_rows():
                for cell in row:
                    if key in str(cell.value):
                        # 替换文本并设置格式
                        new_text = str(cell.value).replace(key, value)
                        cell.value = new_text

    # 保存生成的文件
    if ext == ".docx":
        document.save(os.path.join(OUTPUT_DIR, output_filename))
    elif ext == ".xlsx":
        workbook.save(os.path.join(OUTPUT_DIR, output_filename))

# 定义填充模板文件所需的数据
data = {
    "[NAME]": "John Smith",
    "[AGE]": "30",
    "[CITY]": "New York"
}

# 遍历模板文件，生成对应的输出文件
for template_filename, output_filename in TEMPLATE_FILENAMES.items():
    try:
        fill_template(template_filename, output_filename, data)
        print(f"{template_filename}已成功生成{output_filename}")
    except Exception as e:
        print(f"{template_filename}生成{output_filename}失败: {e}")
