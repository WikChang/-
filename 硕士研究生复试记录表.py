import os
import pandas as pd
from docx import Document
from docxcompose.composer import Composer

def merge_documents(documents, output_path):
    merged_document = Composer(documents[0])

    for document in documents[1:]:
        merged_document.append(document)

    merged_document.save(output_path)
def replace_field_with_value(table, field, value):
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            if field in cell.text:
                # 检查是否有右侧单元格，如果有则填充数据
                if idx + 1 < len(row.cells):
                    row.cells[idx + 1].text = str(value)
                break

def create_and_merge_word_documents(template_path, excel_path, output_path):
    # 读取Excel文件
    df = pd.read_excel(excel_path)

    documents_by_folder = {}

    # 遍历Excel中的每一行数据
    for index, row in df.iterrows():
        # 读取模板文件
        doc = Document(template_path)

        # 将模板中的字段替换为Excel中的数据
        for table in doc.tables:
            replace_field_with_value(table, "考生姓名", row["考生姓名"])
            replace_field_with_value(table, "考生编号", row["考生编号"])
            replace_field_with_value(table, "复试专业名称", row["复试专业名称"])
            replace_field_with_value(table, "复试专业代码", "0"+str(row["复试专业代码"]))
            replace_field_with_value(table, "复试时间", row["复试时间"])
            replace_field_with_value(table, "复试地点", row["复试地点"])

            folder_name = str(row["复试地点"])+'-'+str(row["简略时间"])
            folder_path = os.path.join(output_path, folder_name)
            # 如果文件夹不存在，则创建
            if not os.path.exists(folder_path):
                os.makedirs(folder_path)
            # 生成文件名
            filename =  f"{row['序号']}-华南农业大学2023年硕士研究生复试记录表-{row['复试地点']}-{row['简略时间']}-{row['考生姓名']}.docx"
            output_path = os.path.join(folder_path, filename)

            # 保存生成的Word文件
            doc.save(output_path)
        # 获取复试时间文件夹名称
        folder_name = str(row["复试地点"])+'-'+str(row["简略时间"])

        if folder_name not in documents_by_folder:
            documents_by_folder[folder_name] = []

        documents_by_folder[folder_name].append(doc)

    # 按文件夹合并文档
    for folder_name, documents in documents_by_folder.items():
        output_path = os.path.join(output_path, f"{folder_name}.docx")
        merge_documents(documents, output_path)

# 调用函数，指定模板、Excel和输出文件夹路径
template_path = "./input_file/硕士研究生复试记录表.docx"
excel_path = "./input_file/面试名单.xlsx"
output_path = "./output_file/硕士研究生复试记录表"

create_and_merge_word_documents(template_path, excel_path, output_path)
