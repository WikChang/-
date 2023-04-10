import os
from docx import Document
import pandas as pd
import PyPDF2
import shutil


def split_and_save_word_pages(input_docfile,input_pdffile, output_path):
    # 加载输入的Word文档
    doc = Document(input_docfile)
    namelist = []
    for table in doc.tables:
        name = table.cell(0, 10).text
        namelist.append(name)
    # 遍历输入文档的段落
    with open(input_pdffile, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        try:
            for page_num in range(len(pdf_reader.pages)):
                pdf_writer = PyPDF2.PdfWriter()
                pdf_writer.add_page(pdf_reader.pages[page_num])
                print(page_num)
                output_path = os.path.join(output_path, f'硕士研究生复试评定表_{namelist[page_num]}.pdf')

                with open(output_path, 'wb') as output_pdf:
                    pdf_writer.write(output_pdf)
        except:
            pass
    print(namelist)
def move_files_to_folders(src_folder, excel_file, output_path):
    # 读取Excel文件
    df = pd.read_excel(excel_file, engine='openpyxl')

    # 遍历每一行
    for index, row in df.iterrows():
        name = row['考生姓名']
        location = row['复试地点']
        time = row['简略时间']
        index = row['序号']
        # 创建目标文件夹
        folder_name = f'{location}-{time}'
        folder_path = os.path.join(output_path, folder_name)

        # 如果文件夹不存在，则创建它
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        # 构建源文件路径和目标文件路径
        src_file = os.path.join(src_folder, f'硕士研究生复试评定表_{name}.pdf')
        dst_file = os.path.join(folder_path, f'{str(index)}_硕士研究生复试评定表_{name}.pdf')

        # 移动文件
        if os.path.exists(src_file):
            shutil.move(src_file, dst_file)

# 将文件转换为pdf并分割，需要注意的是，要将评定表转换为pdf，并保存在input_file文件夹中
input_docfile = './input_file/考生评定表.docx'
output_path = './output_file/硕士研究生复试评定表/temp'
input_pdffile =  './input_file/考生评定表.pdf'
split_and_save_word_pages(input_docfile,input_pdffile ,output_path)
#将同一考场的学生放在一个文件夹中
src_folder = './output_file/硕士研究生复试评定表/temp'
excel_file = './input_file面试名单.xlsx'
output_path = './output_file/硕士研究生复试评定表'
move_files_to_folders(src_folder, excel_file, output_path)