# 裁减页边距
# https://blog.csdn.net/weixin_44712173/article/details/123769341

import PyPDF2
import os

def split(page):
    page.mediaBox.lowerLeft = (left_margin, bottom_margin)
    page.mediaBox.lowerRight = (width - right_margin, bottom_margin)
    page.mediaBox.upperLeft = (left_margin, height - top_margin)
    page.mediaBox.upperRight = (width - right_margin, height - top_margin)


left_margin = 60
right_margin = 70
top_margin = 65
bottom_margin = 75

input_file_path = []
output_file_path = []
# os.makedirs('.\\修改')  # 新建文件夹

file_path = '.\\pdf'
file_list = os.listdir(file_path)
for i in file_list:
    if os.path.splitext(i)[1] == '.pdf':
        input_file_path.append('.\\pdf\\' + i)
        output_file_path.append('.\\修改\\' + '修改' + i)

for m in range(len(input_file_path)):
    input_file = PyPDF2.PdfFileReader(open(input_file_path[m], 'rb'))
    output_file = PyPDF2.PdfFileWriter()

    page_info = input_file.getPage(0)
    width = float(page_info.mediaBox.getWidth())
    height = float(page_info.mediaBox.getHeight())
    page_count = input_file.getNumPages()

    for page_num in range(page_count):
        this_page = input_file.getPage(page_num)
        split(this_page)
        output_file.addPage(this_page)

    output_file.write(open(output_file_path[m], 'wb'))
