from utils.my_doc2pdf import my_doc2pdf
from utils.merge_pages import merge_pages
from utils.cut_pages import cut_pages
from utils.merge_pdf import merge_pdf

import tkinter
from tkinter.messagebox import *


def get_args():
    pass


def main():

    # doc2pdf args
    doc2pdf_input_path = 'doc'
    doc2pdf_output_path = 'pdf'

    # merge pages args
    merge_pages_input_path = doc2pdf_output_path
    merge_pages_output_path = 'double_col'
    col_width = 10

    # cut args
    left_margin = 60
    right_margin = 70
    top_margin = 65
    bottom_margin = 75

    # merge pdf args
    merge_pdf_input_path = 'cut'
    merge_pdf_output_path = 'merge'
    merge_pdf_output_file_name = 'merged_pdf'

    showinfo('提示', '正在进行文档转换PDF, 请稍等...')

    my_doc2pdf(input_path=doc2pdf_input_path, output_path=doc2pdf_output_path)

    showinfo('提示', '文档转换PDF成功, 正在进行转换多栏, 请稍等...')

    merge_pages(input_path=merge_pages_input_path, output_path=merge_pages_output_path,
                col_width=col_width)

    showinfo('提示', '多栏转换成功, 正在进行页边距裁减, 请稍等...')

    cut_pages(input_path='double_col', output_path='cut',
              left_margin=left_margin, right_margin=right_margin, top_margin=top_margin, bottom_margin=bottom_margin)

    showinfo('提示', '页边距裁减成功, 正在进行文件合并, 请稍等...')

    merge_pdf(input_path=merge_pdf_input_path, output_path=merge_pdf_output_path,
              output_file_name=merge_pdf_output_file_name)

    showinfo('提示', '文件合并成功, 程序执行结束, 正在退出...')


if __name__ == '__main__':
    main()
