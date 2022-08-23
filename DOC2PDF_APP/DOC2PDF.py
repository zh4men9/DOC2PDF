# coding:utf-8
from utils.my_doc2pdf import my_doc2pdf
from utils.merge_pages import merge_pages
from utils.cut_pages import cut_pages
from utils.merge_pdf import merge_pdf

import tkinter
from tkinter.messagebox import *

import os
import yaml
import shutil


def get_args():

    # 获取yaml文件路径
    yamlPath = './config.yaml'
    args = None
    with open(yamlPath, 'rb') as f:
        # yaml文件通过---分节，多个节组合成一个列表
        date = yaml.safe_load_all(f)

        # salf_load_all方法得到的是一个迭代器，需要使用list()方法转换为列表
        args = list(date)[0]
    f.close()

    return args


def main():
    args = get_args()
    # print(args)
    doc2pdf_item = args['doc2pdf']
    merge_pages_item = args['merge_pages']
    cut_item = args['cut']
    merge_pdf_item = args['merge_pdf']

    # print(doc2pdf_item)
    # print(merge_pages_item)
    # print(cut_item)
    # print(merge_pdf_item)

    # doc2pdf args
    doc2pdf_input_path = doc2pdf_item['doc2pdf_input_path']
    doc2pdf_output_path = doc2pdf_item['doc2pdf_output_path']

    # merge pages args
    merge_pages_input_path = merge_pages_item['merge_pages_input_path']
    merge_pages_output_path = merge_pages_item['merge_pages_output_path']
    col_width = merge_pages_item['col_width']

    # cut args
    cut_input_path = cut_item['cut_input_path']
    cut_output_path = cut_item['cut_output_path']
    left_margin = cut_item['left_margin']
    right_margin = cut_item['right_margin']
    top_margin = cut_item['top_margin']
    bottom_margin = cut_item['bottom_margin']

    # merge pdf args
    merge_pdf_input_path = merge_pdf_item['merge_pdf_input_path']
    merge_pdf_output_path = merge_pdf_item['merge_pdf_output_path']
    merge_pdf_output_file_name = merge_pdf_item['merge_pdf_output_file_name']

    # 判断是否正确放置文件
    if (not os.path.exists(os.path.join(os.getcwd(), doc2pdf_input_path))):
        showerror(
            '错误', f'{os.path.join(os.getcwd(), doc2pdf_input_path)}文件不存在, 请创建!')
        return
    # 对于输出文件夹非空时, 删除文件夹里面文件
    if (os.path.exists(os.path.join(os.getcwd(), merge_pages_output_path))):
        shutil.rmtree(os.path.join(os.getcwd(), merge_pages_output_path))

    if (os.path.exists(os.path.join(os.getcwd(), doc2pdf_output_path))):
        shutil.rmtree(os.path.join(os.getcwd(), doc2pdf_output_path))

    if (os.path.exists(os.path.join(os.getcwd(), merge_pdf_output_path))):
        shutil.rmtree(os.path.join(os.getcwd(), merge_pdf_output_path))

    if (os.path.exists(os.path.join(os.getcwd(), cut_output_path))):
        shutil.rmtree(os.path.join(os.getcwd(), cut_output_path))

    showinfo('提示', '正在进行文档转换PDF, 请稍等...')

    my_doc2pdf(input_path=doc2pdf_input_path, output_path=doc2pdf_output_path)

    showinfo('提示', '文档转换PDF成功, 正在进行转换多栏, 请稍等...')

    merge_pages(input_path=merge_pages_input_path, output_path=merge_pages_output_path,
                col_width=col_width)

    showinfo('提示', '多栏转换成功, 正在进行页边距裁减, 请稍等...')

    cut_pages(input_path=cut_input_path, output_path=cut_output_path,
              left_margin=left_margin, right_margin=right_margin, top_margin=top_margin, bottom_margin=bottom_margin)

    showinfo('提示', '页边距裁减成功, 正在进行文件合并, 请稍等...')

    merge_pdf(input_path=merge_pdf_input_path, output_path=merge_pdf_output_path,
              output_file_name=merge_pdf_output_file_name)

    showinfo('提示', '文件合并成功, 程序执行结束, 正在退出...')


if __name__ == '__main__':
    main()
