[![Security Status](https://www.murphysec.com/platform3/v3/badge/1611146949546246144.svg?t=1)](https://www.murphysec.com/accept?code=c39119a3ae1ad45230415f6b78ea3b1f&type=1&from=2&t=2)

# DOC2PDF

正如DOC2PDF名字可见，这个工具将DOC文档转换成PDF文件，批量转换，解放双手。详细功能如下：

- 文档转换（doc -> pdf， docx -> pdf）
- 多页合并为单页
- 裁剪页边距
- 合并文件

# 1. 功能开发

## 1.1 文档转换

核心思想：利用`win32com`库的`client`打开word文档，利用word的保存为pdf功能实现文档转换。核心代码如下：

``` python
def doc2pdf(fn, output_path='pdf'):  

    # save_path为保存路径
    save_path = fn[:fn.rfind('\\')]
    save_path = save_path[:save_path.rfind('\\')]
    save_path = os.path.join(save_path, output_path)
    save_path = os.path.join(save_path, fn[fn.rfind('\\')+1:-4])
 
    word = client.Dispatch("Word.Application")  # 打开word应用程序
    # for file in files:
    doc = word.Documents.Open(fn)  # 打开word文件
    # 另存为后缀为".pdf"的文件，其中参数17表示为pdf
    doc.SaveAs("{}_doc.pdf".format(save_path), 17)
    doc.Close()  # 关闭原来word文件
    word.Quit()
  
# 转换docx为pdf
  
def docx2pdf(fn, output_path='pdf'):

    # save_path为保存路径
    save_path = fn[:fn.rfind('\\')]
    save_path = save_path[:save_path.rfind('\\')]
    save_path = os.path.join(save_path, output_path)
    save_path = os.path.join(save_path, fn[fn.rfind('\\')+1:-5])

    word = client.Dispatch("Word.Application")  # 打开word应用程序
    # for file in files:
    doc = word.Documents.Open(fn)  # 打开word文件
    # 另存为后缀为".pdf"的文件，其中参数17表示为pdf
    doc.SaveAs("{}_docx.pdf".format(save_path), 17)
    doc.Close()  # 关闭原来word文件
    word.Quit()
```

完整功能封装在`my_doc2pdf`函数中，函数需要两个参数：待转换文档路径`input_path`，以及保存转换后pdf路径`output_path`

完整`my_doc2pdf.py`文件内容如下

``` python
from win32com import client
import os
import sys

# 转换doc为pdf  
def doc2pdf(fn, output_path='pdf'):
    # save_path为保存路径
    save_path = fn[:fn.rfind('\\')]
    save_path = save_path[:save_path.rfind('\\')]
    save_path = os.path.join(save_path, output_path)
    save_path = os.path.join(save_path, fn[fn.rfind('\\')+1:-4])

    word = client.Dispatch("Word.Application")  # 打开word应用程序
    # for file in files:
    doc = word.Documents.Open(fn)  # 打开word文件
    # 另存为后缀为".pdf"的文件，其中参数17表示为pdf
    doc.SaveAs("{}_doc.pdf".format(save_path), 17)
    doc.Close()  # 关闭原来word文件
    word.Quit()

# 转换docx为pdf
def docx2pdf(fn, output_path='pdf'):

    # save_path为保存路径
    save_path = fn[:fn.rfind('\\')]
    save_path = save_path[:save_path.rfind('\\')]
    save_path = os.path.join(save_path, output_path)
    save_path = os.path.join(save_path, fn[fn.rfind('\\')+1:-5])

    word = client.Dispatch("Word.Application")  # 打开word应用程序
    # for file in files:
    doc = word.Documents.Open(fn)  # 打开word文件
    # 另存为后缀为".pdf"的文件，其中参数17表示为pdf
    doc.SaveAs("{}_docx.pdf".format(save_path), 17)
    doc.Close()  # 关闭原来word文件
    word.Quit()

# 获取路径下文件
def get_dir_list(input_path='doc'):

    dir_list = os.listdir(input_path)
    # print(dir_list)
    return dir_list

def my_doc2pdf(input_path='doc', output_path='pdf'):

    dir_list = get_dir_list(input_path)
    cwd_path = os.getcwd()
    cwd_path_input = os.path.join(cwd_path, input_path)
    cwd_path_output = os.path.join(cwd_path, output_path)

    if (not os.path.exists(cwd_path_output)):
        os.makedirs(cwd_path_output)

    for dir_i in dir_list:

        doc_path = os.path.join(cwd_path_input, dir_i)
        # print(doc_path)
        # 判断文件后缀是否是doc或者docx
        if (doc_path.endswith('doc')):
            print(doc_path)
            doc2pdf(doc_path, output_path)
        elif (doc_path.endswith('docx')):
            print(doc_path)
            docx2pdf(doc_path, output_path)

if __name__ == '__main__':

    my_doc2pdf(input_path='doc', output_path='pdf')
```

## 1.2 多页合并为单页

核心思路如下：利用`PageObject.createBlankPage`创建期望长宽的空白页，然后利用`mergeScaledTranslatedPage`将不同页的文件设置到对应位置上。

完整功能封装为`merge_pages`函数，函数需要7个参数：
- 待转换pdf文件路径`input_path`
- 保存转换完成文件路径`output_path`
- 合并文件单页包含的页数`pages`
- 待转换文件单页宽度`single_pages_width`
- 待转换文件单页长度`single_pages_height`
- 每页的距离`col_width`
- 合并方向`axis`（纵向合并和横向合并）

目前只实现2页横向合并为1页功能，完整`merge_pages.py`如下：

``` python
from PyPDF2 import PdfReader, PdfWriter, Transformation, PageObject, PdfFileReader, PdfFileWriter

import os

def merge_pages(input_path='pdf', output_path='double_col', pages=2,
                single_pages_width=420.96002, single_pages_height=595.32001, col_width=10, axis='X'):

    # 默认参数为word打印成双栏的长宽
    # 页面参数配置
    pages_width = single_pages_width*pages
    pages_height = single_pages_height
  
    input_path = os.path.join(os.getcwd(), input_path)
    print(f'input_path:{input_path}')
    output_path = os.path.join(os.getcwd(), output_path)
    print(f'output_path:{output_path}')

    if (not os.path.exists(output_path)):
        os.makedirs(output_path)

    for dirpath, dirnames, filenames in os.walk(input_path):
        for file_name in filenames:
            input_file_path = os.path.join(dirpath, file_name)
            output_file_path = os.path.join(output_path, file_name)

            reader = PdfFileReader(open(input_file_path, 'rb'))
            writer = PdfFileWriter()
            
            for i in range(0, reader.numPages, 2):
                page_1 = reader.getPage(i)
                if (i+1 < reader.numPages):
                    page_2 = reader.getPage(i+1)
                else:
                    page_2 = PageObject.createBlankPage(
                        None, pages_width/2, pages_height)
                # print(page_1.mediaBox.getWidth())
                # print(page_1.mediaBox.getHeight())
                # print(page_2.mediaBox.getWidth())
                # print(page_2.mediaBox.getHeight())

                # pages_width = page_1.mediaBox.getWidth()*2
                # pages_height = page_1.mediaBox.getHeight()

                # Creating a new file double the size of the original

                translated_page = PageObject.createBlankPage(
                    None, pages_width, pages_height)
                # Adding the pages to the new empty page

                translated_page.mergeScaledTranslatedPage(
                    page_1, 1, -col_width, 0, 1)
                translated_page.mergeScaledTranslatedPage(
                    page_2, 1, float(pages_width/2)+col_width, 0, 1)

                writer.addPage(translated_page)

            with open(output_file_path, 'wb') as f:
                writer.write(f)

if __name__ == '__main__':

    merge_pages(input_path='pdf', output_path='double_col')
```

## 1.3 裁剪页边距

核心思路：利用`page.mediaBox`设置每页4个点参数，封装为`split`函数，如下：

``` python
def split(page, width, height, left_margin=60, right_margin=70, top_margin=65, bottom_margin=75):

    page.mediaBox.lowerLeft = (left_margin, bottom_margin)
    page.mediaBox.lowerRight = (width - right_margin, bottom_margin)
    page.mediaBox.upperLeft = (left_margin, height - top_margin)
    page.mediaBox.upperRight = (width - right_margin, height - top_margin)
```

完整功能封装为`cut_pages`函数，函数需要6个参数：
- 待转换pdf文件路径`input_path`
- 保存转换完成文件路径`output_path`
- 裁剪左边距`left_margin`
- 裁剪右边距`right_margin`
- 裁剪上边距`top_margin`
- 裁剪下边距`bottom_margin`

完整`cut_pages.py`如下：

``` python
# 裁减页边距

import PyPDF2
import os

def split(page, width, height, left_margin=60, right_margin=70, top_margin=65, bottom_margin=75):

    page.mediaBox.lowerLeft = (left_margin, bottom_margin)
    page.mediaBox.lowerRight = (width - right_margin, bottom_margin)
    page.mediaBox.upperLeft = (left_margin, height - top_margin)
    page.mediaBox.upperRight = (width - right_margin, height - top_margin)

def cut_pages(input_path='double_col', output_path='cut',
              left_margin=60, right_margin=70, top_margin=65, bottom_margin=75):

    input_file_path = []
    output_file_path = []

    input_path = os.path.join(os.getcwd(), input_path)
    output_path = os.path.join(os.getcwd(), output_path)

    if (not os.path.exists(output_path)):
        os.makedirs(output_path)

    file_list = os.listdir(input_path)
    for i in file_list:
        if os.path.splitext(i)[1] == '.pdf':
            input_file_path.append(os.path.join(input_path, i))
            output_file_path.append(os.path.join(output_path, i))

    for m in range(len(input_file_path)):
        input_file = PyPDF2.PdfFileReader(open(input_file_path[m], 'rb'))
        output_file = PyPDF2.PdfFileWriter()

        page_info = input_file.getPage(0)
        width = float(page_info.mediaBox.getWidth())
        height = float(page_info.mediaBox.getHeight())
        page_count = input_file.getNumPages()

        for page_num in range(page_count):
            this_page = input_file.getPage(page_num)
            split(this_page, left_margin=left_margin, right_margin=right_margin,
                  top_margin=top_margin, bottom_margin=bottom_margin, width=width, height=height)
            output_file.addPage(this_page)

        output_file.write(open(output_file_path[m], 'wb'))
```

## 1.4 合并文件

核心思路：读取每个`input_path`下的pdf文件，循环将每一页保存到一个新的`PdfFileWriter`里，最后保存即可。

完整代码如下：

``` python
import os
from PyPDF2 import PdfMerger, PdfFileWriter, PdfFileReader

def merge_pdf(input_path='cut', output_path='merge', output_file_name='merged_pdf'):

    cwd_path = os.getcwd()
    input_path = os.path.join(cwd_path, input_path)
    output_path = os.path.join(cwd_path, output_path)
    output_file_name = os.path.join(output_path, output_file_name) + '.pdf'

    if (not os.path.exists(output_path)):
        os.makedirs(output_path)

    input_file_list = []
    for dirpath, _, filenames in os.walk(input_path):

        for pdf_file in filenames:
            if (os.path.splitext(pdf_file)[1] == '.pdf'):
                input_file_list.append(os.path.join(dirpath, pdf_file))
  
        writer = PdfFileWriter()
        for single_pdf_file_path in input_file_list:
            reader = PdfFileReader(single_pdf_file_path)

            for page in reader.pages:
                writer.add_page(page)

    writer.write(open(output_file_name, 'wb'))
```

# 2. 开源资料

本项目GitHub仓库：[zh4men9/DOC2PDF (github.com)](https://github.com/zh4men9/DOC2PDF)

参考资料：
- [python - Concatenating multiple page pdf into single page pdf - Stack Overflow](https://stackoverflow.com/questions/59348564/concatenating-multiple-page-pdf-into-single-page-pdf)
- [python - 如何将pdf文件中的两页合并为一页 - IT工具网 (coder.work)](https://www.coder.work/article/5053219)
- [Python获取路径下所有文件名 - skaarl - 博客园 (cnblogs.com)](https://www.cnblogs.com/skaarl/p/10316564.html)
- [(10条消息) 用python裁剪PDF文档_煎饼果子cxk的博客-CSDN博客_python裁剪pdf](https://blog.csdn.net/weixin_44712173/article/details/123769341)

更详细介绍见博客：[(10条消息) DOC2PDF项目博客_zh4men9的博客-CSDN博客](https://blog.csdn.net/qq_32614873/article/details/126470122?spm=1001.2014.3001.5502)

# 使用

为方便使用工具，已将工具打包，打包后的文件在 `DOC2PDF_APP`中

打包命令: `pyinstaller -D -p DOC2PDF_APP DOC2PDF.py`

使用说明见 `使用说明.txt`

## TO DO

* [ ] 多页合并
* [ ] 选择文件进行操作
* [ ] 支持其他文件转PDF（图片、PPT等）
