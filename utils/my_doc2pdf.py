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

    word = client.Dispatch("Word.Application")  # 打开word应用程序
    # for file in files:
    doc = word.Documents.Open(fn)  # 打开word文件
    # 另存为后缀为".pdf"的文件，其中参数17表示为pdf
    doc.SaveAs("{}_doc.pdf".format(save_path), 17)
    doc.Close()  # 关闭原来word文件
    word.Quit()

# 转换docx为pdf


def docx2pdf(fn, output_path='pdf'):
    # save_path为保存路径
    save_path = fn[:fn.rfind('\\')]
    save_path = save_path[:save_path.rfind('\\')]
    save_path = os.path.join(save_path, output_path)
    save_path = os.path.join(save_path, fn[fn.rfind('\\')+1:-5])

    word = client.Dispatch("Word.Application")  # 打开word应用程序
    # for file in files:
    doc = word.Documents.Open(fn)  # 打开word文件
    # 另存为后缀为".pdf"的文件，其中参数17表示为pdf
    doc.SaveAs("{}_docx.pdf".format(save_path), 17)
    doc.Close()  # 关闭原来word文件
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
