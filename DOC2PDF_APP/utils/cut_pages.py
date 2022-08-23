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
