from PyPDF2 import PdfReader, PdfWriter, Transformation, PageObject, PdfFileReader, PdfFileWriter
import os


def merge_pages(input_path='pdf', output_path='double_col', pages=2,
                single_pages_width=420.96002, single_pages_height=595.32001, col_width=10, axis='X'):

    # 默认参数为word打印成双栏的长宽
    # 页面参数配置
    pages_width = single_pages_width*pages
    pages_height = single_pages_height

    input_path = os.path.join(os.getcwd(), input_path)
    # print(f'input_path:{input_path}')
    output_path = os.path.join(os.getcwd(), output_path)
    # print(f'output_path:{output_path}')

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
