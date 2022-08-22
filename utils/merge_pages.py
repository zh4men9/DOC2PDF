from PyPDF2 import PdfReader, PdfWriter, Transformation, PageObject, PdfFileReader, PdfFileWriter

def merge_pages(input_path='pdf', output_path='double_col'):
    file_name = r'D:\Files\MyOwnTools\DOC2PDF\pdf\概率机器学习大作业-周建桥_docx.pdf'
    reader = PdfFileReader(
        open(r"D:\Files\MyOwnTools\DOC2PDF\修改\修改概率机器学习大作业-周建桥_docx.pdf", 'rb'))
    # reader = PdfFileReader(open(r"D:\Files\MyOwnTools\DOC2PDF\pdf\2.pdf",'rb'))
    reader = PdfFileReader(open(file_name, 'rb'))

    two_pages_width = 841.92004 # word打印成双栏的大小
    two_pages_height = 595.32001
    writer = PdfFileWriter()

    for i in range(0, reader.numPages, 2):
        page_1 = reader.getPage(i)
        if (i+1 < reader.numPages):
            page_2 = reader.getPage(i+1)
        else:
            page_2 = PageObject.createBlankPage(
                None, two_pages_width/2, two_pages_height)

        # print(page_1.mediaBox.getWidth())
        # print(page_1.mediaBox.getHeight())

        # print(page_2.mediaBox.getWidth())
        # print(page_2.mediaBox.getHeight())

        # two_pages_width = page_1.mediaBox.getWidth()*2
        # two_pages_height = page_1.mediaBox.getHeight()

        # Creating a new file double the size of the original
        # translated_page = PageObject.createBlankPage(None, page_1.mediaBox.getWidth()*2, page_1.mediaBox.getHeight())
        translated_page = PageObject.createBlankPage(
            None, two_pages_width, two_pages_height)

        # Adding the pages to the new empty page
        translated_page.mergeScaledTranslatedPage(page_1, 1, -10, 0, 1)
        # translated_page.mergePage(page_1)
        translated_page.mergeScaledTranslatedPage(
            page_2, 1, float(two_pages_width/2)+10, 0, 1)

        writer.addPage(translated_page)

    with open('out.pdf', 'wb') as f:
        writer.write(f)

    # reader = PdfFileReader(open(r"D:\Files\MyOwnTools\DOC2PDF\out.pdf", 'rb'))
    # page = reader.getPage(0)
    print(page.mediabox.getHeight())
    print(page.mediabox.getWidth())
