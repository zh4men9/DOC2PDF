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
    for dirpath, _ , filenames in os.walk(input_path):
        
        for pdf_file in filenames:
            if (os.path.splitext(pdf_file)[1]=='.pdf'):
                input_file_list.append(os.path.join(dirpath, pdf_file))
        
        writer = PdfFileWriter()
        for single_pdf_file_path in input_file_list:
            reader = PdfFileReader(single_pdf_file_path)
            
            for page in reader.pages:
                writer.add_page(page)
    
    writer.write(open(output_file_name, 'wb'))
