from utils.my_doc2pdf import my_doc2pdf
from utils.merge_pages import merge_pages
from utils.cut_pages import cut_pages


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
    left_margin=60
    right_margin=70
    top_margin=65
    bottom_margin=75

    # my_doc2pdf(input_path=doc2pdf_input_path, output_path=doc2pdf_output_path)
    
    # merge_pages(input_path=merge_pages_input_path, output_path=merge_pages_output_path,
    # col_width=col_width)
    
    cut_pages(input_path='double_col', output_path='cut',
              top_margin=top_margin, bottom_margin=bottom_margin, width=width, height=height)


if __name__ == '__main__':
    main()
