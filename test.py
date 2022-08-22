import win32print as wp

printer = wp.OpenPrinter('Microsoft Print to PDF')
print_job = wp.StartDocPrinter(printer, 1, (r"D:\Files\MyOwnTools\DOC2PDF\doc\test_doc.pdf", r"D:\Files\MyOwnTools\DOC2PDF\doc\test.pdf", 'RAW'))
wp.StartPagePrinter(printer)