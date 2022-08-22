import win32print

printer_name = 'Microsoft Print to PDF'
# printer_name = 'Adobe PDF'

# print(win32print.EnumPrinters(2))  # (flags, description, name, comment)
# ((8388608, '发送至 OneNote 16,Send to Microsoft OneNote 16 Driver,', '发送至 OneNote 16', ''),
# (8388608, 'Microsoft XPS Document Writer,Microsoft XPS Document Writer v4,', 'Microsoft XPS Document Writer', ''),
# (8388608, 'Microsoft Print to PDF,Microsoft Print To PDF,', 'Microsoft Print to PDF', ''),
# (8388608, 'Foxit PDF Reader Printer,Foxit PDF Reader Printer Driver,', 'Foxit PDF Reader Printer', ''),
# (8388608, 'Fax,Microsoft Shared Fax Driver,', 'Fax', ''),
# (8388608, 'Adobe PDF,Adobe PDF Converter,', 'Adobe PDF', ''))

printer = win32print.OpenPrinter(printer_name)
d = win32print.GetPrinter(printer, 2)
devmode = d['pDevMode']

# print('Status ', d['Status'])
# for n in dir(devmode):
#     print("%s\t%s" % (n, getattr(devmode, n)))
# if d[18]:
#     print("Printer not ready")
# print(':'.join(x.encode('hex') for x in devmode.DriverData))

win32print.SetPrinter(printer, 2, d, 0)

hDC = win32ui.CreateDC()
hDC.CreatePrinterDC(printer_name)
printable_area = hDC.GetDeviceCaps(HORZRES), hDC.GetDeviceCaps(VERTRES)
printer_size = hDC.GetDeviceCaps(
    PHYSICALWIDTH), hDC.GetDeviceCaps(PHYSICALHEIGHT)
printer_margins = hDC.GetDeviceCaps(
    PHYSICALOFFSETX), hDC.GetDeviceCaps(PHYSICALOFFSETY)
#printable_area = (350*3, 1412*3)
#printer_size = (350*3, 1412*3)
print("printer area", printable_area)
print("printer size", printer_size)
print("printer margins", printer_margins)

try:
    hDC.StartDoc(file_name)
    hDC.StartPage()

    dib = ImageWin.Dib(bmp)

    dib.draw(hDC.GetHandleOutput(), (x1, y1, x2, y2))

    hDC.EndPage()
    hDC.EndDoc()
    hDC.DeleteDC()
except win32ui.error as e:
    print("Unexpected error:", e)
