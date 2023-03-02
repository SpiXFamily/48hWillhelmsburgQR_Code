import xlsxwriter
import pyqrcode

# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('data.xls')
worksheet = workbook.add_worksheet()

# Generate QR codes and write them to the worksheet.
data = 'Example text to encode'
qr = pyqrcode.create(data)
worksheet.insert_image('A1', 'data.png', {'image_data': qr.png('data.png', scale=6)})

workbook.close()
