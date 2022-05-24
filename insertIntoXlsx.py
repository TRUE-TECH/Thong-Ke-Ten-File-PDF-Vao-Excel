import xlsxwriter
import pikepdf
import os

folder = r'D:\Auto\pdf\\'
xlsxdes = r'D:\Auto\data.xlsx'

src = os.listdir(folder)

data_pages = []

workbook = xlsxwriter.Workbook(xlsxdes)
worksheet = workbook.add_worksheet()

row = 1
column = 1

row2 = 1
column2 = 2

for i in range(len(src)):
    file = pikepdf.Pdf.open(folder+ src[i])
    totalpages = len(file.pages)
    data_pages.append(totalpages)

for i in range(len(src)):
    worksheet.write(row, column, src[i])
    row += 1

for i in range(len(data_pages)):
    worksheet.write(row2, column2, data_pages[i])
    row2 += 1

workbook.close()
