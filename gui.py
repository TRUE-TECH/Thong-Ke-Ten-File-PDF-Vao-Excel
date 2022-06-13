from tkinter import *
from tkinter import filedialog
import os
import xlsxwriter
import pikepdf

window = Tk()

window.title("get file name and count pages")

window.geometry('350x200')

selectfolder = filedialog.askdirectory(title="select folder") + '/'

selectexel = filedialog.askdirectory(title="select folder to create data.xlsx") + '/data.xlsx'
   
def get_data():

    src = os.listdir(selectfolder)
    data_pages = []

    workbook = xlsxwriter.Workbook(selectexel)
    worksheet = workbook.add_worksheet()

    row = 1
    column = 1

    row2 = 1
    column2 = 2

    for i in range(len(src)):
        file = pikepdf.Pdf.open(selectfolder+ src[i])
        totalpages = len(file.pages)
        data_pages.append(totalpages)

    for i in range(len(src)):
        worksheet.write(row, column, src[i])
        row += 1

    for i in range(len(data_pages)):
        worksheet.write(row2, column2, data_pages[i])
        row2 += 1

    workbook.close()
    window.destroy()

btn2 = Button(window, text="RUN",command= get_data)

btn2.pack(ipadx=5, pady=15)

window.mainloop()