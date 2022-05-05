import win32com.client
import xlwings as xw
import xlwt
from io import StringIO
from tkinter import *
from tkinter import filedialog
import fitz
from openpyexcel import load_workbook
from datetime import datetime
import time
import pyexcel
"""

start_time = datetime.now()
# ExcelApp = win32com.client.GetActiveObject("Excel.Application")
# ExcelApp.Visible = True
# print(ExcelApp .ActiveWorkbook.FullName )
# Создаем книку
book = xlwt.Workbook('utf8')
root = Tk()
root.geometry("300x200")
root.title("02_GRSH_01")

fil = filedialog.askopenfilename()
output_string = StringIO()
doc = fitz.open(fil)

for current_page in range(len(doc)):
    page = doc.load_page(current_page)
    page_text = page.get_text("text")
    text = page_text
text1 = text.split()
text2 = []
x = []
for z in text1:
    if z == 'Null':
        text2.append('0')
    elif z != 'Null':
        text2.append(z)

for item in text2:
    try:
        x.append(int(item))
    except ValueError as e:
        x.append(item)

wb = load_workbook('C:\\Users\\Alexandr\\Desktop\\zere.xlsx')
sheet = wb.worksheets[1]
for i in range(sheet.max_row, 0, -1):
    z = sheet['I' + str(i)].value
    if z is None:
        i = +1
    else:
        i = i + 1
        z = sheet['I' + str(i)].value = x[12]
        break
sheet = wb.worksheets[2]
for b in range(sheet.max_row, 0, -1):
    z = sheet['J' + str(b)].value
    if z is None:
        b = +1
    else:
        b = b + 1
        z = sheet['J' + str(b)].value = x[15]
        print(z)
        break

sheet = wb.worksheets[3]
for c in range(sheet.max_row, 0, -1):
    z = sheet['K' + str(c)].value
    if z is None:
        c = +1
    else:
        c = c + 1
        z = sheet['K' + str(c)].value = x[16]
        print(z)
        break

s = (datetime.now() - start_time)
wb.save('C:\\Users\\Alexandr\\Desktop\\zere.xlsx')
print('Время выполнения :', s)
"""
start_time = datetime.now()
book = xlwt.Workbook('utf8')
root = Tk()
root.geometry("300x200")
root.title("02_GRSH_01")



fil = filedialog.askopenfilename()
output_string = StringIO()
doc = fitz.open(fil)

for current_page in range(len(doc)):
                    page = doc.load_page(current_page)
                    page_text = page.get_text("text")
                    text = page_text
text1 = text.split()
text2 = []
x = []

for z in text1:
                    if z == 'Null':
                        text2.append('0')
                    elif z != 'Null':
                        text2.append(z)

for item in text2:
                    try:
                        x.append(int(item))
                    except ValueError as e:
                        x.append(item)

wb = load_workbook('C:\\Users\\Alexandr\\Desktop\\zere.xlsx')

def param(q, a, v, q1, a1, v1, q2, a2, v2):
    sheet = wb.worksheets[q]
    for i in range(sheet.max_row, 0, -1):
        z = sheet[a + str(i)].value
        if z is None:
            i = +1
        else:
            i = i + 1
            z = sheet[a + str(i)].value = x[v]
            print(z)
            break
    sheet = wb.worksheets[q1]
    for b in range(sheet.max_row, 0, -1):
        z = sheet[a1 + str(b)].value
        if z is None:
            b = +1
        else:
            b = b + 1
            z = sheet[a1 + str(b)].value = x[v1]
            print(z)
            break

    sheet = wb.worksheets[q2]
    for c in range(sheet.max_row, 0, -1):
        z = sheet[a2 + str(c)].value
        if z is None:
            c = +1
        else:
            c = c + 1
            z = sheet[a2 + str(c)].value = x[v2]
            print(z)
            break


    s = (datetime.now() - start_time)
    wb.save('C:\\Users\\Alexandr\\Desktop\\zere.xlsx')
    wb.close()
    print('Время выполнения :', s)


param(2, 'I', 12, 2, 'J', 14, 2, 'K', 8,)

