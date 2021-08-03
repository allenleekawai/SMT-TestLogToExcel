# -*- coding: UTF-8 -*-

print("\nLoading...\n")

import os
import time
from numpy import character
from numpy.lib.function_base import average
import openpyxl
import string

from tkinter import *
from tkinter import messagebox
from tkinter.filedialog import askdirectory, askopenfilename

from openpyxl import Workbook
from openpyxl.styles import Font
from urllib.parse import unquote, urlparse
from openpyxl.styles.colors import Color

def selectFolderPath():
    folderPath_ = askdirectory()
    folderPath.set(folderPath_)

def selectExcelPath():
    excelPath_ = askopenfilename()
    excelPath.set(excelPath_)

def pop_up():
    print("========================================\n")
    print("此為格式錯誤之檔案\n")
    messagebox.showinfo("Warning", wrongFormat)

def printString(n):
    MAX = 304
    # To store result (Excel column name)
    string = ["\0"] * MAX
    # To store current index in str which is result
    i = 0
    rem = 0
    while n >= 1:
        # Find remainder
        rem = n % 26
        rem = int(rem)
        # if remainder is 0, then a
        # 'Z' must be there in output
        if rem == 0:
            string[i] = 'Z'
            i += 1
            n = (n / 26) - 1
        else:
            string[i] = chr((rem - 1) + ord('A'))
            i += 1
            n = n / 26
    string[i] = '\0'

    # Reverse the string and print result
    string = string[::-1]
    return str("".join(string))

root = Tk()
root.title("檔案選擇")
root.resizable(0,0)
root.geometry('500x53')

folderPath = StringVar()
excelPath  = StringVar()

Label(root, text = "資料夾路徑").grid(row = 0, column = 0)
Entry(root, textvariable = folderPath).grid(row = 0, column = 1, ipadx = 126)
Button(root, text = "選擇", command = selectFolderPath).grid(row = 0, column = 2)

Label(root, text = "Excel路徑").grid(row = 1, column = 0)
Entry(root, textvariable = excelPath).grid(row = 1, column = 1, ipadx = 126)
Button(root, text = "選擇", command = selectExcelPath).grid(row = 1, column = 2)

print("請選擇 資料夾路徑 及 Excel 路徑\n")
print("========================================\n")
root.mainloop()

# 讀取資料夾所有檔案，放入陣列中 ( 會判別檔案形式是否正確，若錯誤則從陣列中刪除 )
allFileList = os.listdir(folderPath.get())
fileCount = len(allFileList)

print("讀取資料夾檔案中...\n")

for i in range(len(allFileList)):
    if allFileList[i][-4:] != '.dcl':
        fileCount -= 1

for i in range(fileCount):
    if allFileList[i][-4:] != '.dcl':
        del allFileList[i]
        i -= 1

print("符合條件的 Log，有 %d 筆\n" % (fileCount))
print("========================================\n")

"""
# 打開 Excel，進行 Initial 作業
excelinit = openpyxl.load_workbook(unquote(urlparse(excelPath.get()).path)) #打開Excel文件
for i in range(len(excelinit.sheetnames)):
    excelinit.remove(excelinit['' + excelinit.sheetnames[0]])
excelinit.create_sheet('Data', 0)
excelinit.save(unquote(urlparse(excelPath.get()).path))
"""

# 打開 Excel，進行 Initial 作業
print("打開 Excel 並進行初始化\n")
excelsave = openpyxl.load_workbook(unquote(urlparse(excelPath.get()).path))
sheetsave = excelsave['Data']
#sheet.merge_cells(start_row = 5, start_column = 1, end_row = sheet.max_row, end_column = sheet.max_column)
#sheet.unmerge_cells(start_row = 5, start_column = 1, end_row = sheet.max_row, end_column = sheet.max_column)
while(sheetsave.max_row > 4):
    sheetsave.delete_rows(5)
while(sheetsave.max_column > 303):
    sheetsave.delete_cols(304)
excelsave.save(unquote(urlparse(excelPath.get()).path))
excel = openpyxl.load_workbook(unquote(urlparse(excelPath.get()).path))
sheet = excel['Data']
sheetValue = excel.copy_worksheet(sheet)
sheetValue.title = 'Value'
#sheetValue = excel['Value']

# 建立標題行、結尾空行、運算行
indextitle = ["General Data", "", "ICT Test Date & Time", "", ""]
listtitle  = ["No.", "S/N", "Date", "Time", "Test Pass / Fail"]
end        = []
passTotal  = ["", "", "", "Total Pass"]
failTotal  = ["", "", "", "Total Fail"]
Yield      = ["", "", "", "Total Yield"]
formula    = ["", "", "", "", "Average"]

titleArray2D = []
with open(folderPath.get() + "/" + allFileList[0], 'r+') as f:
    for line in f.readlines():
        titleArray2D.append(line.split(','))

del titleArray2D[0]
del titleArray2D[0]

for i in range(len(titleArray2D)):
    indextitle.append(titleArray2D[i][0][1:])
    listtitle.append(titleArray2D[i][1][1:])

#sheet.append(indextitle)
#sheet.append(listtitle)

passCount   = 1
errorCount  = 0
wrongFormat = []

print("資料寫入中...\n")

# 進行 Log 資料處理，將所有資料以 "," 分開，並放進二維陣列中
for i in range(fileCount):
    global listadd, listtitleCheck, listerror, listvalue
    listadd           = {}
    listadd[i]        = []
    listtitleCheck    = {}
    listtitleCheck[i] = ["No.", "S/N", "Date", "Time", "Test Pass / Fail"]
    listerror         = {}
    listerror[i]      = []

    listvalue         = {}
    listvalue[i]      = []

    array2D = []
    with open(folderPath.get() + "/" + allFileList[i], 'r+') as f:
        print("-", end = "")
        for line in f.readlines():
            array2D.append(line.split(','))

    listadd[i].append(passCount)
    listadd[i].append(array2D[0][4])
    listadd[i].append(array2D[0][5][0:4] + "/" + array2D[0][5][4:6] + "/" + array2D[0][5][6:8])
    listadd[i].append(array2D[0][6][0:2] + ":" + array2D[0][6][2:4] + ":" + array2D[0][6][4:6])
    if array2D[0][0] == "PASS": listadd[i].append("Pass")
    else:                       listadd[i].append("Fail")

    listvalue[i].append(passCount)
    listvalue[i].append(array2D[0][4])
    listvalue[i].append(array2D[0][5][0:4] + "/" + array2D[0][5][4:6] + "/" + array2D[0][5][6:8])
    listvalue[i].append(array2D[0][6][0:2] + ":" + array2D[0][6][2:4] + ":" + array2D[0][6][4:6])
    listvalue[i].append("")
    
    del array2D[0]
    del array2D[0]

    #for x in range(len(array2D)):
    #    if array2D[x][0:4] != "Short" and array2D[i][0:3] != "Open":
    #        correctCount += 1

    for x in range(len(array2D)):
        if array2D[x][0][0:5] == "Short" or array2D[x][0][0:4] == "Open":
            listerror[errorCount] = array2D[-1]
            del array2D[-1]
            del array2D[-1]
            #print(listerror[y])
            errorCount += 1

    #print(array2D)
    
    for j in range(len(array2D)):
        listtitleCheck[i].append(array2D[j][1][1:])
        if array2D[j][11] == " 0\n":    listadd[i].append("Pass")
        else:                           listadd[i].append("Fail")
        #print(array2D[i][11], end='')
        listvalue[i].append(float(array2D[j][9][1:-2]))
        #print(array2D[j][9][1:-2])

    if len(listerror) != -1:
        for z in range((len(listerror) - 1)):
            listadd[i].append("".join(listerror[z]))

    if listtitleCheck[i] == listtitle:
        passCount += 1
        sheet.append(listadd[i])
        sheetValue.append(listvalue[i])
    else:
        wrongFormat.append(allFileList[i])

#print(sheet.max_column)
if sheet.max_column == 303:
    for i in range(5, (sheet.max_column + 1)):
        passCounter = 0
        failCounter = 0
        for j in range(5, passCount + 4):
            if sheet.cell(column = i, row = j).value == "Pass": passCounter += 1
            else:                                               failCounter += 1
        passTotal.append(passCounter)
        failTotal.append(failCounter)
        yieldnum = passCounter / (passCounter + failCounter) * 100
        yieldnum = ('%.3f' % yieldnum)
        Yield.append(yieldnum + "%")
        if (passCounter + failCounter) != (passCount - 1):  print("Something Wrong!")
else:
    for i in range(5, (sheet.max_column)):
        passCounter = 0
        failCounter = 0
        for j in range(5, passCount + 4):
            if sheet.cell(column = i, row = j).value == "Pass": passCounter += 1
            else:                                               failCounter += 1
        passTotal.append(passCounter)
        failTotal.append(failCounter)
        yieldnum = passCounter / (passCounter + failCounter) * 100
        yieldnum = ('%.3f' % yieldnum)
        Yield.append(yieldnum + "%")
        if (passCounter + failCounter) != (passCount - 1):  print("Something Wrong!")
    Yield.append("Short or Open : %d" % errorCount)

sheet.append(end)
sheet.append(passTotal)
sheet.append(failTotal)
sheet.append(Yield)

sheetValue.append(end)
sheetValue.append(formula)

#sheet.merge_cells('A1:B1')
#sheet.merge_cells('C1:D1')
#sheet.merge_cells('A' + str(sheet.max_row - 2) + ':D' + str(sheet.max_row - 2))
#sheet.merge_cells('A' + str(sheet.max_row - 1) + ':D' + str(sheet.max_row - 1))
#sheet.merge_cells('A' + str(sheet.max_row) + ':D' + str(sheet.max_row))

if sheetValue.max_column == 303:
    for i in range(6, (sheet.max_column + 1)):
        #cell_column = '' + printString(i)
        #cell_top = '%s' % (str(5))
        #top = cell_column + cell_top
        top = "F5"
        #cell_buttom = '%s' % (str(passCount + 3))
        #buttom = cell_column + cell_buttom
        buttom = 'F15'
        #cell_average = '%s' % (str(passCount + 5))
        #Average = cell_column + cell_average
        #cell_average = "F17"
        #print(cell_top)
        #print(cell_buttom)
        #print(cell_average)
        #sheetValue[cell_average] = "=AVERAGE({}:{})".format(cell_top, cell_buttom)
        #sheetValue[cell_average] = "=AVERAGE(F5:F15)"
        celltop = sheetValue.cell(row = 5, column = i)
        cellbuttom = sheetValue.cell(row = passCount + 3, column = i)
        sheetValue.cell(row = passCount + 5, column = i).value = "=AVERAGE({}:{})".format(celltop.coordinate, cellbuttom.coordinate)
        #sheetValue['F17'] = "=AVERAGE(F5:F15)"
        print("=AVERAGE({}:{})".format(celltop.coordinate, cellbuttom.coordinate))

for row in sheet:
    for cell in row:
        cell.font = Font(name = "Arial", size = 11)

sheet['A1'].font = Font(name = "Arial", size = 14, color = 'FFFFFF', bold = True)
sheet['A2'].font = Font(name = "Arial", size = 12, color = 'FFFFFF', bold = True)
sheet['C2'].font = Font(name = "Arial", size = 11, color = 'FFFFFF', bold = True)
sheet['F2'].font = Font(name = "Arial", size = 11, bold = True)
sheet['A4'].font = Font(name = "Arial", size = 11, bold = True)
sheet['B4'].font = Font(name = "Arial", size = 11, bold = True)
sheet['C4'].font = Font(name = "Arial", size = 11, bold = True)
sheet['D4'].font = Font(name = "Arial", size = 11, bold = True)
sheet['E4'].font = Font(name = "Arial", size = 11, bold = True)

for row in sheetValue:
    for cell in row:
        cell.font = Font(name = "Arial", size = 11)

sheetValue['A1'].font = Font(name = "Arial", size = 14, color = 'FFFFFF', bold = True)
sheetValue['A2'].font = Font(name = "Arial", size = 12, color = 'FFFFFF', bold = True)
sheetValue['C2'].font = Font(name = "Arial", size = 11, color = 'FFFFFF', bold = True)
sheetValue['F2'].font = Font(name = "Arial", size = 11, bold = True)
sheetValue['A4'].font = Font(name = "Arial", size = 11, bold = True)
sheetValue['B4'].font = Font(name = "Arial", size = 11, bold = True)
sheetValue['C4'].font = Font(name = "Arial", size = 11, bold = True)
sheetValue['D4'].font = Font(name = "Arial", size = 11, bold = True)
sheetValue['E4'].font = Font(name = "Arial", size = 11, bold = True)

print("\n\n資料儲存中...\n")

#print(sheetValue.max_column)

# 儲存編輯好的 Excel
excel.save(unquote(urlparse(excelPath.get()).path))

print("儲存完成！\n")
time.sleep(2)

if wrongFormat != []:   pop_up()
