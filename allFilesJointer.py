from re import template
import xlrd
from openpyxl import Workbook
from xlrd import *


workbook = Workbook()
ws = workbook.active


def allFilesJointer(filenames):

    # print(filenames[0])
    # haha = xlrd.count_records(filenames[0])
    # loc = filenames[0]
    # book = xlrd.open_workbook(loc)
    # # print(xlrd.sheet.Sheet(book).row_len)
    # sheet = book.sheet_by_index(0)
    # print(sheet.nrows)
    # print(sheet.ncols)

    for files in range(len(filenames)):
        tempList = []
        loc = (filenames[files])
        print(loc)
        try:
            wb = xlrd.open_workbook(loc)
        except:
            pass

        sheet = wb.sheet_by_index(0)
        total_rows = sheet.nrows
        total_cols = sheet.ncols
        if total_rows < 1 or total_cols < 1:
            break

        ws.append([str(loc)])  # This is the file name above the sheet
        ws.append([''])  # This is a gap between name and the data
        #########
        for row in range(total_rows):
            tempList.append([])
            for col in range(total_cols):
                try:
                    tempData = str(sheet.cell_value(row, col))
                    tempList[row].append(str(tempData).strip())
                except:
                    pass

        for excelData in tempList:
            ws.append(excelData)

        ws.append([''])

        workbook.save('jointjoint.xlsx')
        tempList = []
        workbook.close()

    # for row in range()
