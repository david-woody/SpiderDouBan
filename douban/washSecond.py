# encoding: utf-8
import os

import xlwt
from xlrd import open_workbook
from xlutils.copy import copy

tittle_style = xlwt.easyxf(
    'font: height 240, name Arial Black, colour_index black, bold on; align: wrap on, vert centre, horiz center;')
normal_style = xlwt.easyxf(
    'font: height 240, name Arial, colour_index black, bold off; align: wrap on, vert centre, horiz center;')

rootdir = os.path.join(os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir)), 'result\\')
flagMap = {}
bookMap = {}
for parent, dirnames, filenames in os.walk(rootdir):  # 三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
    for fileName in filenames:
        filepath = os.path.join(os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir)),
                                'result\\' + fileName)
        # print "FileName:",filepath
        inWb = open_workbook(filepath, formatting_info=True)
        rSheet = inWb.sheet_by_index(0)
        maxRowIndex = rSheet.nrows
        # print "最大行数", maxRowIndex
        for ii in range(1, maxRowIndex):
            print "FileName:", filepath
            flagList = list()
            print rSheet.cell_value(ii, 4)
            if rSheet.cell_value(ii,4).__contains__(";"):
                booksList = rSheet.cell_value(ii, 4).split(";")
            else:
                booksList = rSheet.cell_value(ii, 4).split(",")
            print  len(booksList)
            for book in booksList:
                book = book.strip()
                if bookMap.has_key(book):
                    if bookMap[book].__contains__(rSheet.cell_value(ii, 2)):
                        continue
                    else:
                        bookMap[book].append(rSheet.cell_value(ii, 2))
                else:
                    flagList.append(rSheet.cell_value(ii, 2))
                    bookMap[book] = flagList
outWb = open_workbook("allTags.xls", formatting_info=True)
newbook = copy(outWb)
try:
    oSheet = newbook.get_sheet(1)
except Exception as e:
    oSheet = newbook.add_sheet('书名标签结果', cell_overwrite_ok=True)
for A in range(100):
    oSheet.col(A).width = 4800
oSheet.write(0, 0, '书名', tittle_style)
oSheet.write(0, 1, '标签集', tittle_style)
index = 1
for book in bookMap:
    print book, len(bookMap[book])
    oSheet.write(index, 0, book, normal_style)
    newFlagList = list(set(bookMap[book]))
    newFlagList.sort()
    oSheet.write(index, 1, ",".join(newFlagList), normal_style)
    index += 1
newbook.save("allTags.xls")
# if bookName == None and len(bookName) == 0:
#     continue
