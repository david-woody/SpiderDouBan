# encoding: utf-8
import os

import xlwt
from xlrd import open_workbook

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
        inWb = open_workbook(filepath, formatting_info=True)
        rSheet = inWb.sheet_by_index(0)
        maxRowIndex = rSheet.nrows
        # print "最大行数", maxRowIndex
        for ii in range(1, maxRowIndex):
            flagList = list()
            booksList = rSheet.cell_value(ii, 4).split(",")
            for book in booksList:
                if ii==1:
                 # print book
                 if bookMap.has_key(book):
                    bookMap[book].append(rSheet.cell_value(ii, 2))
                 else:
                    flagList.append(rSheet.cell_value(ii, 2))
                    bookMap[book] = flagList
for book in bookMap:
    print book,len(bookMap[book])
                # if bookName == None and len(bookName) == 0:
                #     continue
