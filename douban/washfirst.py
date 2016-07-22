# encoding: utf-8
import os

import xlwt
from xlrd import open_workbook
from xlwt import Workbook

tittle_style = xlwt.easyxf(
    'font: height 240, name Arial Black, colour_index black, bold on; align: wrap on, vert centre, horiz center;')
normal_style = xlwt.easyxf(
    'font: height 240, name Arial, colour_index black, bold off; align: wrap on, vert centre, horiz center;')

rootdir = os.path.join(os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir)), 'result\\')
flagMap = {}
for parent, dirnames, filenames in os.walk(rootdir):  # 三个参数：分别返回1.父目录 2.所有文件夹名字（不含路径） 3.所有文件名字
    for fileName in filenames:
        filepath = os.path.join(os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir)),
                                'result\\' + fileName)
        inWb = open_workbook(filepath, formatting_info=True)
        rSheet = inWb.sheet_by_index(0)
        maxRowIndex = rSheet.nrows
        # print "最大行数", maxRowIndex
        for ii in range(1, maxRowIndex):
            if flagMap.has_key(rSheet.cell_value(ii, 2)):
                flagMap[rSheet.cell_value(ii, 2)] = flagMap[rSheet.cell_value(ii, 2)] + int(rSheet.cell_value(ii, 3))
            else:
                flagMap[rSheet.cell_value(ii, 2)] = int(rSheet.cell_value(ii, 3))
flagList = flagMap.iteritems()
flagList = sorted(flagList, key=lambda d: d[1], reverse=True)

book = Workbook(encoding='utf-8', style_compression=0)
sheet = book.add_sheet('豆瓣所有用户标签结果', cell_overwrite_ok=True)
for A in range(100):
    sheet.col(A).width = 4800
sheet.write(0, 0, '标签', tittle_style)
sheet.write(0, 1, '次数', tittle_style)
index = 1
for flag in flagList:
    sheet.write(index, 0, flag[0], normal_style)
    sheet.write(index, 1, flag[1], normal_style)
    index += 1
book.save("allTags.xls")