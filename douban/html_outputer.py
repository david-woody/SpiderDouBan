# encoding: utf-8
import os

import xlwt
from xlrd import open_workbook
from xlutils.copy import copy
from xlwt.Workbook import Workbook

## Style variable Begin
tittle_style = xlwt.easyxf(
    'font: height 240, name Arial Black, colour_index black, bold on; align: wrap on, vert centre, horiz center;')
subtittle_left_style = xlwt.easyxf(
    'font: height 240, name Arial, colour_index brown, bold on, italic on; align: wrap on, vert centre, horiz left;'      "borders: top double, bottom double, left double;")
subtittle_right_style = xlwt.easyxf(
    'font: height 240, name Arial, colour_index brown, bold on, italic on; align: wrap on, vert centre, horiz left;'      "borders: top double, bottom double, right double;")
subtittle_top_and_bottom_style = xlwt.easyxf(
    'font: height 240, name Arial, colour_index black, bold off, italic on; align: wrap on, vert centre, horiz left;'      "borders: top double, bottom double;")
blank_style = xlwt.easyxf(
    'font: height 650, name Arial, colour_index brown, bold off; align: wrap on, vert centre, horiz left;'      "borders: top double, bottom double, left double, right double;")
normal_style = xlwt.easyxf(
    'font: height 240, name Arial, colour_index black, bold off; align: wrap on, vert centre, horiz center;')


## Style variable End
class HtmlOutputer(object):
    def __init__(self):
        if os.path.exists("result.xls") == False:
            # 判断文件是否存在，存在则跳过，不存在则新建
            self.book = Workbook(encoding='utf-8', style_compression=0)
            self.sheet = self.book.add_sheet('豆瓣标签结果', cell_overwrite_ok=True)
            ## Column Width Determine Begin
            for A in range(100):
                self.sheet.col(A).width = 4800
            ## Column Width Determine Begin
            self.sheet.write(0, 0, '编号', tittle_style)
            self.sheet.write(0, 1, '用户', tittle_style)
            self.sheet.write(0, 2, '标签', tittle_style)
            self.sheet.write(0, 3, '次数', tittle_style)
            self.sheet.write(0, 4, '标注过', tittle_style)
            self.rowx = 1
            self.index = 1
            # sheet.write_merge(0, 5, 0, 1, "哈哈哈", normal_style)
        else:
            rb = open_workbook('result.xls', formatting_info=True)  # 注意这里的workbook首字母是小写
            r_sheet = rb.sheet_by_index(0)
            self.rowx = r_sheet.nrows
            firstColsValue = r_sheet.col_values(0)[::-1]
            for index in firstColsValue:
                if str(index).__eq__("") == False:
                    self.index = int(index)
                    break
            self.index = self.index + 1
            # 管道作用
            self.book = copy(rb)
            # 通过get_sheet()获取的sheet有write()方法
            self.sheet = self.book.get_sheet(0)
            self.sheet.write(0, 7, 'lalal', tittle_style)
        return

    def writeData(self, rowx, colx, data):
        self.sheet.write(rowx, colx, '标注过', tittle_style)

    def writeDataDefault(self, username, allTags, allBooks):
        self.sheet.write(self.rowx, 0, self.index, normal_style)
        self.sheet.write(self.rowx, 1, username, normal_style)
        for tag in allTags:
            self.sheet.write(self.rowx, 2, tag, normal_style)
            self.sheet.write(self.rowx, 3, allBooks[tag].__len__(), normal_style)
            self.sheet.write(self.rowx, 4, ",".join(allBooks[tag]), normal_style)
            self.rowx = self.rowx + 1


    def save(self, name):
        self.book.save(name)
