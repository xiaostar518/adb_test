#! /usr/bin/env python
# coding=utf-8
import os
import xlrd
import xlwt
from xlutils.copy import copy
import time
from xlwt import Style

currentTime = time.strftime('%Y_%m_%d_%H_%M_%S', time.localtime(time.time()))
print currentTime

# 文件名
excelName = "demo_" + currentTime + ".xls"
# sheet名
sheet_name = 'demo'
# 表头
row0 = [u'ID', u'Time', u'Content']  # flies = os.listdir("app");

isGoOn = True

style = xlwt.easyxf('font:height 240, color-index red, bold on;align: wrap on, vert centre, horiz center')


def get_current_time():
    return time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))


def set_style(name, height, bold=False):
    style = xlwt.XFStyle()  # 初始化样式

    font = xlwt.Font()  # 为样式创建字体
    font.name = name
    font.bold = bold
    font.color_index = 4
    font.height = height

    style.font = font
    return style


def excel_init():
    try:
        # 打开文件
        readWorkbook = xlrd.open_workbook(excelName)
    except:
        print u"demo.xls 文件不存在"
        isHaveExcel = False
    else:
        print u"demo.xls 文件存在"
        isHaveExcel = True

    if isHaveExcel is False:
        create_excel(excelName, sheet_name, row0)
        return

    try:
        # 获取所有sheet
        print readWorkbook.sheet_names()  # [u'sheet1', u'sheet2']
        demo_sheet = readWorkbook.sheet_by_name(sheet_name)
    except:
        print u"sheet demo 不存在"
        isHaveSheet = False
    else:
        print u"sheet demo 存在"
        isHaveSheet = True

    if isHaveSheet is False:
        create_excel(excelName, sheet_name, row0)
        return

    # 获取所有row
    print(u"一共有：%d行" % demo_sheet.nrows)
    if demo_sheet.nrows == 0:
        create_excel(excelName, sheet_name, row0)
    else:
        print demo_sheet.row_values(0)
        arr = []
        for i in demo_sheet.row_values(0):
            arr.append(i)
            print i
        if row0 == arr:
            print u"索引行标准"
        else:
            print u"索引行不标准"
            create_excel(excelName, sheet_name, row0)


def create_excel(excelName, sheetName, row0):
    # 创建工作簿
    writeWorkbook = xlwt.Workbook(encoding='utf-8')
    print u"创建文件"

    # 创建sheet
    demo_sheet = writeWorkbook.add_sheet(sheetName)
    print u"创建sheet demo"

    # 生成表头
    for i in range(len(row0)):
        demo_sheet.write(0, i, row0[i], set_style('Times New Roman', 220, True))

    # 保存文件
    writeWorkbook.save(excelName)
    print u"创建文件成功"


def writeExcel(row, col, str, style=Style.default_style):
    rb = xlrd.open_workbook(excelName, formatting_info=True)
    wb = copy(rb)
    ws = wb.get_sheet(0)
    ws.write(row, col, str, style)
    wb.save(excelName)


def write_loop():
    i = 0
    j = 1
    while True:
        print "循环次数" + bytes(i)
        i += 1
        # text = "adb devices"
        text = "adb logcat -c && adb logcat -d -v time"
        # text = "adb logcat -d -v time"
        print text
        content = os.popen(text)
        out = content.read()
        # content.close()
        # print out
        print out.split('\n')
        data = out.split('\n')
        print len(data)
        for line in data:
            writeExcel(j, 0, j, set_style('Times New Roman', 220))
            writeExcel(j, 1, get_current_time(), set_style('Times New Roman', 220))
            writeExcel(j, 2, line, set_style('Times New Roman', 220))
            j += 1
            print 'writeExcel:' + bytes(j)


excel_init()

write_loop()
