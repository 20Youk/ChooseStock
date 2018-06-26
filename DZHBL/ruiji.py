# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:
from PIL import Image
import pytesseract
import os
import xlsxwriter

# filePath = '../../file/ruiji/RP0830029902.tif'
# pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe'
# x2 = 580
# y2 = 550
# w2 = 200
# h2 = 40
# im = Image.open(filePath)
# RR_date = im.crop((x2, y2, x2 + w2, y2 + h2))
# RR_date.save('../../file/ImageAn/RR_date.jpeg')
# new_date = pytesseract.image_to_string(RR_date).replace(' ', '')
# print u'识别送货单时间为:' + new_date

wb = xlsxwriter.Workbook('../../file/ImageAn_ruiji.xlsx')
ws = wb.add_worksheet('Sheet1')
field = [u'序号', u'收货日期', u'收货单单号', u'收货单金额']
ws.write_row(0, 0, field)
pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe'
srcPath = '../../file/ruiji/'
x1 = 320
y1 = 335
w1 = 200
h1 = 30
x2 = 580
y2 = 550
w2 = 200
h2 = 40
x3 = 585
y3 = 730
w3 = 200
h3 = 30
allList = []
count = 1
for item in os.listdir(srcPath):
    print u'第%d次识别....' % count
    filePath = srcPath + item
    im = Image.open(filePath)
    RR_num = im.crop((x1, y1, x1 + w1, y1 + h1))
    RR_text = pytesseract.image_to_string(RR_num).replace(' ', '')
    RR_num.save('../../file/ImageAn/RR_num.jpeg')
    print u'识别送货单号为:' + RR_text

    RR_date = im.crop((x2, y2, x2 + w2, y2 + h2))
    RR_date.save('../../file/ImageAn/RR_date.jpeg')
    new_date = pytesseract.image_to_string(RR_date).replace(' ', '')
    print u'识别送货单时间为:' + new_date

    RR_totle = im.crop((x3, y3, x3 + w3, y3 + h3))
    RR_totle.save('../../file/ImageAn/RR_totle.jpeg')
    totle_text = pytesseract.image_to_string(RR_totle).replace(' ', '')
    print u'识别送货单金额为:' + totle_text
    ws.write_row(count, 0, [count, new_date, RR_text, totle_text])
    count += 1
wb.close()
print 'Done!!!'
