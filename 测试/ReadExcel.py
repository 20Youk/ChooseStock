# -*- coding:utf8 -*-
# import csv
#
# csvFile = csv.reader(open('..\..\excel\chicang.csv', 'rb'))
# i = 0
# dateList = []
# codeList = []
# mvList = []
# buyList = []
# netList = []
# for item in csvFile:
#     i += 1
#     if i > 1:
#         dateList.append(item[1])
#         codeList.append(item[6])
#         mvList.append(item[8])
#         buyList.append(item[10])
#         netList.append(item[12])
import xlrd
wb = xlrd.open_workbook('..\..\excel\chicang.xlsx')
sheet = wb.sheet_by_index(0)
dateList = sheet.col_values(1, start_rowx=1)
codeList = sheet.col_values(6, start_rowx=1)
mvList = sheet.col_values(8, start_rowx=1)
buyList = sheet.col_values(10, start_rowx=1)
netList = sheet.col_values(12, start_rowx=1)
print 'Done!'
