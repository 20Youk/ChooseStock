# -*- coding:utf8 -*-
import xlrd
from xlutils.copy import copy
try:
    wb = xlrd.open_workbook('C:\Users\Administrator\Desktop\lin_Quality_HS300_20170802.xlsx')
    sheet = wb.sheet_by_index(1)
    returnList = sheet.col(1, start_rowx=1)
    proportion = []
    for i in range(0, len(returnList), 5):
        oneList = [cmp(j.value, 0) for j in returnList[i: i + 5]]
        upCount = oneList.count(1)
        downCount = oneList.count(-1)
        keepCount = oneList.count(0)
        if i + 5 >= len(returnList):
            proportion.append([sheet.cell(len(returnList) - 1, 0).value, float(upCount) / float(upCount + downCount + keepCount)])
        else:
            proportion.append([str(sheet.cell(i + 5, 0).value), float(upCount) / float(upCount + downCount + keepCount)])
    result = sum([p[1] for p in proportion]) / len(proportion)
    wb1 = copy(wb)
    sheet1 = wb1.get_sheet(1)
    sheet1.write(4, 12, u'周胜率')
    sheet1.write(4, 13, result)
    sheet2 = wb1.add_sheet('sheet3')
    field = [u'交易日', u'周胜率']
    for k in range(0, len(field)):
        sheet2.write(0, k, field[k])
    for l in range(0, len(proportion)):
        for k in range(0, len(field)):
            sheet2.write(l + 1, k, proportion[l][k])
    wb1.save(r'C:\Users\Administrator\Desktop\lin_Quality_HS300_20170802.xlsx')
except IOError, e:
    print '执行失败\n', e
else:
    print 'Done!!!'

















































































































































