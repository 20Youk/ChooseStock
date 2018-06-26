# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:更改文件名称
import os
import xlrd

filePath = '../../file/1'
# rb = xlrd.open_workbook('../../file/ruiji.xlsx')
# rs = rb.sheet_by_index(0)
# numList = rs.col_values(2, start_rowx=2)
fileList = os.listdir(filePath)
count = 0
for item in fileList:
    allFile = filePath + item
    os.rename(allFile, filePath + u'融资明细' + item)
    count += 1
print 'Done!!!'
