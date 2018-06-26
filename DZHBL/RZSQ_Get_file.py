# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:融资申请附件批量下载
import xlrd
import os
import urllib

# Excel文件路径
srcPath = u'../../file/融资申请票据.xlsx'
# 设置下载后存放的存储路径
path = '../../file'

# 读取Excel
rb = xlrd.open_workbook(srcPath)
rs = rb.sheet_by_index(0)
code_list = rs.col_values(1, start_rowx=1)
category_list = rs.col_values(2, start_rowx=1)
supplier_list = rs.col_values(3, start_rowx=1)
filename_list = rs.col_values(4, start_rowx=1)
link_list = rs.col_values(5, start_rowx=1)
lastSupplier = ''

# 定义下载函数downLoadPicFromURL（本地文件夹，网页URL）
for i in range(0, len(code_list)):
    dirPath = os.path.join(path, supplier_list[i].strip())
    if supplier_list[i] != lastSupplier and not os.path.exists(dirPath):
        os.mkdir(dirPath)
        os.makedirs(dirPath + '/ApplyExcel')
        os.makedirs(dirPath + '/PO')
        os.makedirs(dirPath + '/Receipt')
        os.makedirs(dirPath + '/Invoice')
    if category_list[i] == 1:
        dest_dir = dirPath + '/ApplyExcel/' + filename_list[i]
    elif category_list[i] == 2:
        dest_dir = dirPath + '/PO/' + filename_list[i]
    elif category_list[i] == 3:
        dest_dir = dirPath + '/Receipt/' + filename_list[i]
    else:
        dest_dir = dirPath + '/Invoice/' + filename_list[i]
    url = link_list[i]
    urllib.urlretrieve(url, dest_dir)
    lastSupplier = supplier_list[i]
print 'Done!!!'

