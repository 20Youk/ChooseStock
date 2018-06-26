# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:导出酒店返点数据
import xlsxwriter
import pymssql
import datetime
import os
import sys

host, user, password, database = '172.18.4.47', 'dbreader', 'dzhbl1234', 'Baoli'
sqlList = ['[Hotel_Rebate]']
today = datetime.date.today()
filePath = '/home/vftpuser/public/酒店返点数据/%s' % today
if not os.path.exists(filePath):
    os.mkdir(filePath)
excel = filePath + '/酒店返点基础数据%s.xlsx' % today
# 初始化日期
lastDay = today - datetime.timedelta(days=+today.day)
monthFirstDay = lastDay - datetime.timedelta(days=lastDay.day - 1)
if len(sys.argv) == 2:
    monthFirstDay = datetime.datetime.strptime(sys.argv[1], '%Y-%m-%d')
if len(sys.argv) == 3:
    monthFirstDay = datetime.datetime.strptime(sys.argv[1], '%Y-%m-%d')
    lastDay = datetime.datetime.strptime(sys.argv[2], '%Y-%m-%d')
print 'data-time is :' + monthFirstDay.strftime('%Y-%m-%d') + ' ~ ' + lastDay.strftime('%Y-%m-%d')
timeList = [[monthFirstDay.strftime('%Y-%m-%d'), lastDay.strftime('%Y-%m-%d')]]
# 连接数据库
conn = pymssql.connect(host, user, password, database, charset='utf8')
cursor = conn.cursor()
# 执行语句
sqlDataList = []
fieldList = []
for i in range(0, len(sqlList)):
    sql = 'exec ' + sqlList[i] + " @SDate = '%s', @EDate = '%s'" % (timeList[i][0], timeList[i][1])
    cursor.execute(sql)
    data = cursor.fetchall()
    field = cursor.description
    sqlDataList.append(data)
    fieldList.append(field)
conn.close()
wb = xlsxwriter.Workbook(excel)
ws1 = wb.add_worksheet('Sheet1')
grey = wb.add_format({'border': 1, 'align': 'vcenter', 'bg_color': '#696969', 'font_size': 9, 'font_color': 'black'})
style = wb.add_format({'border': 1, 'align': 'vcenter', 'font_size': 9, 'font_color': 'black'})
ws1.merge_range(0, 0, 0, 11, u'返点计算', grey)
title = u'制作人：ice   日期：%s   返点计算期间：%s-%s' % (today, monthFirstDay, lastDay)
print title
ws1.merge_range('A2:L2', title, grey)
ws1.write_row(2, 0, [item[0] for item in fieldList[0]], style)
sum_rebate = 0
for i in range(0, len(sqlDataList[0])):
    sum_rebate += sqlDataList[0][i][10]
    for j in range(0, len(fieldList[0])):
        ws1.write(i + 3, j, sqlDataList[0][i][j], style)
count = len(sqlDataList[0]) + 3
print sum_rebate
ws1.write(count, 9, u'合计', grey)
ws1.write(count, 10, sum_rebate, grey)
ws1.set_column('A:A', 13)
ws1.set_column('B:B', 24)
ws1.set_column('C:C', 35)
wb.close()
print 'DONE!!!'
