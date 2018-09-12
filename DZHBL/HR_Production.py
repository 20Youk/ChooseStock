# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:导出HR绩效报表
import xlsxwriter
import pymssql
from dateutil.relativedelta import relativedelta
import datetime
import os
import sys

host, user, password, database = '***', '***', '***', '***'
sqlList = ['[HR_Business_Product]', '[HR_Opration_Product]', '[HR_Group_Product]', '[HR_Group_Product]', '[HR_Group_Product]']
today = datetime.date.today()
filePath = '/home/vftpuser/public/HR绩效报表/%s' % today
if not os.path.exists(filePath):
    os.mkdir(filePath)
excel = filePath + '/HR绩效报表%s.xlsx' % today
# 初始化日期
lastDay = today - datetime.timedelta(days=+1)
monthFirstDay = lastDay - datetime.timedelta(days=lastDay.day - 1)
if len(sys.argv) == 2:
    monthFirstDay = datetime.datetime.strptime(sys.argv[1], '%Y-%m-%d')
if len(sys.argv) == 3:
    monthFirstDay = datetime.datetime.strptime(sys.argv[1], '%Y-%m-%d')
    lastDay = datetime.datetime.strptime(sys.argv[2], '%Y-%m-%d')
print 'data-time is :' + monthFirstDay.strftime('%Y-%m-%d') + ' ~ ' + lastDay.strftime('%Y-%m-%d')
lastMonthToday = lastDay - relativedelta(months=+1)
lastMonthFirstDay = monthFirstDay - relativedelta(months=+1)
lastYearToday = lastDay - relativedelta(years=+1)
lastYearFirstDay = monthFirstDay - relativedelta(years=+1)
timeList = [[monthFirstDay.strftime('%Y-%m-%d'), lastDay.strftime('%Y-%m-%d')],
            [monthFirstDay.strftime('%Y-%m-%d'), lastDay.strftime('%Y-%m-%d')],
            [monthFirstDay.strftime('%Y-%m-%d'), lastDay.strftime('%Y-%m-%d')],
            [lastMonthFirstDay.strftime('%Y-%m-%d'), lastMonthToday.strftime('%Y-%m-%d')],
            [lastYearFirstDay.strftime('%Y-%m-%d'), lastYearToday.strftime('%Y-%m-%d')]]
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
ws1.write_row(0, 0, [item[0] for item in fieldList[0]])
for i in range(0, len(sqlDataList[0])):
    for j in range(0, len(fieldList[0])):
        ws1.write(i + 1, j, sqlDataList[0][i][j])
ws2 = wb.add_worksheet('Sheet2')
ws2.write_row(0, 0, [item[0] for item in fieldList[1]])
for i in range(0, len(sqlDataList[1])):
    for j in range(0, len(fieldList[1])):
        ws2.write(i + 1, j, sqlDataList[1][i][j])
ws3 = wb.add_worksheet('Sheet3')
ws3.write_row(0, 0, [item[0] for item in fieldList[2]])
for i in range(0, len(sqlDataList[2])):
    ws3.write_row(i + 1, 0, sqlDataList[2][i])
    ws3.write_row(i + 1, 4, sqlDataList[3][i])
    ws3.write_row(i + 1, 8, sqlDataList[4][i])
wb.close()
print 'Done!!!'
