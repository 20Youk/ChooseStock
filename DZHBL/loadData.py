# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用: 连接Mysql数据库，导出数据到Excel
import MySQLdb
import xlsxwriter

filePath = '../../file/SupplierInfo20180716.xlsx'
host = '***'
user = '***'
ps = '***'
db = '***'
db = MySQLdb.connect(host, user, ps, db, charset='utf8')
db.autocommit(on=True)
cursor = db.cursor()
sql = '''select SupplierName, Industry from SupplierInfo
'''


cursor.execute(sql)
data = cursor.fetchall()
# print 'Database version is : %s' % data
field = cursor.description
db.close()
wb = xlsxwriter.Workbook(filePath)
ws = wb.add_worksheet('Sheet1')
for i in range(0, len(field)):
    ws.write(0, i, field[i][0].decode('utf8'))
for i in range(0, len(data)):
    for j in range(0, len(field)):
        ws.write(i + 1, j, data[i][j])
wb.close()
print 'Done'

