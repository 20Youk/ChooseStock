# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用: 连接Mysql数据库，导出数据到Excel
import MySQLdb
import xlsxwriter

filePath = '../../file/UnvoiceInfo20190221.xlsx'
host = '39.108.218.254'
user = 'root'
ps = 'Gccf,1234'
db = 'GCCFSI'
db = MySQLdb.connect(host, user, ps, db, charset='utf8')
db.autocommit(on=True)
cursor = db.cursor()
sql = '''select GoodName 品名,
Unit 单位,
  Amount 数量,
  UnitPrice 单价
from InvoiceGoods
where InvoiceNum is not NULL
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

