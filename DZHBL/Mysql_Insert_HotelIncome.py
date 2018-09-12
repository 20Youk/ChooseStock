# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用: 连接Mysql数据库，插入酒店收入信息到对应表
import MySQLdb
import pandas as pd

# filePath = '../../file/SupplierInfo.xlsx'
filePath = '../../file/data_result.xlsx'
df = pd.read_excel(filePath, sheet_name=0)
df = df.fillna('NULL')
allValues = df.values
host = '***'
user = '***'
ps = '***'
db = '***'
db = MySQLdb.connect(host, user, ps, db, charset='utf8')
db.autocommit(on=True)
cursor = db.cursor()
# sql = '''insert INTO SupplierInfo(SupplierCode, SupplierName, Industry, City, Contact1, TelPhone1, Mail1,
# Contact2, TelPhone2, Mail2, Contact3, TelPhone3, Mail3, Contact4, TelPhone4, Mail4)
#     values(%d, '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s') '''
sql = '''INSERT into HotelIncome(Hotel, RMonth, OccupancyRate, BanquetIncome)
    values('%s', '%s', '%f', '%f') '''


for i in range(0, len(allValues)):
    oldValue = allValues[i]
    newValue = [oldValue[0], oldValue[1], float(oldValue[2]), float(oldValue[3])]
    changedSql = sql % tuple(newValue)
    cursor.execute(changedSql)
# db.commit()
# data = cursor.fetchall()
# print 'Database version is : %s' % data
db.close()
print 'Done'

