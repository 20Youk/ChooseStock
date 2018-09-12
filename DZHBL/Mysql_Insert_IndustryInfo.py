# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用: 连接Mysql数据库，插入行业分类信息到对应表
import MySQLdb
import xlrd
import pandas as pd

# filePath = '../../file/SupplierInfo.xlsx'
filePath = '../../file/IndustryInfo.xlsx'
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
sql = '''insert INTO IndustryInfo(IndustryCode0, IndustryName0, IndustryCode1, IndustryName1)
    values(%d, '%s', '%d', '%s') '''


for i in range(0, len(allValues)):
    oldValue = allValues[i]
    newValue = [int(oldValue[0]), oldValue[1], int(oldValue[2]), oldValue[3]]
    changedSql = sql % tuple(newValue)
    cursor.execute(changedSql)
# db.commit()
# data = cursor.fetchall()
# print 'Database version is : %s' % data
db.close()
print 'Done'

