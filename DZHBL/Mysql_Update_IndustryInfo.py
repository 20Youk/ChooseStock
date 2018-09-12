# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用: 连接Mysql数据库，插入行业分类信息到对应表
import MySQLdb
import pandas as pd

# filePath = '../../file/SupplierInfo.xlsx'
filePath = '../../file/SupplierInfo180627.xlsx'
df = pd.read_excel(filePath, sheet_name=0)
df = df.where(df.notnull(), 'NULL')
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
# sql = '''UPDATE SupplierInfo set Industry = '%s' where SupplierCode = '%d' '''
# sql = '''UPDATE SupplierInfo set City = '%s' where SupplierCode = '%d' '''
# sql = '''UPDATE SupplierInfo set Industry='%s', City='%s', Contact1='%s', TelPhone1='%s', Mail1='%s',
#             Contact2='%s', TelPhone2='%s', Mail2='%s', Contact3='%s', TelPhone3='%s', Mail3='%s', Contact4='%s', TelPhone4='%s', Mail4='%s',
#             RegisteredCapital=%f, EmployeeNum=%d, Litigation=%d, RegisteredCurrency='%s' where SupplierCode=%d '''
sql = '''UPDATE SupplierInfo set Industry = '%s' where SupplierCode = '%d' '''
for i in range(0, len(allValues)):
    oldValue = allValues[i]
    # newValue = [oldValue[2], oldValue[3], oldValue[4], str(oldValue[5]) if oldValue[5] else oldValue[5], oldValue[6],
    #             oldValue[7], str(oldValue[8]) if oldValue[8] else oldValue[8], oldValue[9], oldValue[10],
    #             str(oldValue[11]) if oldValue[11] else oldValue[11], oldValue[12], oldValue[13],
    #             str(oldValue[14]) if oldValue[14] else oldValue[14], oldValue[15],
    #             float(oldValue[16]) if oldValue[16] else oldValue[16], int(oldValue[17]) if oldValue[17] else oldValue[17],
    #             int(oldValue[18]) if oldValue[18] else oldValue[18], oldValue[19], int(oldValue[0])]
    newValue = [oldValue[2], int(oldValue[0])]
    changedSql = sql % tuple(newValue)
    changedSql = changedSql.replace("'NULL'", "NULL")
    cursor.execute(changedSql)
# db.commit()
# data = cursor.fetchall()
# print 'Database version is : %s' % data
db.close()
print 'Done'

