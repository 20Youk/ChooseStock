# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用: 连接Mysql数据库，插入酒店收入信息到对应表
import MySQLdb
import xlrd
import pandas as pd

# filePath = '../../file/SupplierInfo.xlsx'
filePath = '../../file/SupplierCommucation.xls'
df = pd.read_excel(filePath, sheet_name=0)
df = df.fillna('NULL')
allValues = df.values
host = '39.108.218.254'
user = 'root'
ps = 'Gccf,1234'
db = 'GCCFSI'
db = MySQLdb.connect(host, user, ps, db, charset='utf8')
db.autocommit(on=True)
cursor = db.cursor()
# sql = '''insert INTO SupplierInfo(SupplierCode, SupplierName, Industry, City, Contact1, TelPhone1, Mail1,
# Contact2, TelPhone2, Mail2, Contact3, TelPhone3, Mail3, Contact4, TelPhone4, Mail4)
#     values(%d, '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s') '''
sql = '''INSERT into SupplierCommunication(WorkNum, Informant, Department, InformantTime, VisitObject,
        Company, VisitType, Content, Remark) VALUES('%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s')'''


for i in range(0, len(allValues)):
    oldValue = allValues[i]
    newValue = [str(oldValue[0]), oldValue[1], oldValue[2], str(oldValue[3]), oldValue[4], oldValue[5], oldValue[6], oldValue[7], oldValue[8]]
    changedSql = sql % tuple(newValue)
    changedSql = changedSql.replace("'NULL'", "NULL")
    cursor.execute(changedSql)
# db.commit()
# data = cursor.fetchall()
# print 'Database version is : %s' % data
db.close()
print 'Done'

