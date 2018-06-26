# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用: 连接Mysql数据库，导出数据到Excel
import MySQLdb
import xlsxwriter

filePath = '../../file/com_180619.xlsx'
host = '39.108.218.254'
user = 'root'
ps = 'Gccf,1234'
db = 'GCCFSI'
db = MySQLdb.connect(host, user, ps, db, charset='utf8')
db.autocommit(on=True)
cursor = db.cursor()
sql = '''select * from SupplierCommunication where Company in (
'SSWM照明广告',
'广州创首贸易有限公司',
'深圳欧尚艺术设计有限公司',
'深圳市明捷电器',
'深圳市森源家具',
'深圳市南方联合酒店设备有限公司',
'苏州联诺信息技术有限公司',
'深圳迪烨大马贸易有限公司',
'深圳市福馨柯贸易有限公司',
'深圳市佳厨厨具有限公司',
'深圳市美登家具有限公司',
'深圳市旺廚廚房設備有限公司',
'北京西科盛世通酒店会展设备制造有限公司',
'惠州市尊宝智能控制股份有限公司',
'深圳市勤峻实业有限公司',
'深圳市圣象木业有限公司',
'深圳市恒安兴酒店用品集团股份有限公司'
)
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

