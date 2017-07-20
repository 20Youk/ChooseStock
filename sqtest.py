# -*- coding:utf8 -*-
import pymssql
import xlwt
import datetime
server = '192.168.8.200'
user = 'sa'
password = 'wind123@pa'
database = 'PortfolioData'
conn = pymssql.connect(server, user, password, database, charset='utf8')
sql2 = '''select a.SecCode, AVG(b.OperatingNetFlow) / STDEV(b.OperatingNetFlow - a.NetIncome) Adj from IncomeInfo a , CashFlow b where a.SecCode = b.SecCode and a.ReportingDate = b.ReportingDate and a.ReportingDate <= '2006' and a.ReportingDate >= DATEADD(YEAR, -5, '2006') and RIGHT('0'+ltrim(MONTH(a.ReportingDate)),2) = '12' group by a.SecCode having STDEV(b.OperatingNetFlow - a.NetIncome) <> 0 '''
cursor = conn.cursor()
date1 = '2007-05-01'
date2 = '2017-01-06'
today = datetime.datetime.now().strftime('%Y%m%d')
cursor.execute('exec GetNewChooseCode @fDate = %s', '2006-04-30')
# cursor.execute('exec [GetPriceChangeNum] @startDate = %s, @endDate = %s',('2017-07-17','2017-07-18'))
# cursor.execute(sql)
# data1 = cursor.fetchone()
sqldata = cursor.fetchall()
field1 = cursor.description
conn.close()
print sqldata, '\n', len(sqldata)
# 将sql执行结果插入excle
workbook = xlwt.Workbook(encoding='utf8')
sheet = workbook.add_sheet('sheet1', cell_overwrite_ok=True)
for i in range(0, len(field1)):
    sheet.write(0, i, field1[i][0])
for row in range(1, len(sqldata) + 1):
    for col in range(0, len(field1)):
        sheet.write(row, col, sqldata[row - 1][col])
workbook.save(r'C:\Users\Administrator\Desktop\sqltest%s.xls' % today)