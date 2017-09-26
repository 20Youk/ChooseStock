# -*- coding:utf8 -*-
import pymssql
import xlwt
import datetime

server = '192.168.8.200'
user = 'sa'
password = 'wind123@pa'
database = 'PortfolioData'
# 增加判断，筛选出半年创新高的股票
sql1 = '''with h1 as (select a.SecCode2,%s as TStarDate, %s as TEndDate, (select max(AdjClose) from SecPrice where SecCode = a.SecCode and FDate >= DATEADD(MONTH, -1, %s) and FDate < %s) / (select max(AdjClose) from SecPrice where SecCode = a.SecCode and FDate >= DATEADD(MONTH, -6, %s)  and FDate < %s) PricePro from SecInfo a  where SecCode2 in(%s) and SecType = 'A' )
        select SecCode2, TStarDate, TEndDate  from h1 where PricePro >= 1'''
sql2 = '''with h1 as (select [DAY], (ROW_NUMBER() over (order by [Day]) - 1) / 5 r  from TradeDay a where [Day] > = %s and [Day] < %s)
		select MIN([DAY]) TStartDate, MAX([DAY]) TEndDate from h1 group by r '''
conn = pymssql.connect(server, user, password, database, charset='utf8')
cursor = conn.cursor()
# 执行选股收益率统计sql
today = datetime.datetime.now().strftime('%Y%m%d')
daylist = ('-04-30', '-08-30')
sqldata1 = []
sqldata2 = []
startyear = 2006
k = 0
while k <= 10:
    year = startyear + k
    for day in daylist:
        startday = str(year) + day
        cursor.execute('exec [GetNewChooseCode] @fDate = %s', startday)
        data = cursor.fetchall()
        cursor.execute(sql2, (data[0][0], data[0][1]))
        dayData = cursor.fetchall()
        codeList = []
        for j in range(0, len(data)):
            codeList.append(data[j][2])
        codeString = ('%s,' * len(codeList))[:-1]
        s = '%s'
        sql11 = sql1 % (s, s, s, s, s, s, codeString)
        codeList1 = tuple(codeList)
        for m in range(0, len(dayData)):
            codeList.insert(0, dayData[m][0])
            codeList.insert(1, dayData[m][1])
            codeList.insert(2, dayData[m][0])
            codeList.insert(3, dayData[m][0])
            codeList.insert(4, dayData[m][0])
            codeList.insert(5, dayData[m][0])
            cursor.execute(sql11, tuple(codeList))
            data1 = cursor.fetchall()
            if data1:
                sqldata1.extend(data1)
            codeList = list(codeList1)
    k += 1
    sqldata1.sort(key=lambda x: x[1])
# 将sql执行结果插入excle
workbook = xlwt.Workbook(encoding='utf8')
sheet1 = workbook.add_sheet('sheet1', cell_overwrite_ok=True)
sheet2 = workbook.add_sheet('sheet2', cell_overwrite_ok=True)
field1 = ['证券代码', '周期开始日', '周期截止日']
for i in range(0, len(field1)):
    sheet1.write(0, i, field1[i])
for row in range(1, len(sqldata1) + 1):
    for col in range(0, len(field1)):
        sheet1.write(row, col, sqldata1[row - 1][col])
# for i in range(0, len(field2)):
#     sheet2.write(0, i, field1[i][0])
# for row in range(1, len(sqldata2) + 1):
#     for col in range(0, len(field2)):
#         sheet2.write(row, col, sqldata2[row - 1][col])
workbook.save(r'C:\Users\Administrator\Desktop\QualityCode%s.xls' % today)
conn.close()
print 'Done!!!'
