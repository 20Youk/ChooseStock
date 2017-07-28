# -*- coding:utf8 -*-
import pymssql
import xlwt
import datetime

server = '192.168.8.200'
user = 'sa'
password = 'wind123@pa'
database = 'PortfolioData'
# 增加判断，筛选出半年创新高的股票
sql1 = '''declare @tStartDate date
        declare @tEndDate date
        declare @topNum integer
        set @topNum = %d
        set @tStartDate = %s
        set @tEndDate = %s
        begin
            with h1 as (select a.SecCode2, @tStartDate as TStarDate, @tEndDate as TEndDate, (select max(AdjClose) from SecPrice a1 where a1.SecCode = a.SecCode and a1.FDate >= DATEADD(MONTH, -1, @tStartDate) and a1.FDate < @tStartDate)
            / (select max(AdjClose) from SecPrice where SecCode = a.SecCode and FDate >= DATEADD(MONTH, -6, @tStartDate)  and FDate < @tStartDate) PricePro,
            (select a1.AdjPreClose from SecPrice a1 where a1.SecCode = a.SecCode and a1.FDate = @tStartDate) / (select a1.AdjPreClose from SecPrice a1 where a1.SecCode = a.SecCode and a1.FDate = dateadd(month, -6, @tStartDate)) - 1 as HalfPctChg
            from SecInfo a  where SecCode2 in(%s) and SecType = 'A' )
            select top (@topNum) SecCode2, TStarDate, TEndDate, HalfPctChg  from h1 where PricePro >= 1 order by HalfPctChg desc
        end'''
sql2 = '''with h1 as (select [DAY], (ROW_NUMBER() over (order by [Day]) - 1) / 5 r  from TradeDay a where [Day] > = %s and [Day] < %s)
        , h2 as (select MIN([DAY]) TStartDate, MAX([DAY]) TEndDate from h1 group by r )
        select TStartDate, TEndDate, (select AdjClose from SecPrice where SecCode = '000906.SH' and FDate = h2.TEndDate) / (select AdjClose from SecPrice where SecCode = '000906.SH' and FDate = h2.TStartDate) - 1 IndexReturn from h2'''
sql3 = '''select b.SecCode2, a.FDate, case when a.FDate = %s then a.PctChg else a.AdjClose / (select a1.AdjClose from StockMovingAvg a1 where a1.SecCode = a.SecCode and a1.FDate = %s) - 1 end as PctChg from StockMovingAvg a, SecInfo b where a.SecCode = b.SecCode and b.SecType= 'A' and a.FDate >= %s and a.FDate <= %s and b.SecCode2 in (%s) order by a.FDate '''
conn = pymssql.connect(server, user, password, database, charset='utf8')
cursor = conn.cursor()
# 执行选股收益率统计sql
today = datetime.datetime.now().strftime('%Y%m%d')
daylist = ('-04-30', '-08-30')
sqldata1 = []
finalCode = []
finalDate = []
finalReturn = []
returnSummary = []
startyear = 2006
k = 0
while k <= 10:
    year = startyear + k
    for day in daylist:
        startday = str(year) + day
        cursor.execute('exec [GetNewChooseCode] @fDate = %s', startday)
        data = cursor.fetchall()
# 以10天为一个周期
        cursor.execute(sql2, (data[0][0], data[0][1]))
        dayData = cursor.fetchall()
        codeList = []
        for j in range(0, len(data)):
            codeList.append(data[j][2])
        codeNum = int(len(codeList) / 1.2 - len(codeList) / 1.2 % 10)
        codeString = ('%s,' * len(codeList))[:-1]
        s = '%s'
        sql11 = sql1 % (codeNum, s, s, codeString)
        codeList1 = tuple(codeList)
        for m in range(0, len(dayData)):
            codeList.insert(0, dayData[m][0])
            codeList.insert(1, dayData[m][1])
# 筛选半年创新高股票
            cursor.execute(sql11, tuple(codeList))
            data1 = cursor.fetchall()
            if data1:
                sqldata1.extend(data1)
                newCode = []
                for i in range(0, len(data1)):
                    newCode.append(data1[i][0])
                newCodeStr = ('%s,' * len(newCode))[:-1]
                sql31 = sql3 %(s, s, s, s, newCodeStr)
                newCode.insert(0, dayData[m][0])
                newCode.insert(1, dayData[m][0])
                newCode.insert(2, dayData[m][0])
                newCode.insert(3, dayData[m][1])
# 计算每个周期收益率
                cursor.execute(sql31, tuple(newCode))
                newData = cursor.fetchall()
                newData.sort(key=lambda x: x[1])
                newCodeList = []
                newDateList = []
                newReturnList = []
                upCount = 0
                downCount = 0
                keepCount = 0
                upSum = 0
                downSum = 0
                for i in range(0, len(newData)):
                    if newData[i][0] not in newCodeList:
                        if int(newData[i][2] * 10000) >= 1000 or int(newData[i][2] * 10000) <= -500 or newData[i][1] == dayData[m][1]:
                            newCodeList.append(newData[i][0])
                            newDateList.append(newData[i][1])
                            newReturnList.append(newData[i][2])
                            if newData[i][2] > 0:
                                upCount += 1
                                upSum += newData[i][2]
                            elif newData[i][2] < 0:
                                downCount += 1
                                downSum += newData[i][2]
                            else:
                                keepCount += 1
                returnSummary.append([dayData[m][0], dayData[m][1], sum(newReturnList) / codeNum, dayData[m][2], sum(newReturnList) / codeNum - dayData[m][2], upCount, downCount, keepCount, upSum / codeNum, downSum / codeNum])
                finalCode.extend(newCodeList)
                finalDate.extend(newDateList)
                finalReturn.extend(newReturnList)
            codeList = list(codeList1)
    k += 1
sqldata1.sort(key=lambda x: x[1])
# 将sql执行结果插入excle
workbook = xlwt.Workbook(encoding='utf8')
sheet1 = workbook.add_sheet('sheet1', cell_overwrite_ok=True)
sheet2 = workbook.add_sheet('sheet2', cell_overwrite_ok=True)
field1 = ['Code', 'TStartDate', 'TEndDate']
field2 = ['Code', 'ReturnDate', 'Return']
field3 = ['TStartDate', 'TEndDate', 'TReturn', 'IndexReturn', 'AbReturn', 'UpCount', 'DownCount', 'KeepCount', 'UpReturn', 'DownReturn']
# 导出筛选后的代码
for i in range(0, len(field1)):
    sheet1.write(0, i, field1[i])
for row in range(1, len(sqldata1) + 1):
    for col in range(0, len(field1)):
        sheet1.write(row, col, sqldata1[row - 1][col])
# 导出周期内每只股票的收益率
for i in range(0, len(field2)):
    sheet2.write(0, i, field2[i])
for row in range(1, len(finalCode) + 1):
    sheet2.write(row, 0, finalCode[row - 1])
    sheet2.write(row, 1, finalDate[row - 1])
    sheet2.write(row, 2, finalReturn[row - 1])
# 统计周期内收益情况
for i in range(0, len(field3)):
    sheet2.write(0, i + 4, field3[i])
for row in range(1, len(returnSummary) + 1):
    for col in range(0, len(field3)):
        sheet2.write(row, col + 4, returnSummary[row - 1][col])
workbook.save(r'C:\Users\Administrator\Desktop\QualityReturn%s.xls' % today)
conn.close()
print 'Done!!!'
