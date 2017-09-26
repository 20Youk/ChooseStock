# -*- coding:utf8 -*-
import pymssql
import xlsxwriter
import datetime


def connsqlserver(hostname, database, username, password, sql):
    conn = pymssql.connect(hostname, username, password, database, charset='utf8')
    cursor = conn.cursor()
    cursor.execute(sql)
    sqldata = cursor.fetchall()
    field = cursor.description
    return sqldata, field


def writeintoexcel(filepath, sheetname, field, sqldata):
    wb = xlsxwriter.Workbook(filepath)
    sheet = wb.add_worksheet(sheetname)
    for i in range(0, len(field)):
        sheet.write(0, i, field[i][0])
    for row in range(1, len(sqldata) + 1):
        for col in range(0, len(field)):
            sheet.write(row, col, sqldata[row - 1][col])
    wb.close()


if __name__ == '__main__':
    hostName = '192.168.8.200'
    userName = 'apiread'
    passWord = 'wind,1234'
    dataBase = 'PortfolioData'
    # sql1 = '''select SecCode as 证券交易代码, IndustryName1 as 申万一级行业 from StockIndustry where IndustryType = '1' and EndDate is null  '''
    sql1 = '''with c2 as (select [Day], DATENAME(YEAR, [Day]) * 100 + MONTH([Day]) [Month] from TradeDay where [Day] >= '2016-11-01')
            , c1 as (select max([Day]) MaxDay from c2 group by [Month])
            , c0 as (select a.* from IndexConstituent a, c1 where FDate = MaxDay and IndexCode in( '000016.SH','000905.SH','000300.SH'))
            , t0 as (select distinct(FDate) FDate from c0)
            , t1 as (select FDate, LEAD(DATEADD(day, -1, FDate)) over (order by FDate) EndDate from t0 )
            , t22 as (select FDate, '000000.SH' as IndexCode, SecCode, 0.4 as [Weight] from c0 a where IndexCode = '000300.SH' and SecCode not in (select SecCode from c0 b where b.FDate = a.FDate and b.IndexCode = '000016.SH'))
            , t2 as (select * from t22 union select * from c0)
            , t3 as (select a.FDate, ISNULL(t1.EndDate, '2050-12-31') EndDate, a.IndexCode, a.SecCode, a.[Weight], b.IndustryName1 from t2 a, StockIndustry b, SecInfo c, t1 where a.SecCode = c.SecCode and b.SecCode = c.SecCode2 and c.SecType = 'A' and b.IndustryType = '1' and b.EndDate is null and a.FDate = t1.FDate)
            select a.FDate, IndexCode, IndustryName1, sum([Weight] * PctChg) / sum([Weight]) [Return], sum([Weight] * 0.01) [Weight] from t3, StockMovingAvg a where a.SecCode = t3.SecCode and a.FDate >= t3.FDate and a.FDate <= t3.EndDate  group by a.FDate, t3.IndexCode, IndustryName1 order by 1,2,3
		'''
    sql2 = '''declare @fDate date
            set @fDate = '2016-12-30'
            begin
                with c2 as (select [Day], DATENAME(YEAR, [Day]) * 100 + MONTH([Day]) [Month] from TradeDay where [Day] >= DATEADD(MONTH, -2, @fDate))
                   , c1 as (select max([Day]) MaxDay from c2 group by [Month])
                   , c0 as (select a.* from IndexConstituent a, c1 where FDate = MaxDay and IndexCode in( '000016.SH','000300.SH'))
                   , t2 as (select distinct(FDate) FDate from c0)
                   , t1 as (select FDate, LEAD(DATEADD(day, -1, FDate)) over (order by FDate) EndDate from t2 )
                   , t0 as (select a.FDate, isnull(t1.EndDate, '2050-12-31') as EndDate, '000000.SH' as IndexCode, SecCode, 0.4 as [Weight] from c0 a, t1 where a.FDate = t1.FDate and IndexCode = '000300.SH' and SecCode not in (select SecCode from c0 b where b.FDate = a.FDate and b.IndexCode = '000016.SH'))
                select a.FDate, '000000.SH' as SecCode, AVG(a.AdjClose / a.AdjPreClose - 1) PCtChg from t0, SecPrice a where t0.SecCode = a.SecCode and a.FDate >= t0.FDate and a.FDate <= EndDate and a.FDate >= @fDate group by a.FDate
                union
                select FDate, SecCode, PctChg from StockMovingAvg where SecCode in ('000300.SH', '000905.SH') and FDate >= @fDate
            end
            '''
    sql3 = '''select a.SecCode2, b.IndustryName1 from SecInfo a, StockIndustry b where a.SecCode2 = b.SecCode and a.SecType = 'A' and b.IndustryType = '1' and b.EndDate is null order by 1 '''
    today = datetime.datetime.now().strftime('%Y%m%d')
    filePath1 = 'C:\Users\Administrator\Desktop\IndexIndRe.xlsx'
    filePath2 = 'C:\Users\Administrator\Desktop\IndexReturn.xlsx'
    filePath3 = 'C:\Users\Administrator\Desktop\SW_Industry.xlsx'
    sheetName1 = 'Sheet1'
    sqlData1, field1 = connsqlserver(hostName, dataBase, userName, passWord, sql1)
    writeintoexcel(filePath1, sheetName1, field1, sqlData1)
    sqlData2, field2 = connsqlserver(hostName, dataBase, userName, passWord, sql2)
    writeintoexcel(filePath2, sheetName1, field2, sqlData2)
    sqlData3, field3 = connsqlserver(hostName, dataBase, userName, passWord, sql3)
    writeintoexcel(filePath3, sheetName1, field3, sqlData3)
    print '执行成功!'