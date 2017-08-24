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
    try:
        hostName = '192.168.8.200'
        userName = 'apiread'
        passWord = 'wind,1234'
        dataBase = 'PortfolioData'
        # sql1 = '''select SecCode as 证券交易代码, IndustryName1 as 申万一级行业 from StockIndustry where IndustryType = '1' and EndDate is null  '''
        sql1 = '''declare @fDate date
                set @fDate = '2016-12-30'
                begin
                    with t0 as (select FDate, LEAD(dateadd(day, -1, FDate)) over (order by SecCode, FDate) EndDate, SecCode from IndexConstituent a where IndexCode = '000300.SH' and FDate >= @fDate and SecCode not in
                    (select SecCode from IndexConstituent b where b.IndexCode = '000016.SH' and b.FDate =
                    (select top 1 FDate from IndexConstituent c where c.IndexCode = '000016.SH' and c.FDate <= a.FDate order by FDate desc)))
                    select a.FDate, '000000.SH' as SecCode, AVG(a.AdjClose / a.AdjPreClose - 1) PCtChg from t0, SecPrice a where t0.SecCode = a.SecCode and t0.FDate <= a.FDate and t0.EndDate >= a.FDate group by a.FDate
                    union
                    select FDate, SecCode, PctChg from StockMovingAvg where SecCode in ('000300.SH', '000905.SH') and FDate >= @fDate
                end '''
        today = datetime.datetime.now().strftime('%Y%m%d')
        filePath = 'C:\Users\Administrator\Desktop\/test_%s.xlsx' % today
        sheetName = 'Sheet1'
        sqlData, field1 = connsqlserver(hostName, dataBase, userName, passWord, sql1)
        writeintoexcel(filePath, sheetName, field1, sqlData)
    except Exception, e:
        print '执行失败，请检查错误信息:\n %s' % e
    else:
        print '执行成功!'