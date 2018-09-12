# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:查询上个月放款数据汇总
import xlsxwriter
import pymssql
import datetime
from dateutil.relativedelta import relativedelta


def conn_sql(hostname, database, username, password, sql):
    conn = pymssql.connect(hostname=hostname, database=database, user=username, password=password, charset='utf8')
    cursor = conn.cursor()
    cursor.execute(sql)
    sqldata = cursor.fetchall()
    field = cursor.description
    return sqldata, field

if __name__ == '__main__':
    host = '***'
    db = '***'
    user = '***'
    pw = '***'
    sql1 = '''with t2 as (select t1.DocEntry, convert(decimal(18,2), sum(t1.LineTotal)) as SumLineTotle, convert(decimal(18,2), sum(t1.ALineTotal)) as SumALineTotle
                from U_RZQR1 t1 group by t1.DocEntry)
                select t1.RZHCode as '融资申请编号',
                t.CardName as '融资方名称',
                t.BCardName as '买方酒店名称',
                case when t1.RzType = 1 then '票前融资' else '票后融资' end as 融资方式,
                year(t1.RMoth)+MONTH(t1.RMoth)*0.01 as 融资月份,
                convert(date, t.SDate, 120) as '账期开始日期',
                convert(date, t.EDate, 120) as '帐期结束日期',
                t2.SumLineTotle as 应收款总额,
                t2.SumALineTotle as 应放款总额,
                convert(date, t.DocDate, 120) as '放款日期'
                 from U_OPFK1 t, U_RZQR t1 , t2, U_OPFK t3
                 where t.DocEntry = t3.DocEntry
                 and t3.DocType = 'F'
                 and t.QCode = t1.HCode
                 and t1.DocType = 'Q'
                 and t1.DocEntry = t2.DocEntry
                 and convert(date, t.DocDate, 120) >= '%s'
                 and convert(date, t.DocDate, 120) < '%s'
                 order by convert(date, t.DocDate, 120) '''
    firstDay = datetime.datetime.now()
    lastDay = datetime.datetime.now() + relativedelta(months=+1)
    changedSql = sql1 % (firstDay.strftime('%Y-%m-%d'), lastDay.strftime('%Y-%m-%d'))
    sqlData1, field1 = conn_sql(host, db, user, pw, changedSql)
    excelPath = '../file/loan_%s.xlsx' % firstDay.strftime('%Y-%m-%d')
    wb = xlsxwriter.Workbook(excelPath)
    ws = wb.add_worksheet(u'放款明细')
    for j in range(0, len(field1)):
        ws.write(0, j, field1[0])
    dataList = list(sqlData1)
    for i in range(0, len(sqlData1)):
        ws.write_row(i + 1, 0, list(sqlData1[i]))
    wb.close()
