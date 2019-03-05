# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:导出DZHBL数据库的融资申请中及已放款数据到Excel
import pymssql
import xlsxwriter
import time
import os


def read_sql(host, user, password, database, sql):
    conn = pymssql.connect(host, user, password, database, charset='utf8')
    cursor = conn.cursor()
    cursor.execute(sql)
    data = cursor.fetchall()
    field = cursor.description
    conn.close()
    return data, field


def write_to_excel(filepath, data, field):
    wb = xlsxwriter.Workbook(filepath)
    ws = wb.add_worksheet('Sheet1')
    newstyle = wb.add_format()
    sumstyle = wb.add_format()
    newstyle.set_border(2)
    newstyle.set_font_size(9)
    newstyle.set_font_name(u'宋体')
    newstyle.set_align('left')      # 左对齐
    newstyle.set_align('vcenter')   # 垂直居中
    sumstyle.set_border(2)
    sumstyle.set_font_size(9)
    sumstyle.set_font_name(u'宋体')
    sumstyle.set_align('left')
    sumstyle.set_align('vcenter')
    sumstyle.set_pattern(1)
    sumstyle.set_bg_color('yellow')
    k = 1
    n = 0
    for i in range(0, len(field)):
        ws.write(0, i, field[i][0], newstyle)
    for i in range(0, len(data)):
        print data[i - 1][11], type(data[i - 1][11]), not data[i - 1][11]
        if data[i][11] and (not data[i - 1][11]):
            n = i
            ws.write(k, 8, u'未放款合计', sumstyle)
            ws.write(k, 9, sum([data[m][9] for m in range(0, i)]), sumstyle)
            ws.write(k, 10, sum([data[m][10] for m in range(0, i)]), sumstyle)
            k += 1
        for j in range(0, len(field)):
            ws.write(k, j, data[i][j], newstyle)
        k += 1
    ws.write(k, 8, u'放款合计', sumstyle)
    ws.write(k, 9, sum([data[m][9] for m in range(n, len(data))]), sumstyle)
    ws.write(k, 10, sum([data[m][10] for m in range(n, len(data))]), sumstyle)
    wb.close()
    return

if __name__ == '__main__':
    month = int(time.strftime('%m'))
    today = time.strftime('%Y%m%d')
    # filePath = '../../file/%s' % today
    filePath = '/home/vftpuser/public/融资申请中及已放款数据/%s' % today
    if not os.path.exists(filePath):
        os.mkdir(filePath)
    excel = filePath + '/%d月融资申请中与已放款数据%s.xlsx' % (month, today)
    server = '***'
    userName = '***'
    passWord = '***'
    dataBase = '***'
    sql1 = '''declare @oneday datetime
            set @oneday = convert(date, dateadd(dd, -day(getdate())+1, GETDATE()), 120);
            with t0 as (select t1.DocEntry, convert(decimal(18,2), sum(t1.LineTotal)) as SumLineTotle, convert(decimal(18,2), sum(t1.ALineTotal)) as SumALineTotle from U_RZQR1 t1 group by t1.DocEntry)
            , t00 as (select t1.DocEntry, t2.BaseEntry,t2.QCode, t2.BDate, t2.BTotal from U_OPFK t1, U_OPFK1 t2 where t1.DocEntry = t2.DocEntry and  t2.BaseEntry is not null and t1.DocType = 'T')
            , t01 as (SELECT b.BaseEntry, b.DocDate FROM U_OPFK A INNER JOIN U_OPFK1 B ON A.DocEntry=B.DocEntry WHERE ISNULL(B.DocDate,'')<>'' AND A.DocType='F')
            , t as (select DocEntry,sum(LineTotal) LineTotle, RType from U_BLRZ group by DocEntry, RType)
            select t.DocEntry as 系统编号,
                t00.DocEntry as 结算申请单编号,
                t.RZHCode as 融资申请编号,
                t.CardName as 融资方名称,
                t.BCardName as 买房酒店名称,
                case when t.RzType = 1 then '票前融资' else '票后融资' end as 融资方式,
                SUBSTRING(convert(varchar(100), t.RMoth, 120), 0, 8) as 融资月份,
                convert(date, t.SDate, 120) as 账期开始日期,
                convert(date, t.EDate, 120) as 账期结束日期,
                t0.SumLineTotle as 应收款总额,
                t0.SumALineTotle as 应放款总额,
                convert(date, t01.DocDate, 120) as 放款日期,
                convert(date, t00.BDate, 120) as 回款日期,
                t00.BTotal as 回款总额,
                CASE WHEN
                EXISTS(SELECT 1 FROM U_OPFK A
                INNER JOIN U_OPFK1 B ON A.DocEntry=B.DocEntry AND B.BaseEntry=t.DocEntry
                WHERE B.TKTotal IS NOT NULL AND A.DocType='T') 　
                THEN '√' ELSE NULL END AS 财务结算放款确认,
                convert(date, dateadd(mm, 1, dateadd(dd, -day(t.RMoth)+1, t.RMoth)), 120) as 实际账期开始日期,
                DATEDIFF(dd,convert(date, dateadd(mm, 1, dateadd(dd, -day(t.RMoth)+1, t.RMoth)), 120), convert(date, t00.BDate, 120)) as 实际账期,
                convert(int,t1.JC) as 系统设置帐期
            from U_RZQR t inner join t0 on t.DocEntry = t0.DocEntry
            inner join U_JCSJ1 t1 on t.BCardName = t1.BCardName and t.CardName = t1.CardName
            left join t00 on t.DocEntry = t00.BaseEntry
            left join t01 on t.DocEntry = t01.BaseEntry
            where t.RZHCode is not null
            and (convert(date, t.SDate, 120) >= @oneday or convert(date, t01.DocDate, 120)>= @oneday)
            union
            SELECT
            t0.DocEntry as '系统编号',null as 结算申请单编号,
            T0.DocCode AS '融资编号',T1.CardName AS '融资方名称',T2.CardName AS '买方酒店名称',
            case when t.RType = 1 then '票前融资' else '票后融资' end as 融资方式,
            SUBSTRING(convert(varchar(100), t0.U_SDate, 120), 0, 8) AS '融资月份',
            convert(date,t0.createDate, 120) as '账期开始日期',convert(date, dateadd(DD, t3.jc, t0.createDate), 120) as '账期结束日期',
            t.LineTotle as '应收款总额',
            t.LineTotle * t4.RRate as 应放款总额,null as 放款日期,null as	回款日期,null as 回款总额,null as 财务结算放款确认,null as	实际账期开始日期,null as 实际账期,null as 系统设置帐期
            FROM T_BLRZ T0
            INNER JOIN T_OCRD T1 ON T0.OutCID=T1.BPAccount
            INNER JOIN T_OCRD T2 ON T0.AcctCardName=T2.CardCode
            INNER JOIN t on t0.DocEntry=t.DocEntry
            inner join U_JCSJ1 t3 on t1.CardName = t3.CardName and t2.CardName = t3.BCardName
            left join U_SXFA t4 on t1.CardName = t4.CardName and t2.CardName = t4.BCardName and t4.BLType = 'B'
            where t0.DocCode not in (select RZHCode from U_RZQR where DocType= 'Q')
            and convert(date,t0.createDate, 120) >= @oneday
            order by 放款日期, 结算申请单编号 desc

            '''
    sqlData, field1 = read_sql(server, userName, passWord, dataBase, sql1)
    write_to_excel(excel, sqlData, field1)
    print 'Done!'
