#-*- coding:utf8 -*-
import pymssql
import xlwt
import datetime
server = '192.168.8.200'
user = 'sa'
password = 'wind123@pa'
database='PortfolioData'
conn = pymssql.connect(server,user,password, database,charset='utf8')
sql =  '''
declare @fDate date
set @fDate = '2006-04-30'
begin
	declare @startdate date
	declare @endDate date
	declare @reportingDate date
	declare @lastYearDate date
	declare @lastReportingDate date
	set @startdate = @fDate
	set @lastYearDate = DATENAME(YEAR, dateadd(year, -1, @startDate)) + '-12-31'
	if datename(month, @startDate) = '04' begin
		set @endDate = DATEADD(MONTH, 4, @startDate)
		set @reportingDate = DATENAME(YEAR, @startDate)+'-03-31'
		set @lastReportingDate = DATENAME(YEAR, dateadd(year, -1, @startDate)) + '-03-31'
	end
	else begin
		set @endDate = DATEADD(MONTH, 8, @startDate)
		set @reportingDate = DATENAME(YEAR, @startDate)+'-06-30'
		set @lastReportingDate = DATENAME(YEAR, dateadd(year, -1, @startDate)) + '-06-30'
	end
	begin
		---筛选三年前上市、调仓日当天可交易、证券名称非S\*\P开头
		with s0 as (select min([DAY]) Fdate from TradeDay where [DAY] >= @startdate and [Day] < @endDate)
		, s1 as (select c.SecCode2 from SecPrice a, StockNameHistory b, s0, SecInfo c where a.SecCode = c.SecCode and c.SecType = 'A' and b.SecCode = c.SecCode2 and a.FDate = s0.Fdate and b.StartDate <= a.FDate and isnull(b.EndDate, '2050-12-31') >= a.FDate and a.TradeStatus = '1' and substring(b.SecNameAfter, 1, 1) not in ('S','*','P') and c.IPODate <= datename(year, DATEADD(YEAR, -3, @startdate)) + '-01-01')
		---计算总资产、总负债、现金、可交易资产、短期借款、长期借款、净利润的各个TTM值
		, b1 as (select a.SecCode, a.ReportingDate, a.TotleAssets, a.TotleLiab, a.MonetaryCap, a.TradableAssets, a.StBorrow, a.ItBorrow, b.NetIncome from Balanceinfo a, IncomeInfo b where  a.SecCode = b.SecCode and a.ReportingDate = b.ReportingDate and a.ReportingDate = @reportingDate)
		, b2 as (select a.SecCode, a.ReportingDate, a.TotleAssets, a.TotleLiab, a.MonetaryCap, a.TradableAssets, a.StBorrow, a.ItBorrow, b.NetIncome from Balanceinfo a, IncomeInfo b where  a.SecCode = b.SecCode and a.ReportingDate = b.ReportingDate and a.ReportingDate = @lastYearDate)
		, b3 as (select a.SecCode, a.ReportingDate, a.TotleAssets, a.TotleLiab, a.MonetaryCap, a.TradableAssets, a.StBorrow, a.ItBorrow, b.NetIncome from Balanceinfo a, IncomeInfo b where  a.SecCode = b.SecCode and a.ReportingDate = b.ReportingDate and a.ReportingDate = @lastReportingDate)
		, b4 as (select b1.SecCode, b1.ReportingDate, b1.TotleAssets + b2.TotleAssets - b2.TotleAssets TotleAssets, b1.TotleLiab + b2.TotleLiab + b3.TotleLiab TotleLiab, b1.MonetaryCap + b2.MonetaryCap - b3.MonetaryCap MonetaryCap, b1.TradableAssets + b2.TradableAssets - b3.TradableAssets TradableAssets, b1.StBorrow + b2.StBorrow - b3.StBorrow StBorrow, b1.ItBorrow + b2.ItBorrow - b3.ItBorrow LtBorrow,
					b1.NetIncome + b2.NetIncome - b3.NetIncome NetIncome from b1, b2, b3 where b1.SecCode = b2.SecCode and b2.SecCode = b3.SecCode)
		, b5 as (select SecCode, ReportingDate, TotleAssets, case when TotleLiab < 0 then 0 else TotleLiab end TotleLiab, case when StBorrow < 0 then 0 else StBorrow end StBorrow, case when LtBorrow < 0 then 0 else LtBorrow end LtBorrow from b4)
		---1、净利润 / 可投资资本，标准化取[-3,3] (季报TTM)
		, k1 as (select SecCode, ReportingDate, TotleAssets - TotleLiab - MonetaryCap - TradableAssets as InvAssets, NetIncome from b4)
		, k2 as (select SecCode, case when NetIncome > 0 and InvAssets <= 0 then 10 when NetIncome < 0 and InvAssets <= 0 then -10 else NetIncome / InvAssets end as NiIaPor  from k1)
		, k3 as (select AVG(NiIaPor) [Avg], STDEV(NiIaPor) [Stdev] from k2)
		, k4 as (select SecCode, (NiIaPor - [AVG]) / [STDEV] as NiIaStd from k2, k3 where (NiIaPor - [AVG]) / [STDEV] <= 3 and (NiIaPor - [AVG]) / [STDEV] >= -3)
		---2、总资产 / 总负债 、总负债 / (短期借款 + 长期借款),标准化取[-3,3],相加取平均 (季报TTM)
		, l1 as (select SecCode, ReportingDate, case when TotleLiab = 0 then 100 else TotleAssets / TotleLiab end AssLiab, case when StBorrow = 0 and LtBorrow = 0 then 100 else TotleLiab / (StBorrow + LtBorrow) end LiabBorrow from b5)
		, l2 as (select AVG(AssLiab) AvgAssLiab, STDEV(AssLiab) StdevAssLiab, AVG(LiabBorrow) AvgLiabBor, STDEV(LiabBorrow) StdevLiabBor from l1)
		, l3 as (select SecCode, ((AssLiab - AvgAssLiab) / StdevAssLiab + (LiabBorrow - AvgLiabBor) / StdevLiabBor) / 2 AssLiabBor from l1, l2 where (AssLiab - AvgAssLiab) / StdevAssLiab >= -3 and (AssLiab - AvgAssLiab) / StdevAssLiab <= 3 and (LiabBorrow - AvgLiabBor) / StdevLiabBor >= -3 and (LiabBorrow - AvgLiabBor) / StdevLiabBor <=3)
		---3、过去六年的年报 平均ROE_BASIC，标准化取[-3,3]  (年报)
		, r1 as (select SecCode, avg(RoeBasic) AvgRoe, STDEV(RoeBasic) StdevRoe, count(*) [count] from IncomeInfo where ReportingDate <= @lastYearDate and ReportingDate >= DATEADD(YEAR, -5, @lastYearDate) and RIGHT('0'+ltrim(MONTH(ReportingDate)),2) = '12' group by SecCode having STDEV(RoeBasic) is not null )
		, r2 as (select SecCode, (case when [count] = 2 and AvgRoe >= 10 then AvgRoe * 0.5 when [count] = 3 and AvgRoe >= 10 then AvgRoe * 0.75 when [count] >= 4 and AvgRoe >= 10 then AvgRoe when [count] = 2 and AvgRoe < 10 then AvgRoe * 0.5 * 0.5 when [count] = 3 and AvgRoe < 10 then AvgRoe * 0.75 * 0.5 when [count] >= 4 and AvgRoe < 10 then AvgRoe * 0.5 end) / StdevRoe Adj from r1 where StdevRoe <> 0 and [count] >= 2)
		, r3 as (select avg(Adj) AvgAdj, STDEV(Adj) StdevAdj from r2)
		, r4 as (select SecCode, (Adj - AvgAdj) / StdevAdj Adj2 from r2, r3 where (Adj - AvgAdj) / StdevAdj >= -3 and (Adj - AvgAdj) / StdevAdj <= 3)
		----4、过去六年经营产生的净现金平均值 - 净利润，标准化取[-3,3]  (年报)
		, c1 as (select a.SecCode, AVG(b.OperatingNetFlow) / STDEV(b.OperatingNetFlow - a.NetIncome) Adj from IncomeInfo a , CashFlow b where a.SecCode = b.SecCode and a.ReportingDate = b.ReportingDate and a.ReportingDate <= @lastYearDate and a.ReportingDate >= DATEADD(YEAR, -5, @lastYearDate) and RIGHT('0'+ltrim(MONTH(a.ReportingDate)),2) = '12' group by a.SecCode having STDEV(b.OperatingNetFlow - a.NetIncome) <> 0)
		, c2 as (select AVG(Adj) AvgAdj, STDEV(Adj) StdevAdj from c1)
		, c3 as (select SecCode, (Adj - AvgAdj) / StdevAdj Adj2 from c1, c2 where (Adj - AvgAdj) / StdevAdj >= -3 and (Adj - AvgAdj) / StdevAdj <= 3)
		-----将1~4结果取平均
		, re1 as (select * from k4 union select * from l3 union select * from r4 union select * from c3)
		, re4 as (select SecCode, AVG(NiIaStd) AvgRe from re1 group by SecCode)
		----最终筛选:四年净利润都 > 0,三年同比平均 > 10%,最近一个季报净利润 同比 > 15%
		, n1 as (select SecCode, ReportingDate, NetIncome from IncomeInfo where ReportingDate <= @lastYearDate and ReportingDate >= DATEADD(YEAR, -3, @lastYearDate) and RIGHT('0'+ltrim(MONTH(ReportingDate)),2) = '12' )
				/* 四年净利润都 > 0 */
		, n2 as (select SecCode from IncomeInfo where ReportingDate <= @lastYearDate and ReportingDate >= DATEADD(YEAR, -3, @lastYearDate) and RIGHT('0'+ltrim(MONTH(ReportingDate)),2) = '12' and NetIncome > 0 group by SecCode having count(*) = 4)
				/* 三年同比平均 > 0.1 */
		, n3 as (select SecCode, MIN(ReportingDate) MinDate from n1 group by SecCode)
		, n4 as (select n1.SecCode, ReportingDate, NetIncome, case when n1.ReportingDate = n3.MinDate then null else LEAD(NetIncome) over(order by n1.SecCode, n1.ReportingDate desc) end LastNetIncome from n1, n3 where n1.SecCode = n3.SecCode)
		, n5 as (select SecCode, case when NetIncome > 0 and LastNetIncome = 0 then 1 when NetIncome < 0 and LastNetIncome = 0 then -1 when NetIncome = 0 and LastNetIncome = 0 then 0 else (NetIncome - LastNetIncome) / ABS(LastNetIncome) end NetIncomePor from n4 where LastNetIncome is not null)
		, n6 as (select SecCode, AVG(NetIncomePor) AvgPor from n5 group by SecCode having AVG(NetIncomePor) > 0.1)
				/* 最近一个季报净利润 同比 > 0.15 */
		, n7 as (select SecCode, ReportingDate ,NetIncome LastNetIncome from IncomeInfo where ReportingDate = DATEADD(year, -1, @reportingDate))
		, n8 as (select a.SecCode, case when NetIncome > 0 and LastNetIncome = 0 then 1 when NetIncome < 0 and LastNetIncome = 0 then -1 when NetIncome = 0 and LastNetIncome = 0 then 0 else (NetIncome - LastNetIncome) / ABS(LastNetIncome) end NetIncomePor from IncomeInfo a, n7 where a.ReportingDate = @reportingDate and a.SecCode = n7.SecCode)
				/* 净利润整合筛选 */
		, n9 as (select n8.SecCode from n8, n6, n2, s1 where n8.SecCode = n6.SecCode and n6.SecCode = n2.SecCode and n2.SecCode = s1.SecCode2 and n8.NetIncomePor > 0.15)
		----将所有结果合并
		, t1 as (select count(*) MaxID from re4 where re4.SecCode in (select SecCode from n9))
		, t2 as (select ROW_NUMBER() over(order by re4.AvgRe desc) ID, SecCode, AvgRe  from re4 where re4.SecCode in (select SecCode from n9))
		, t as (select t2.SecCode from t2, t1 where ID < =  t1.MaxID / 2.7)
		, d1 as (select a.FDate, a.SecCode, (a.[AdjClose] - a.AdjPreClose) / a.AdjPreClose [Return]
					from SecPrice a, SecInfo b, t where a.SecCode = b.SecCode and b.SecCode2 = t.SecCode and b.SecType = 'A' and a.FDate > = @startDate and a.FDate < @endDate and a.FDate not in (select s0.Fdate from s0)
				union
				select a.FDate, a.SecCode, ((a.AdjHigh + a.AdjLow) / 2 - a.AdjPreClose) / a.AdjPreClose [Return]
					from SecPrice a, SecInfo b, t where a.SecCode = b.SecCode and b.SecCode2 = t.SecCode and b.SecType = 'A' and a.FDate > = @startDate and a.FDate < @endDate and a.FDate in (select s0.Fdate from s0))
		--select FDate, AVG([Return]) - 0.003 * 2 / 240 [Return], (select (a1.[AdjClose] - a1.AdjPreClose) / a1.AdjPreClose from SecPrice a1 where a1.SecCode = '000906.SH' and a1.FDate = d1.FDate) as IndexReturn   from d1 group by FDate
		select * from b4
	end
end
'''
sql2 = '''select a.SecCode, AVG(b.OperatingNetFlow) / STDEV(b.OperatingNetFlow - a.NetIncome) Adj from IncomeInfo a , CashFlow b where a.SecCode = b.SecCode and a.ReportingDate = b.ReportingDate and a.ReportingDate <= '2006' and a.ReportingDate >= DATEADD(YEAR, -5, '2006') and RIGHT('0'+ltrim(MONTH(a.ReportingDate)),2) = '12' group by a.SecCode having STDEV(b.OperatingNetFlow - a.NetIncome) <> 0 '''
cursor = conn.cursor()
date1 = '2007-05-01'
date2 = '2017-01-06'
today = datetime.datetime.now().strftime('%Y%m%d')
cursor.execute('exec GetNewChooseCode @fDate = %s', '2006-04-30')
#cursor.execute('exec [GetPriceChangeNum] @startDate = %s, @endDate = %s',('2017-07-17','2017-07-18'))
#cursor.callproc('GetPriceChangeNum',('2017-07-18', '2017-07-18'))
#cursor.execute(sql)
#data1 = cursor.fetchone()
sqldata = cursor.fetchall()
field1 = cursor.description
conn.close
print  sqldata,'\n',len(sqldata)
#将sql执行结果插入excle
workbook = xlwt.Workbook(encoding='utf8')
sheet = workbook.add_sheet('sheet1',cell_overwrite_ok=True)
for i in range(0,len(field1)):
    sheet.write(0,i,field1[i][0])
for row in range(1,len(sqldata)+1):
    for col in range(0,len(field1)):
        sheet.write(row,col,sqldata[row-1][col])
workbook.save(r'./sqtest%s.xls'% today)