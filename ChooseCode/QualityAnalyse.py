# -*- coding:utf8 -*-
import pymssql
import xlsxwriter
import datetime
# 质量选股模型
server = '192.168.8.200'
user = 'sa'
password = 'wind123@pa'
database = 'PortfolioData'
conn = pymssql.connect(server, user, password, database, charset='utf8')
cursor = conn.cursor()
# 执行选股收益率统计sql
today = datetime.datetime.now().strftime('%Y%m%d')
# sql0:查询股票21天周期的平均收益率. sql1: 净利润/可投资资本指标筛选结果, sql2: （总资产/总负债 + 总负债/（长期+短期））/2, sql3: 六年平均roe指标, sql4:六年净现金 - 净利润
# sql0 = '''with t0 as (select [DAY], (ROW_NUMBER() over(order by [Day]) - 1) / 21 ID from TradeDay where [Day] >= '2006-04-30' and [Day] <= '2017-06-30')
# , t1 as (select min([Day]) MinDate, MAX([DAY]) MaxDate from t0 group by ID)
# select t1.MinDate, b.SecCode2, AVG(a.PctChg) [Return] from t1, StockMovingAvg a, SecInfo b where a.SecCode = b.SecCode and b.SecType = 'A' and a.FDate >= t1.MinDate and a.FDate <= t1.MaxDate group by b.SecCode2, t1.MinDate
# '''
sql0 = '''with t0 as (select [DAY], (ROW_NUMBER() over(order by [Day]) - 1) / 21 ID from TradeDay where [Day] >= '2006-04-30' and [Day] <= '2017-06-30')
, t1 as (select min([Day]) MinDate, MAX([DAY]) MaxDate from t0 group by ID)
, t2 as (select t1.MinDate, MaxDate, b.SecCode, b.SecCode2, min(FDate) MinFDate, max(FDate) MaxFDate from t1, SecPrice a, SecInfo b where a.SecCode = b.SecCode and b.SecType = 'A' and a.FDate >= t1.MinDate and a.FDate <= MaxDate group by b.SecCode, b.SecCode2, MinDate, MaxDate)
, t3 as (select t2.MinDate, t2.SecCode2, a.AdjClose from t2, SecPrice a  where a.SecCode = t2.SecCode and a.FDate = t2.MinFDate )
, t4 as (select t2.MinDate, t2.SecCode2, a.AdjClose from t2, SecPrice a  where a.SecCode = t2.SecCode and a.FDate = t2.MaxFDate)
select t3.MinDate, t3.SecCode2, t4.AdjClose / t3.AdjClose -1 [Return] from t3, t4 where t4.SecCode2 = t3.SecCode2 and t4.MinDate = t3.MinDate
'''
sql1 = '''---以21个交易日为一周期，获取周期开始及结束日期
with t0 as (select [DAY], (ROW_NUMBER() over(order by [Day]) - 1) / 21 ID from TradeDay where [Day] >= '2006-04-30' and [Day] <= '2017-06-30')
,t1 as (select min([Day]) MinDate, MAX([DAY]) MaxDate from t0 group by ID)
---筛选三年前上市、调仓日当天可交易、证券名称非S\*\P开头
,s0 as (select t1.MinDate, t1.MaxDate, c.SecCode2, c.SecCode from SecPrice a, StockNameHistory b, SecInfo c, t1 where a.SecCode = c.SecCode and c.SecType = 'A' and b.SecCode = c.SecCode2 and a.FDate = t1.MinDate and b.StartDate <= a.FDate and isnull(b.EndDate, '2050-12-31') >= a.FDate and a.TradeStatus = '1' and substring(b.SecNameAfter, 1, 1) not in ('S','*','P') and c.IPODate <= datename(year, DATEADD(YEAR, -3, t1.MinDate)) + '-01-01')
---计算总资产、总负债、现金、可交易资产、短期借款、长期借款、净利润的各个TTM值(上季报 + 去年年报 - 去年同期季报)
, bb as (select SecCode, max(ReportingDate) MaxReportingDate from Balanceinfo group by SecCode)
, b0 as (select a.SecCode, a.ReportingDate, a.PublicDate, case when a.ReportingDate = bb.MaxReportingDate then '2050-12-31' else lead(a.PublicDate) over(order by a.SecCode, a.ReportingDate, a.PublicDate) end as NextPublicDate, a.TotleAssets, a.TotleLiab, a.MonetaryCap, a.TradableAssets, a.StBorrow, a.ItBorrow, b.NetIncome from Balanceinfo a, IncomeInfo b, bb where  a.SecCode = b.SecCode and a.ReportingDate = b.ReportingDate and a.SecCode = bb.SecCode)
, b1 as (select MinDate, MaxDate, SecCode, ReportingDate, PublicDate, NextPublicDate, TotleAssets, TotleLiab, MonetaryCap, TradableAssets, StBorrow, ItBorrow, NetIncome from b0, t1 where PublicDate <= MinDate and NextPublicDate > MinDate)
, b2 as (select MinDate, MaxDate, b0.SecCode, b0.ReportingDate yearReport, b1.ReportingDate, b0.TotleAssets + b1.TotleAssets TotleAssets, b0.TotleLiab + b1.TotleLiab TotleLiab, b0.MonetaryCap + b1.MonetaryCap MonetaryCap, b0.TradableAssets + b1.TradableAssets TradableAssets, b0.StBorrow + b1.StBorrow StBorrow, b0.ItBorrow + b1.ItBorrow ItBorrow, b0.NetIncome + b1.NetIncome NetIncome from b0, b1 where b0.SecCode = b1.SecCode and b0.ReportingDate = datename(year, DATEADD(YEAR, -1, b1.ReportingDate)) + '-12-31')
, b3 as (select MinDate, MaxDate, b0.SecCode, b0.ReportingDate LastReport, b2.ReportingDate, b2.TotleAssets - b0.TotleAssets TotleAssets, b2.TotleLiab - b0.TotleLiab TotleLiab, b2.MonetaryCap - b0.MonetaryCap MonetaryCap, b2.TradableAssets - b0.TradableAssets TradableAssets, b2.StBorrow - b0.StBorrow StBorrow, b2.ItBorrow - b0.ItBorrow LtBorrow, b2.NetIncome - b0.NetIncome NetIncome from b0, b2 where b0.SecCode = b2.SecCode and b0.ReportingDate = DATEADD(year, -1, b2.ReportingDate))
, b4 as (select MinDate, MaxDate, SecCode, ReportingDate, TotleAssets, MonetaryCap, TradableAssets, NetIncome, case when TotleLiab < 0 then 0 else TotleLiab end TotleLiab, case when StBorrow < 0 then 0 else StBorrow end StBorrow, case when LtBorrow < 0 then 0 else LtBorrow end LtBorrow from b3)
---1、净利润 / 可投资资本，标准化取[-3,3] (季报TTM)
, k1 as (select MinDate, MaxDate, SecCode, ReportingDate, TotleAssets - TotleLiab - MonetaryCap - TradableAssets as InvAssets, NetIncome from b4)
, k2 as (select MinDate, MaxDate, SecCode, NetIncome / InvAssets as NiIaPor  from k1 where InvAssets > 0)
, k3 as (select MinDate, AVG(NiIaPor) [Avg], STDEV(NiIaPor) [Stdev] from k2 group by MinDate)
select k2.MinDate, SecCode,  case when (NiIaPor - [AVG]) / [STDEV] > 3 then 3 when (NiIaPor - [AVG]) / [STDEV] < -3 then -3 else (NiIaPor - [AVG]) / [STDEV] end as NiIaStd from k2, k3 where k2.MinDate = k3.MinDate and SecCode in (select SecCode2 from s0 where s0.MinDate = k2.MinDate) order by 3,1
 '''
sql2 = '''---以21个交易日为一周期，获取周期开始及结束日期
with t0 as (select [DAY], (ROW_NUMBER() over(order by [Day]) - 1) / 21 ID from TradeDay where [Day] >= '2006-04-30' and [Day] <= '2017-06-30')
,t1 as (select min([Day]) MinDate, MAX([DAY]) MaxDate from t0 group by ID)
---筛选三年前上市、调仓日当天可交易、证券名称非S\*\P开头
,s0 as (select t1.MinDate, t1.MaxDate, c.SecCode2, c.SecCode from SecPrice a, StockNameHistory b, SecInfo c, t1 where a.SecCode = c.SecCode and c.SecType = 'A' and b.SecCode = c.SecCode2 and a.FDate = t1.MinDate and b.StartDate <= a.FDate and isnull(b.EndDate, '2050-12-31') >= a.FDate and a.TradeStatus = '1' and substring(b.SecNameAfter, 1, 1) not in ('S','*','P') and c.IPODate <= datename(year, DATEADD(YEAR, -3, t1.MinDate)) + '-01-01')
---计算总资产、总负债、现金、可交易资产、短期借款、长期借款、净利润的各个TTM值
, bb as (select SecCode, max(ReportingDate) MaxReportingDate from Balanceinfo group by SecCode)
, b0 as (select a.SecCode, a.ReportingDate, a.PublicDate, case when a.ReportingDate = bb.MaxReportingDate then '2050-12-31' else lead(a.PublicDate) over(order by a.SecCode, a.ReportingDate, a.PublicDate) end as NextPublicDate, a.TotleAssets, a.TotleLiab, a.MonetaryCap, a.TradableAssets, a.StBorrow, a.ItBorrow, b.NetIncome from Balanceinfo a, IncomeInfo b, bb where  a.SecCode = b.SecCode and a.ReportingDate = b.ReportingDate and a.SecCode = bb.SecCode)
, b1 as (select MinDate, MaxDate, SecCode, ReportingDate, PublicDate, NextPublicDate, TotleAssets, TotleLiab, MonetaryCap, TradableAssets, StBorrow, ItBorrow, NetIncome from b0, t1 where PublicDate <= MinDate and NextPublicDate > MinDate)
, b2 as (select MinDate, MaxDate, b0.SecCode, b0.ReportingDate yearReport, b1.ReportingDate, b0.TotleAssets + b1.TotleAssets TotleAssets, b0.TotleLiab + b1.TotleLiab TotleLiab, b0.MonetaryCap + b1.MonetaryCap MonetaryCap, b0.TradableAssets + b1.TradableAssets TradableAssets, b0.StBorrow + b1.StBorrow StBorrow, b0.ItBorrow + b1.ItBorrow ItBorrow, b0.NetIncome + b1.NetIncome NetIncome from b0, b1 where b0.SecCode = b1.SecCode and b0.ReportingDate = datename(year, DATEADD(YEAR, -1, b1.ReportingDate)) + '-12-31')
, b3 as (select MinDate, MaxDate, b0.SecCode, b0.ReportingDate LastReport, b2.ReportingDate, b2.TotleAssets - b0.TotleAssets TotleAssets, b2.TotleLiab - b0.TotleLiab TotleLiab, b2.MonetaryCap - b0.MonetaryCap MonetaryCap, b2.TradableAssets - b0.TradableAssets TradableAssets, b2.StBorrow - b0.StBorrow StBorrow, b2.ItBorrow - b0.ItBorrow LtBorrow, b2.NetIncome - b0.NetIncome NetIncome from b0, b2 where b0.SecCode = b2.SecCode and b0.ReportingDate = DATEADD(year, -1, b2.ReportingDate))
, b4 as (select MinDate, MaxDate, SecCode, ReportingDate, TotleAssets, MonetaryCap, TradableAssets, NetIncome, case when TotleLiab < 0 then 0 else TotleLiab end TotleLiab, case when StBorrow < 0 then 0 else StBorrow end StBorrow, case when LtBorrow < 0 then 0 else LtBorrow end LtBorrow from b3)
---2、总资产 / 总负债 、总负债 / (短期借款 + 长期借款),标准化取[-3,3],相加取平均 (季报TTM)
, l1 as (select MinDate, SecCode, ReportingDate, case when TotleLiab = 0 then 100 else TotleAssets / TotleLiab end AssLiab, case when StBorrow = 0 and LtBorrow = 0 then 100 else TotleLiab / (StBorrow + LtBorrow) end LiabBorrow from b4)
, l2 as (select MinDate, AVG(AssLiab) AvgAssLiab, STDEV(AssLiab) StdevAssLiab, AVG(LiabBorrow) AvgLiabBor, STDEV(LiabBorrow) StdevLiabBor from l1 group by MinDate)
select l1.MinDate, SecCode, (case when (AssLiab - AvgAssLiab) / StdevAssLiab > 3 then 3 when (AssLiab - AvgAssLiab) / StdevAssLiab < -3 then -3 else (AssLiab - AvgAssLiab) / StdevAssLiab end + case when (LiabBorrow - AvgLiabBor) / StdevLiabBor > 3 then 3 when
(LiabBorrow - AvgLiabBor) / StdevLiabBor < -3 then -3 else (LiabBorrow - AvgLiabBor) / StdevLiabBor end ) / 2  AssLiabStd from l1, l2 where l1.MinDate = l2.MinDate and SecCode in (select SecCode2 from s0 where s0.MinDate = l1.MinDate) order by 1,3
'''
sql3 = '''---以21个交易日为一周期，获取周期开始及结束日期
with t0 as (select [DAY], (ROW_NUMBER() over(order by [Day]) - 1) / 21 ID from TradeDay where [Day] >= '2006-04-30' and [Day] <= '2017-06-30')
,t1 as (select min([Day]) MinDate, MAX([DAY]) MaxDate from t0 group by ID)
---筛选三年前上市、调仓日当天可交易、证券名称非S\*\P开头
, s0 as (select t1.MinDate, t1.MaxDate, c.SecCode2, c.SecCode from SecPrice a, StockNameHistory b, SecInfo c, t1 where a.SecCode = c.SecCode and c.SecType = 'A' and b.SecCode = c.SecCode2 and a.FDate = t1.MinDate and b.StartDate <= a.FDate and isnull(b.EndDate, '2050-12-31') >= a.FDate and a.TradeStatus = '1' and substring(b.SecNameAfter, 1, 1) not in ('S','*','P') and c.IPODate <= datename(year, DATEADD(YEAR, -3, t1.MinDate)) + '-01-01')
---3、过去六年的年报 平均ROE_BASIC，标准化取[-3,3]  (年报)
, r1 as (select SecCode, MinDate, avg(RoeBasic) AvgRoe, STDEV(RoeBasic) StdevRoe, count(*) [count] from IncomeInfo, t1 where ReportingDate < MinDate and ReportingDate >= DATEADD(YEAR, -5, MinDate) and RIGHT('0'+ltrim(MONTH(ReportingDate)),2) = '12' and SecCode in (select SecCode2 from s0 where s0.MinDate = t1.MinDate) group by SecCode, MinDate having STDEV(RoeBasic) is not null )
, r2 as (select SecCode, MinDate, (case when [count] = 2 and AvgRoe >= 10 then AvgRoe * 0.5 when [count] = 3 and AvgRoe >= 10 then AvgRoe * 0.75 when [count] >= 4 and AvgRoe >= 10 then AvgRoe when [count] = 2 and AvgRoe < 10 then AvgRoe * 0.5 * 0.5 when [count] = 3 and AvgRoe < 10 then AvgRoe * 0.75 * 0.5 when [count] >= 4 and AvgRoe < 10 then AvgRoe * 0.5 end) / StdevRoe Adj from r1 where StdevRoe <> 0 and [count] >= 2)
, r3 as (select MinDate, avg(Adj) AvgAdj, STDEV(Adj) StdevAdj from r2 group by MinDate)
select r2.MinDate, SecCode,case when (Adj - AvgAdj) / StdevAdj > 3 then 3 when (Adj - AvgAdj) / StdevAdj < -3 then -3 else(Adj - AvgAdj) / StdevAdj end as Adj2 from r2, r3 where r2.MinDate = r3.MinDate order by 1,3
'''
sql4 = '''---以21个交易日为一周期，获取周期开始及结束日期
with t0 as (select [DAY], (ROW_NUMBER() over(order by [Day]) - 1) / 21 ID from TradeDay where [Day] >= '2006-04-30' and [Day] <= '2017-06-30')
, t1 as (select min([Day]) MinDate, MAX([DAY]) MaxDate from t0 group by ID)
---筛选三年前上市、调仓日当天可交易、证券名称非S\*\P开头
, s0 as (select t1.MinDate, t1.MaxDate, c.SecCode2, c.SecCode from SecPrice a, StockNameHistory b, SecInfo c, t1 where a.SecCode = c.SecCode and c.SecType = 'A' and b.SecCode = c.SecCode2 and a.FDate = t1.MinDate and b.StartDate <= a.FDate and isnull(b.EndDate, '2050-12-31') >= a.FDate and a.TradeStatus = '1' and substring(b.SecNameAfter, 1, 1) not in ('S','*','P') and c.IPODate <= datename(year, DATEADD(YEAR, -3, t1.MinDate)) + '-01-01' )
----4、过去六年经营产生的净现金平均值 - 净利润，标准化取[-3,3]  (年报)
, c1 as (select MinDate, a.SecCode, AVG(b.OperatingNetFlow) / STDEV(b.OperatingNetFlow - a.NetIncome) Adj from IncomeInfo a , CashFlow b, t1 where a.SecCode = b.SecCode and a.ReportingDate = b.ReportingDate and a.ReportingDate < MinDate and a.ReportingDate >= DATEADD(YEAR, -5, MinDate) and RIGHT('0'+ltrim(MONTH(a.ReportingDate)),2) = '12' and a.SecCode in (select SecCode2 from s0 where s0.MinDate = t1.MinDate) group by a.SecCode, MinDate having STDEV(b.OperatingNetFlow - a.NetIncome) <> 0 )
, c2 as (select MinDate, AVG(Adj) AvgAdj, STDEV(Adj) StdevAdj from c1 group by MinDate)
select c1.MinDate, SecCode, case when (Adj - AvgAdj) / StdevAdj < -3 then -3 when  (Adj - AvgAdj) / StdevAdj > 3 then 3 else (Adj - AvgAdj) / StdevAdj end as Adj2 from c1, c2 where c1.MinDate = c2.MinDate order by 1, 3
'''
cursor.execute(sql0)
data0 = cursor.fetchall()
cursor.execute(sql1)
data1 = cursor.fetchall()
cursor.execute(sql2)
data2 = cursor.fetchall()
cursor.execute(sql3)
data3 = cursor.fetchall()
cursor.execute(sql4)
data4 = cursor.fetchall()
dict0 = {(item[0], item[1]): [item[2]] for item in data0}
dict1 = {(item[0], item[1]): [item[2]] for item in data1}
dict2 = {(item[0], item[1]): [item[2]] for item in data2}
dict3 = {(item[0], item[1]): [item[2]] for item in data3}
dict4 = {(item[0], item[1]): [item[2]] for item in data4}
for i in dict1.keys():
    dict1[i].append(dict0[i][0])
for i in dict2.keys():
    dict2[i].append(dict0[i][0])
for i in dict3.keys():
    dict3[i].append(dict0[i][0])
for i in dict4.keys():
    dict4[i].append(dict0[i][0])
sqldata1 = sorted([[j[0], j[1], dict1[j][0], float(dict1[j][1])] for j in dict1.keys()], key=lambda x: (x[0], x[1]))
sqldata2 = sorted([[j[0], j[1], dict2[j][0], float(dict2[j][1])] for j in dict2.keys()], key=lambda x: (x[0], x[1]))
sqldata3 = sorted([[j[0], j[1], dict3[j][0], float(dict3[j][1])] for j in dict3.keys()], key=lambda x: (x[0], x[1]))
sqldata4 = sorted([[j[0], j[1], dict4[j][0], float(dict4[j][1])] for j in dict4.keys()], key=lambda x: (x[0], x[1]))
# 将sql执行结果插入excle
workbook = xlsxwriter.Workbook('C:\Users\Administrator\Desktop\QualityAnalyse_%s.xlsx' % today)
sheet1 = workbook.add_worksheet(u'可投资资本指标')
sheet2 = workbook.add_worksheet(u'负债指标')
sheet3 = workbook.add_worksheet(u'六年roe指标')
sheet4 = workbook.add_worksheet(u'净现金指标')
field1 = ['MinDate', 'Code', 'TradeAssAdj', 'Return']
field2 = ['MinDate', 'Code', 'LiabAdj', 'Return']
field3 = ['MinDate', 'Code', 'RoeAdj', 'Return']
field4 = ['MinDate', 'Code', 'CashAdj', 'Return']
for k in range(0, len(field1)):
    sheet1.write(0, k, field1[k])
for row in range(1, len(sqldata1) + 1):
    for col in range(0, len(field1)):
        sheet1.write(row, col, sqldata1[row - 1][col])
for k in range(0, len(field2)):
    sheet2.write(0, k, field2[k])
for row in range(1, len(sqldata2) + 1):
    for col in range(0, len(field2)):
        sheet2.write(row, col, sqldata2[row - 1][col])
for k in range(0, len(field3)):
    sheet3.write(0, k, field3[k])
for row in range(1, len(sqldata3) + 1):
    for col in range(0, len(field3)):
        sheet3.write(row, col, sqldata3[row - 1][col])
for k in range(0, len(field4)):
    sheet4.write(0, k, field4[k])
for row in range(1, len(sqldata4) + 1):
    for col in range(0, len(field4)):
        sheet4.write(row, col, sqldata4[row - 1][col])
workbook.close()
conn.close()
print 'Done!!!'
