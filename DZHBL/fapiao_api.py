# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:

import urllib2
import ssl
import json
import pandas as pd
import MySQLdb

filePath = '../../file/fapiao.xlsx'
df = pd.read_excel(filePath, sheet_name=0)
df = df.fillna('')
allValues = df.values
dbHost = '***'
user = '***'
ps = '***'
dbName = '***'
db = MySQLdb.connect(dbHost, user, ps, dbName, charset='utf8')
db.autocommit(on=True)
cursor = db.cursor()
sql1 = '''INSERT SupplierInvoice VALUES ('%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', %.2f, %.2f, '%s', '%s') '''
sql2 = '''insert InvoiceGoods VALUES ('%s', '%s', %.2f, %.2f, %.2f, %.2f, %.2f)'''

host = 'https://fapiao.market.alicloudapi.com'
path = '/invoice/query'
method = 'GET'
appcode = 'c473a197e0c44a62ace226143c71784d'
ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE
querys = 'fpdm=%s&fphm=%s&kprq=%s&noTaxAmount=%s&checkCode=%s'
for jtem in allValues:
    bodys = {}
    if jtem[4]:
        checkCode = '0' * (6 - len(str(long(jtem[4])))) + str(long(jtem[4]))
        noTaxAmount = ''
    else:
        checkCode = ''
        noTaxAmount = '%.2f' % jtem[3]
    url = host + path + '?' + querys % (
        str(long(jtem[0])), str(long(jtem[1])), str(long(jtem[2])), noTaxAmount, checkCode)
    print '获取发票代码为%s，发票号码为%s，发票日期为%s的发票信息' % (str(long(jtem[0])), str(long(jtem[1])), str(long(jtem[2])))
    request = urllib2.Request(url)
    request.add_header('Authorization', 'APPCODE ' + appcode)
    response = urllib2.urlopen(request, context=ctx)
    content = response.read()
    retDict = json.loads(content)
    if retDict:
        sqlValues1 = [retDict['fpdm'], retDict['fphm'], retDict['fplx'],
                      retDict['code'], retDict['gfMc'], retDict['gfNsrsbh'],
                      retDict['gfContact'].split(' ')[0], retDict['gfContact'].split(' ')[1],
                      retDict['gfBank'].split(' ')[0],
                      retDict['gfBank'].split(' ')[1], retDict['xfMc'], retDict['xfNsrsbh'],
                      retDict['xfContact'].split(' ')[0],
                      retDict['xfContact'].split(' ')[1], retDict['xfBank'].split(' ')[0],
                      retDict['xfBank'].split(' ')[1],
                      float(retDict['sumamount']), float(retDict['goodsamount']), retDict['del'], retDict['kprq']]
        cursor.execute(sql1 % tuple(sqlValues1))
        goodsData = retDict['goodsData']
        for item in goodsData:
            sqlValues2 = [item['name'], item['unit'], float(item['amount']), float(item['priceUnit']),
                          float(item['priceSum']), float(item['taxRate']) * 0.01, float(item['taxSum'])]
            cursor.execute(sql2 % tuple(sqlValues2))
print 'Done!!!'
