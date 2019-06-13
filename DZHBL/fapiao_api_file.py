# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:

import urllib2
import ssl
import json
import pandas as pd
import MySQLdb
import ConfigParser


def exp_num(s):
    global s1, s2
    for i in range(0, len(s)):
        if s[i].isdigit():
            s1 = s[:i]
            s2 = s[i:]
            break
    return s1, s2

filePath = '../../file/fapiao.xlsx'
df = pd.read_excel(filePath, sheet_name=0)
df = df.fillna('')
allValues = df.values
con = ConfigParser.ConfigParser()
config_path = '../../config/GCCFSI.txt'
with open(config_path, 'r') as f:
    con.readfp(f)
    dbHost = con.get('info', 'server')
    user = con.get('info', 'username')
    ps = con.get('info', 'password')
    dbName = con.get('info', 'database')
db = MySQLdb.connect(dbHost, user, ps, dbName, charset='utf8')
db.autocommit(on=True)
cursor = db.cursor()
sql1 = '''INSERT SupplierInvoice VALUES ('%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', %.2f, %.2f, '%s', '%s') '''
sql2 = '''insert InvoiceGoods VALUES ('%s', '%s', %.2f, %.2f, %.2f, %.2f, %.2f, '%s')'''

host = 'https://fapiao.market.alicloudapi.com'
path = '/invoice/query'
method = 'GET'
appcode = 'c473a197e0c44a62ace226143c71784d'
ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE
querys = 'fpdm=%s&fphm=%s&kprq=%s&noTaxAmount=%s&checkCode=%s'
for jtem in allValues:
    try:
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
        if retDict['success']:
            gfBankName, gfBankCode = exp_num(retDict['gfBank'])
            xfBankName, xfBankCode = exp_num(retDict['xfBank'])
            sqlValues1 = [retDict['fpdm'], retDict['fphm'], retDict['fplx'],
                          retDict['code'], retDict['gfMc'], retDict['gfNsrsbh'],
                          retDict['gfContact'], gfBankName, gfBankCode,
                          retDict['xfMc'], retDict['xfNsrsbh'],
                          retDict['xfContact'], xfBankName, xfBankCode,
                          float(retDict['sumamount']), float(retDict['goodsamount']), retDict['del'], retDict['kprq']]
            cursor.execute(sql1 % tuple(sqlValues1))
            goodsData = retDict['goodsData']
            for item in goodsData:
                sqlValues2 = [item['name'], item['unit'], float(item['amount']), float(item['priceUnit']),
                              float(item['priceSum']), float(item['taxRate']) * 0.01, float(item['taxSum']), retDict['fphm']]
                cursor.execute(sql2 % tuple(sqlValues2))
        else:
            print retDict['data']
    except(urllib2.URLError, Exception, IOError), e:
        print '程序出错，异常信息：%s' % e
        continue
print 'Done!!!'
