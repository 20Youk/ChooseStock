# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:核验发票

import urllib2
import ssl
import json
import MySQLdb
import ConfigParser
import pymssql
import xlsxwriter
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import make_header
import time
import logging
import sys


log_path = './log/alert/alert_inve_%s' % time.strftime('%Y%m%d') + '.txt'
con = ConfigParser.ConfigParser()
config_path = './config/alert_fapiao.txt'
with open(config_path, 'r') as f:
    con.readfp(f)
    dbHost = con.get('info', 'server')
    user = con.get('info', 'username')
    ps = con.get('info', 'password')
    dbName = con.get('info', 'database')
    srcHost = con.get('info', 'src_server')
    srcUser = con.get('info', 'src_username')
    srcPs = con.get('info', 'src_password')
    srcDBName = con.get('info', 'src_database')
    srcSql = con.get('info', 'src_sql')
    appcode = con.get('info', 'appcode')
    fromaddr = con.get('info', 'from_addr')
    from_pw = con.get('info', 'from_pw')
    mailServer = con.get('info', 'mail_server')
    mailPort = con.getint('info', 'mail_port')
    excel_name = con.get('info', 'excel_name') + time.strftime('%Y%m%d') + '.xlsx'
    excel_path = con.get('info', 'excel_path') + excel_name
    submit = con.get('info', 'submit')
    toaddr = con.get('info', 'to_addr')
    ccaddr = con.get('info', 'cc_addr')
# 设置日志输出
reload(sys)
sys.setdefaultencoding('gbk')
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
# 设置日志输出格式
formatter = logging.Formatter('%(asctime)s [%(levelname)s]  %(name)s : %(message)s')
# 设置日志文件路径、告警级别过滤、输出格式
fh = logging.FileHandler(log_path)
fh.setLevel(logging.WARN)
fh.setFormatter(formatter)
# 设置控制台告警级别、输出格式
ch = logging.StreamHandler()
# ch.setLevel(logging.INFO)
ch.setFormatter(formatter)
# 载入配置
logger.addHandler(fh)
logger.addHandler(ch)


def exp_num(s):
    global s1, s2
    for i in range(0, len(s)):
        if s[i].isdigit():
            s1 = s[:i]
            s2 = s[i:]
            break
    return s1, s2


def tran_number(s):
    global d
    d = 'NULL'
    try:
        d = str('%.2f' % float(s))
        return d
    except (ValueError, TypeError):
        pass
    return d


def write_to_excel(filepath, data1, field1):
    wb = xlsxwriter.Workbook(filepath)
    ws1 = wb.add_worksheet(u'发票查验')
    newstyle = wb.add_format()
    newstyle.set_border(2)
    newstyle.set_font_size(9)
    newstyle.set_font_name(u'宋体')
    newstyle.set_align('left')      # 左对齐
    newstyle.set_align('vcenter')   # 垂直居中
    length = 10
    for i in range(0, len(field1)):
        ws1.write(0, i, field1[i][0], newstyle)
    ws1.write(0, len(field1), u'发票查验结果', newstyle)
    for i in range(0, len(data1)):
        for j in range(0, len(field1) + 1):
            ws1.write(i + 1, j, data1[i][j], newstyle)
    ws1.set_column('A:A', length)
    ws1.set_column('B:B', length)
    ws1.set_column('I:I', length)
    wb.close()
    return


def send_email(from_addr, to_addr, cc_addr, subject, password, data1, field1):
    textstring = '''<p><strong>发票查验失败清单</strong></p>'''
    sheetstring = '''<table width="500" border="2" bordercolor="black" cellspacing="2">
                    <tr>
                    '''
    titles = ''
    for ktem in field1:
        titles = titles + '<td><strong>' + str(ktem[0].encode('utf8')) + '</strong></td>'
    titles = titles + '<td><strong>' + '发票查验结果' + '</strong></td>'
    sheetstring = sheetstring + titles + '</tr>'
    onestring = ''
    if data1:
        try:
            for i in range(0, len(data1)):
                onestring += '<tr>'
                for j in range(0, len(field1) + 1):
                    onedata = data1[i][j]
                    if type(u'a') == type(onedata):
                        onedata = str(onedata.encode('utf8'))
                    else:
                        onedata = str(onedata)
                    onestring = '''{0} <td>{1}</td>'''.format(onestring, onedata)
                onestring += '</tr>'
            onestring = textstring + sheetstring + onestring + '''</table>'''
            # else:
            #     onestring = '查询无结果，请检查任务执行情况，谢谢!'
            msg = MIMEMultipart()
            msg['From'] = u'<%s>' % from_addr
            msg['To'] = to_addr
            msg['Cc'] = cc_addr
            msg['Subject'] = subject
            # --这是文字部分--
            part = MIMEText(onestring, 'html', 'utf-8')
            msg.attach(part)
            # ---这是附件部分---
            # xlsx类型附件
            file_msg = MIMEText(open(excel_path, 'rb').read(), 'base64', 'UTF-8')
            file_msg["Content-Type"] = 'application/octet-stream;name="%s"' % make_header([(excel_name, 'UTF-8')]).encode('UTF-8')
            file_msg["Content-Disposition"] = 'attachment;filename= "%s"' % make_header([(excel_name, 'UTF-8')]).encode('UTF-8')
            msg.attach(file_msg)
            smtp = smtplib.SMTP_SSL(mailServer, mailPort)
            smtp.login(from_addr, password)
            smtp.sendmail(from_addr, to_addr.split(',') + cc_addr.split(','), msg.as_string())
            logger.info(u'邮件发送成功....')
            return True
        except Exception, e1:
            logger.error(e1, exc_info=True)
            return False


def get_src_data():
    conn = pymssql.connect(srcHost, srcUser, srcPs, srcDBName, charset='utf8')
    src_cur = conn.cursor()
    src_cur.execute(srcSql)
    data = src_cur.fetchall()
    field = src_cur.description
    conn.close()
    return data, field


def checking(data1):
    db = MySQLdb.connect(dbHost, user, ps, dbName, charset='utf8')
    db.autocommit(on=True)
    cursor = db.cursor()
    sql1 = '''INSERT INTO SupplierInvoice (InvoiceCode, InvoiceNum, InvoiceType, MarkCode, Buyer, BuyerTaxId,
              BuyerAddress, BuyerBank, BuyerBankNum, Seller, SellerTaxId, SellerAddress, SellerBank,
              SellerBankNum, TotalPrice, NoTaxPrice, IsDel, InvoiceDate) VALUES(
              '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s',  %s,  %s, '%s', '%s')  '''
    sql2 = '''INSERT into InvoiceGoods (GoodName, Unit, Amount, UnitPrice, Price, TaxRate, Tax, InvoiceNum) VALUES (
              '%s', '%s', %s,  %s,  %s,  %s,  %s, '%s') '''
    host = 'https://fapiao.market.alicloudapi.com'
    path = '/invoice/query'
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    querys = 'fpdm=%s&fphm=%s&kprq=%s&noTaxAmount=%s&checkCode=%s'
    alert_data = []
    for jtem in data1:
        one_data = list(jtem)
        try:
            if jtem[7]:
                check_code = '0' * (6 - len(str(long(jtem[7])))) + str(long(jtem[7]))
                no_tax_amount = ''
            else:
                check_code = ''
                no_tax_amount = '%.2f' % jtem[6]
            url = host + path + '?' + querys % (
                str(jtem[3]), str(jtem[4]), str(jtem[5]), no_tax_amount, check_code)
            request = urllib2.Request(url)
            request.add_header('Authorization', 'APPCODE ' + appcode)
            response = urllib2.urlopen(request, context=ctx)
            content = response.read()
            retdict = json.loads(content)
            if retdict['success']:
                if retdict['del'] == 'Y':
                    one_data.append(u'发票已作废')
                    alert_data.append(one_data)
                gfbankname, gfbankcode = exp_num(retdict['gfBank'])
                xfbankname, xfbankcode = exp_num(retdict['xfBank'])
                sqlvalues1 = [retdict['fpdm'], retdict['fphm'], retdict['fplx'],
                              retdict['code'], retdict['gfMc'], retdict['gfNsrsbh'],
                              retdict['gfContact'], gfbankname, gfbankcode,
                              retdict['xfMc'], retdict['xfNsrsbh'],
                              retdict['xfContact'], xfbankname, xfbankcode,
                              tran_number(retdict['sumamount']), tran_number(retdict['goodsamount']), retdict['del'], retdict['kprq']]
                cursor.execute(sql1 % tuple(sqlvalues1))
                goodsdata = retdict['goodsData']
                for item in goodsdata:
                    sqlvalues2 = [item['name'], item['unit'], tran_number(item['amount']), tran_number(item['priceUnit']),
                                  tran_number(item['priceSum']), tran_number(item['taxRate']), tran_number(item['taxSum']), retdict['fphm']]
                    cursor.execute(sql2 % tuple(sqlvalues2))
                logger.info(u'获取发票代码为%s，发票号码为%s，发票日期为%s的发票信息' % (str(jtem[3]), str(jtem[4]), str(jtem[5])))
            else:
                print retdict['data']
        except(urllib2.URLError, Exception, IOError), e:
            logger.error(e, exc_info=True)
            one_data.append(str(e))
            alert_data.append(one_data)
            continue
    return alert_data

if __name__ == "__main__":
    sqlData, sqlField = get_src_data()
    alert_rec = checking(sqlData)
    if alert_rec:
        write_to_excel(excel_path, alert_rec, sqlField)
        if send_email(fromaddr, toaddr, ccaddr, submit, from_pw, alert_rec, sqlField):
            print '邮件发送成功！'
        else:
            print '邮件发送失败，请检查程序，谢谢！'