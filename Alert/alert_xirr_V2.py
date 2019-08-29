# -*- coding:utf8 -*-
# 应用: 计算全公司整体XIRR
import pymssql
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import ConfigParser
import logging
import sys
import time
import datetime
from dateutil.relativedelta import relativedelta
from scipy import optimize
import xlsxwriter
from email.header import make_header
from decimal import *

log_path = './log/alert/alert_irr_%s' % time.strftime('%Y%m%d') + '.txt'
con = ConfigParser.ConfigParser()
config_path = './config/config.txt'
alert_path = './config/alert_irr.txt'
with open(config_path, 'r') as f:
    con.readfp(f)
    hostname = con.get('info', 'server')
    user = con.get('info', 'username')
    sqlpassword = con.get('info', 'password')
    database = con.get('info', 'database')
    fromaddr = con.get('info', 'from_addr')
    from_pw = con.get('info', 'from_pw')
    mailServer = con.get('info', 'mail_server')
    mailPort = con.getint('info', 'mail_port')
with open(alert_path, 'r') as g:
    con.readfp(g)
    toaddr = con.get('info', 'to_addr')
    ccaddr = con.get('info', 'cc_addr')
    sql = con.get('info', 'sql')
    submit = con.get('info', 'submit')
    excel_name = con.get('info', 'excel_name') + time.strftime('%Y%m%d') + '.xlsx'
    excel_path = con.get('info', 'excel_path') + excel_name

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


def xnpv(rate, cashflows):
    return sum([cf / (1 + rate) ** ((t - cashflows[0][0]).days / 365.0) for (t, cf) in cashflows])


def xirr(cashflows, guess=0.1):
    try:
        re = optimize.newton(lambda r: xnpv(r, cashflows), guess)
        return str(round(re, 4) * 100) + '%', round(re, 4) * 100
    except:
        print('Calc Wrong')


def connect_db(startday, endday):
    # 连接数据库
    conn = pymssql.connect(hostname, user, sqlpassword, database, charset='utf8')
    cursor = conn.cursor()
    cursor.execute(sql + "'%s','%s'" % (datetime.datetime.strftime(startday, '%Y-%m-%d'), datetime.datetime.strftime(endday, '%Y-%m-%d')))
    data = cursor.fetchall()
    des = cursor.description
    sql1 = '''select convert(numeric(18,2),sum(t1.SJYFTotal)) Total from U_OPFK1 t1 inner join U_OPFK t0 on t0.DocEntry = t1.DocEntry
                where convert(date, t1.DocDate, 120) >= '%s' and convert(date, t1.DocDate, 120) <= '%s'
                and t0.DocType = 'F'
          '''
    sql2 = '''select convert(numeric(18,2), sum(isnull(t1.BTotal, t1.HGTOtal))) BTotal
                from U_OPFK t0 inner join U_OPFK1 t1 on t0.DocEntry = t1.DocEntry
                where t0.DocType = 't' and convert(date, t1.TkDate, 120) >= '%s'
                and convert(date, t1.TkDate, 120) <= '%s' '''
    cursor.execute(sql1 % (datetime.datetime.strftime(startday, '%Y-%m-%d'), datetime.datetime.strftime(endday, '%Y-%m-%d')))
    data1 = cursor.fetchall()
    cursor.execute(sql2 % (datetime.datetime.strftime(startday, '%Y-%m-%d'), datetime.datetime.strftime(endday, '%Y-%m-%d')))
    data2 = cursor.fetchall()
    cursor.close()
    if data:
        alldata = [(datetime.datetime.strptime(item[0], '%Y-%m-%d'), float(item[1])) for item in data]
        allxirr, g1 = xirr(alldata)
        btotal = float(data2[0][0])
        return allxirr, btotal, float(data1[0][0]), data, des
    else:
        return 0, 0, 0


def write_to_excel(filepath, data1, field1):
    wb = xlsxwriter.Workbook(filepath)
    ws1 = wb.add_worksheet(u'IRR')
    newstyle = wb.add_format()
    newstyle.set_border(2)
    newstyle.set_font_size(9)
    newstyle.set_font_name(u'宋体')
    newstyle.set_align('left')  # 左对齐
    newstyle.set_align('vcenter')  # 垂直居中
    for i in range(0, len(field1)):
        ws1.write(0, i, field1[i][0], newstyle)
    for i in range(0, len(data1)):
        for j in range(0, len(field1)):
            ws1.write(i + 1, j, data1[i][j], newstyle)
    wb.close()
    return


def send_email(from_addr, to_addr, cc_addr, subject, password, irr1, income1, ftotal1, irr2, income2, ftotal2, data1, field1):
    try:
        today = time.strftime('%Y-%m-%d')
        if ftotal2 == 0:
            textstring = '''<p><strong>您好，<br/>&emsp;截止%s，上周IRR（剔除资金成本）为：%s，上周总放款为：%s元，上周总回款为：%s元。 </strong></p>''' % (today, irr1, format(ftotal1, ','), format(income1, ','))
        else:
            textstring = '''<p><strong>您好，<br/>&emsp;截止%s，上周IRR（剔除资金成本）为：%s，上周总放款为：%s元，上周总回款为：%s元。<br/>
                            &emsp;截止%s，上月IRR（剔除资金成本）为：%s，上月总放款为：%s元，上月总回款为：%s元。
                            </strong></p>''' % (today, irr1, format(ftotal1, ','), format(income1, ','), today, irr2, format(ftotal2, ','), format(income2, ','))
        sheetstring = '''<table width="500" border="2" bordercolor="black" cellspacing="2"><tr>'''
        titles = ''
        for item in field1:
            titles = titles + '<td><strong>' + str(item[0].encode('utf8')) + '</strong></td>'
        sheetstring = sheetstring + titles + '</tr>'
        onestring = ''
        for i in range(0, len(data1)):
            onestring += '<tr>'
            for j in range(0, len(field1)):
                onedata = data1[i][j]
                if type(u'a') == type(onedata):
                    onedata = str(onedata.encode('utf8'))
                elif type(Decimal('0.1')) == type(onedata):
                    onedata = format(data1[i][j], ',')
                else:
                    onedata = str(onedata)
                onestring = '''{0} <td>{1}</td>'''.format(onestring, onedata)
            onestring += '</tr>'
        onestring = textstring + sheetstring + onestring + '''</table>'''
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
        file_msg["Content-Type"] = 'application/octet-stream;name="%s"' % make_header(
                [(excel_name, 'UTF-8')]).encode('UTF-8')
        file_msg["Content-Disposition"] = 'attachment;filename= "%s"' % make_header([(excel_name, 'UTF-8')]).encode(
            'UTF-8')
        msg.attach(file_msg)
        smtp = smtplib.SMTP_SSL(mailServer, mailPort)
        smtp.login(from_addr, password)
        smtp.sendmail(from_addr, to_addr.split(',') + cc_addr.split(','), msg.as_string())
        logger.info(u'邮件发送成功....')
        return True
    except Exception, e:
        logger.error(e, exc_info=True)
        return False


if __name__ == "__main__":
    now = datetime.datetime.now()
    MonDay = now + datetime.timedelta(days=-7)
    SunDay = now + datetime.timedelta(days=-1)
    if MonDay.day <= 7:
        monLastDay = MonDay + datetime.timedelta(days=-MonDay.day)
        monFirstDay = MonDay + datetime.timedelta(days=1 - MonDay.day) - relativedelta(months=+1)
        monthIrr, monthInc, monthTol = connect_db(monFirstDay, monLastDay)
        yearFirstDay = monFirstDay - relativedelta(months=+monFirstDay.month - 1)
        yearIrr, yearInc, yearTol = connect_db(yearFirstDay, SunDay)
    else:
        monthIrr, monthInc, monthTol = 0, 0, 0
    lastIrr, lastIncome, lastFTotal, lastData, field = connect_db(MonDay, SunDay)
    write_to_excel(excel_path, lastData, field)
    if lastFTotal > 0:
        if send_email(fromaddr, toaddr, ccaddr, submit, from_pw, lastIrr, lastIncome, lastFTotal, monthIrr, monthInc, monthTol, lastData, field):
            print '邮件发送成功！'
        else:
            print '邮件发送失败，请检查程序，谢谢！'
    else:
        print '执行无结果'
