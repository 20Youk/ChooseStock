# -*- coding:utf8 -*-
# 应用：发送保证金对账单
import pymssql
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import make_header
import ConfigParser
import logging
import sys
import time
import xlsxwriter
import datetime
from dateutil.relativedelta import relativedelta
import os

log_path = './log/alert/alert_bzdzd_%s' % time.strftime('%Y%m%d') + '.txt'
con = ConfigParser.ConfigParser()
config_path = './config/config.txt'
alert_path = './config/alert_bzdzd.txt'
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
    ccaddr = con.get('info', 'cc_addr')
    sql = con.get('info', 'sql')
    excel_path = con.get('info', 'excel_path').decode('utf8')

# 设置日志输出
reload(sys)
sys.setdefaultencoding('gbk')
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
# 设置日志输出格式
formatter = logging.Formatter('%(asctime)s [%(levelname)s]  %(name)s : %(message)s')
# 设置日志文件路径、告警级别过滤、输出格式
fh = logging.FileHandler(log_path)
fh.setLevel(logging.INFO)
fh.setFormatter(formatter)
# 设置控制台告警级别、输出格式
ch = logging.StreamHandler()
# ch.setLevel(logging.INFO)
ch.setFormatter(formatter)
# 载入配置
logger.addHandler(fh)
logger.addHandler(ch)


def connect_db():
    # 连接数据库
    conn = pymssql.connect(hostname, user, sqlpassword, database, charset='utf8')
    cursor = conn.cursor()
    cursor.execute(sql)
    data = cursor.fetchall()
    field = cursor.description
    cursor.close()
    return data, field


def write_to_excel(filepath, data1, field1):
    all_list = []
    if data1:
        today = datetime.date.today() - relativedelta(days=+1)
        startday = today - relativedelta(months=+1)
        firstdata = data1[0][0]
        email = data1[0][-1]
        allpath = filepath + firstdata + startday.strftime('%Y%m%d').decode('utf8') + u'~' + today.strftime('%Y%m%d').decode('utf8') + u'保证金对账单.xlsx'
        all_list.append([email, allpath, firstdata + startday.strftime('%Y%m%d').decode('utf8') + u'~' + today.strftime('%Y%m%d').decode('utf8') + u'保证金对账单'])
        wb = xlsxwriter.Workbook(allpath)
        ws1 = wb.add_worksheet(u'保证金对账单')
        ws1.set_column('A:B', 42)
        newstyle = wb.add_format()
        newstyle.set_border(2)
        newstyle.set_font_size(9)
        newstyle.set_font_name(u'宋体')
        newstyle.set_align('left')      # 左对齐
        newstyle.set_align('vcenter')   # 垂直居中
        grey = wb.add_format({'border': 1, 'align': 'vcenter', 'bg_color': '#696969', 'font_size': 9, 'font_color': 'black'})
        ws1.merge_range('A1:H1', u'保证金对账单')
        ws1.write(1, 0, u'日期：' + startday.strftime('%Y-%m-%d') + '~' + today.strftime('%Y-%m-%d'))
        m = 0
        for i in [0, 1, 8, 9]:
            ws1.write(2, m, field1[i][0], grey)
            ws1.write(3, m, data1[0][i], newstyle)
            m += 1
        for i in range(0, 8):
            ws1.write(4, i, field1[i][0], grey)
            ws1.write(5, i, data1[0][i], newstyle)
        k = 6
        for i in range(1, len(data1)):
            if data1[i][0] == data1[i - 1][0] and data1[i][1] == data1[i - 1][1]:
                ws1.write_row(k, 0, data1[i][:8], newstyle)
                k += 1
            elif data1[i][0] == data1[i - 1][0] and data1[i][1] != data1[i - 1][1]:
                k += 1
                m = 0
                for j in [0, 1, 8, 9]:
                    ws1.write(k, m, field1[j][0], grey)
                    ws1.write(k + 1, m, data1[i][j], newstyle)
                    m += 1
                for j in range(0, 8):
                    ws1.write(k + 2, j, field1[j][0], grey)
                    ws1.write(k + 3, j, data1[i][j], newstyle)
                k += 4
            else:
                wb.close()
                firstdata = data1[i][0]
                email = data1[i][-1]
                allpath = filepath + firstdata + startday.strftime('%Y%m%d').decode('utf8') + u'~' + today.strftime('%Y%m%d').decode('utf8') + u'保证金对账单.xlsx'
                all_list.append([email, allpath, firstdata + startday.strftime('%Y%m%d').decode('utf8') + u'~' + today.strftime('%Y%m%d').decode('utf8') + u'保证金对账单'])
                wb = xlsxwriter.Workbook(allpath)
                ws1 = wb.add_worksheet(u'保证金对账单')
                ws1.set_column('A:B', 42)
                newstyle = wb.add_format()
                newstyle.set_border(2)
                newstyle.set_font_size(9)
                newstyle.set_font_name(u'宋体')
                newstyle.set_align('left')      # 左对齐
                newstyle.set_align('vcenter')   # 垂直居中
                grey = wb.add_format({'border': 1, 'align': 'vcenter', 'bg_color': '#696969', 'font_size': 9, 'font_color': 'black'})
                ws1.merge_range('A1:H1', u'保证金对账单')
                ws1.write(1, 0, u'日期：' + startday.strftime('%Y-%m-%d') + '~' + today.strftime('%Y-%m-%d'))
                m = 0
                for j in [0, 1, 8, 9]:
                    ws1.write(2, m, field1[j][0], grey)
                    ws1.write(3, m, data1[i][j], newstyle)
                    m += 1
                for j in range(0, 8):
                    ws1.write(4, j, field1[j][0], grey)
                    ws1.write(5, j, data1[i][j], newstyle)
                k = 6
        wb.close()
    return all_list


def send_email(from_addr, cc_addr, password, data_list):
        for item in data_list:
            try:
                onestring = item[2] + u'，请查收附件！'
                msg = MIMEMultipart()
                msg['From'] = u'<%s>' % from_addr
                msg['To'] = item[0]
                msg['Cc'] = cc_addr
                msg['Subject'] = item[2]
                # --这是文字部分--
                part = MIMEText(onestring, 'html', 'utf-8')
                msg.attach(part)
                # ---这是附件部分---
                # xlsx类型附件
                file_msg = MIMEText(open(item[1], 'rb').read(), 'base64', 'UTF-8')
                file_msg["Content-Type"] = 'application/octet-stream;name="%s"' % make_header([(os.path.basename(item[1]), 'UTF-8')]).encode('UTF-8')
                file_msg["Content-Disposition"] = 'attachment;filename= "%s"' % make_header([(os.path.basename(item[1]), 'UTF-8')]).encode('UTF-8')
                msg.attach(file_msg)
                smtp = smtplib.SMTP_SSL(mailServer, mailPort)
                smtp.login(from_addr, password)
                smtp.sendmail(from_addr, item[0].split(';') + cc_addr.split(','), msg.as_string())
                logger.info(u'邮件发送成功, 收件人:' + item[2] + u', 收件邮箱:%s' % (item[0]))
            except Exception, e:
                logger.error(e, exc_info=True)
                pass
            continue


if __name__ == "__main__":
    sqlData, sqlField = connect_db()
    send_list = write_to_excel(excel_path, sqlData, sqlField)
    if send_list:
        send_email(fromaddr, ccaddr, from_pw, send_list)
    else:
        logger.info(u'查询无结果')
    print 'DONE!!!'
