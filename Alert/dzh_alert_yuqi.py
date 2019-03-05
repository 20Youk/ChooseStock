# -*- coding:utf8 -*-
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

log_path = './log/alert/alert_yuqi_%s' % time.strftime('%Y%m%d') + '.txt'
con = ConfigParser.ConfigParser()
config_path = './config/config.txt'
alert_path = './config/alert_yuqi.txt'
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
    wb = xlsxwriter.Workbook(filepath)
    ws1 = wb.add_worksheet(u'每日逾期统计')
    newstyle = wb.add_format()
    newstyle.set_border(2)
    newstyle.set_font_size(9)
    newstyle.set_font_name(u'宋体')
    newstyle.set_align('left')      # 左对齐
    newstyle.set_align('vcenter')   # 垂直居中
    length1 = 10
    length2 = 10
    for i in range(0, len(field1)):
        ws1.write(0, i, field1[i][0], newstyle)
    for i in range(0, len(data1)):
        for j in range(0, len(field1)):
            ws1.write(i + 1, j, data1[i][j], newstyle)
        length1 = max(length1, len(data1[i][2]))
        length2 = max(length2, len(data1[i][1]))
    ws1.set_column('C:C', length1)
    ws1.set_column('B:B', length2)
    ws1.set_column('D:J', 20)
    wb.close()
    return


def send_email(from_addr, to_addr, cc_addr, subject, password, data1, field1):
    try:
        sheetstring = '''<table width="1000" border="2" bordercolor="black" cellspacing="2">
                        <tr>
                        '''
        titles = ''
        for item in field1:
            titles = titles + '<td><strong>' + str(item[0].encode('utf8')) + '</strong></td>'
        sheetstring = sheetstring + titles + '</tr>'
        onestring = ''
        if data1:
            for i in range(0, len(data1)):
                onestring += '<tr>'
                for j in range(0, len(field1)):
                    onedata = data1[i][j]
                    if type(u'a') == type(onedata):
                        onedata = str(onedata.encode('utf8'))
                    else:
                        onedata = str(onedata)
                    onestring = '''{0} <td>{1}</td>'''.format(onestring, onedata)
                onestring += '</tr>'
            onestring = sheetstring + onestring + '''</table>'''
        else:
            onestring = '查询无结果，谢谢!'
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
    except Exception, e:
        logger.error(e, exc_info=True)
        return False


if __name__ == "__main__":
    sqlData, sqlField = connect_db()
    write_to_excel(excel_path, sqlData, sqlField)
    if send_email(fromaddr, toaddr, ccaddr, submit, from_pw, sqlData, sqlField):
        print '邮件发送成功！'
    else:
        print '邮件发送失败，请检查程序，谢谢！'
