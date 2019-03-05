# -*- coding:utf8 -*-
import pymssql
import smtplib
from email.mime.text import MIMEText
import ConfigParser
import logging
import sys
import time

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


def send_email(from_addr, to_addr, subject, password):
    # 连接数据库
    conn = pymssql.connect(hostname, user, sqlpassword, database, charset='utf8')
    cursor = conn.cursor()
    # sql1 = '''exec yuqi 5'''
    sheetstring = '''<table width="500" border="2" bordercolor="black" cellspacing="2">
                    <tr>
                    '''
    titles = ''
    cursor.execute(sql)
    data1 = cursor.fetchall()
    field = cursor.description
    for item in field:
        titles = titles + '<td><strong>' + str(item[0].encode('utf8')) + '</strong></td>'
    sheetstring = sheetstring + titles + '</tr>'
    cursor.close()
    onestring = ''
    if data1:
        try:
            for i in range(0, len(data1)):
                onestring += '<tr>'
                for j in range(0, len(field)):
                    onedata = data1[i][j]
                    if type(u'a') == type(onedata):
                        onedata = str(onedata.encode('utf8'))
                    else:
                        onedata = str(onedata)
                    onestring = onestring + ''' <td>''' + onedata + '''</td>'''
                onestring += '</tr>'
            onestring = sheetstring + onestring + '''</table>'''
            # else:
            #     onestring = '查询无结果，请检查任务执行情况，谢谢!'
            msg = MIMEText(onestring, 'html', 'utf-8')
            msg['From'] = u'<%s>' % from_addr
            msg['To'] = to_addr
            msg['Cc'] = from_addr
            msg['Subject'] = subject
            smtp = smtplib.SMTP_SSL(mailServer, mailPort)
            smtp.login(from_addr, password)
            smtp.sendmail(from_addr, to_addr, msg.as_string())
            logger.info(u'邮件发送成功....')
            return True
        except Exception, e:
            logger.error(e, exc_info=True)
            return False


if __name__ == "__main__":
    if send_email(fromaddr, toaddr, submit, from_pw):
        print '邮件发送成功！'
    else:
        print '邮件发送失败，请检查程序，谢谢！'
