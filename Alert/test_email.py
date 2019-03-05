# -*- coding:utf8 -*-
import pymssql
import smtplib
from email.mime.text import MIMEText
import ConfigParser
import logging
import sys
import time

log_path = '../../log/alert/alert_test_%s' % time.strftime('%Y%m%d') + '.txt'
con = ConfigParser.ConfigParser()
config_path = '../../config/config.txt'
alert_path = '../../config/alert_test.txt'
with open(config_path, 'r') as f:
    con.readfp(f)
    fromaddr = con.get('info', 'from_addr')
    from_pw = con.get('info', 'from_pw')
    mailServer = con.get('info', 'mail_server')
    mailPort = con.getint('info', 'mail_port')
with open(alert_path, 'r') as g:
    con.readfp(g)
    toaddr = con.get('info', 'to_addr')
    ccaddr = con.get('info', 'cc_addr')
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
    sheetstring = '''
                    <p><strong>中华金融2019年业绩报表</strong></p>
<p><strong>日期: 2019/2/1 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    已过工作日: 21 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 时间进度:92%</strong></p>
<table border=2 cellpadding=2 cellspacing=2 width=1500pt
       style='font-size:xx-small;'>
 <tr style='height:37.5pt'>
  <td rowspan=2 height=50 style='height:37.5pt'>月份</td>
  <td colspan=4>总放款业绩</td>
  <td colspan=4>新客户放款</td>
  <td colspan=3>新客户数</td>
  <td colspan=4>老客户放款</td>
  <td colspan=4>收入</td>
  <td colspan=2>库存</td>
 </tr>
 <tr style='height:37.5pt'>
  <td>目标</td>
  <td>已放款</td>
  <td>申请中</td>
  <td>达成率</td>
  <td>目标</td>
  <td>已放款</td>
  <td>申请中</td>
  <td>达成率</td>
  <td>目标</td>
  <td>实际</td>
  <td>达成率</td>
  <td>目标</td>
  <td>已放款</td>
  <td>申请中</td>
  <td>达成率</td>
  <td>目标</td>
  <td>已放款</td>
  <td>申请中</td>
  <td>达成率</td>
  <td>目标</td>
  <td>实际</td>
 </tr>
 <tr style='height:27.0pt;word-wrap:break-word;'>
  <td>2019年1月</td>
  <td>12,981,715</td>
  <td>10,994,407</td>
  <td>420,168</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>87.9%</td>
  <td>4,391,715 </td>
  <td>608,975 </td>
  <td>0 </td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>13.9%</td>
  <td>15 </td>
  <td>2</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>13.7%</td>
  <td>8,590,000 </td>
  <td>10,385,432 </td>
  <td>420,168 </td>
  <td>125.8%</td>
  <td>374,522 </td>
  <td>437,005 </td>
  <td>8,475 </td>
  <td>118.9%</td>
  <td>17,276,715 </td>
  <td>13,703,102 </td>
 </tr>
 <tr style='height:27.0pt'>
  <td>2019年2月</td>
  <td>12,158,515 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>1,499,610 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>5 </td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>10,658,905 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>379,345 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;color:#9C0006;
  font-weight:400;text-decoration:none;text-underline-style:none;text-line-through:
  none;'>　</td>
  <td>18,776,325 </td>
  <td>　</td>
 </tr>
 <tr style='height:27.0pt'>
  <td>2019年3月</td>
  <td>16,623,625 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>4,605,945 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>15 </td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>12,017,680 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>332,473 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;color:#9C0006;
  font-weight:400;text-decoration:none;text-underline-style:none;text-line-through:
  none;'>　</td>
  <td>23,382,270 </td>
  <td>　</td>
 </tr>
 <tr style='height:27.0pt'>
  <td>2019年4月</td>
  <td>20,586,880 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>6,534,015 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>22 </td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>14,052,865 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>411,738 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;color:#9C0006;
  font-weight:400;text-decoration:none;text-underline-style:none;text-line-through:
  none;'>　</td>
  <td>29,916,285 </td>
  <td>　</td>
 </tr>
 <tr style='height:27.0pt'>
  <td>2019年5月</td>
  <td>25,264,235 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>7,498,050 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>25 </td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>17,766,185 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>505,285 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;color:#9C0006;
  font-weight:400;text-decoration:none;text-underline-style:none;text-line-through:
  none;'>　</td>
  <td>37,414,335 </td>
  <td>　</td>
 </tr>
 <tr style='height:27.0pt'>
  <td>2019年6月</td>
  <td>30,905,625 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>8,462,085 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>28 </td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>22,443,540 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>618,113 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;color:#9C0006;
  font-weight:400;text-decoration:none;text-underline-style:none;text-line-through:
  none;'>　</td>
  <td>45,876,420 </td>
  <td>　</td>
 </tr>
 <tr style='height:27.0pt'>
  <td>2019年7月</td>
  <td>37,082,590 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>9,319,005 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>31 </td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>27,763,585 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>741,652 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;color:#9C0006;
  font-weight:400;text-decoration:none;text-underline-style:none;text-line-through:
  none;'>　</td>
  <td>55,195,425 </td>
  <td>　</td>
 </tr>
 <tr style='height:27.0pt'>
  <td>2019年8月</td>
  <td>43,866,540 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>10,175,925 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>34 </td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>33,690,615 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>877,331 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;color:#9C0006;
  font-weight:400;text-decoration:none;text-underline-style:none;text-line-through:
  none;'>　</td>
  <td>65,371,350 </td>
  <td>　</td>
 </tr>
 <tr style='height:27.0pt'>
  <td>2019年9月</td>
  <td>51,543,115 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>11,354,190 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>38 </td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>40,188,925 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>1,030,862 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;color:#9C0006;
  font-weight:400;text-decoration:none;text-underline-style:none;text-line-through:
  none;'>　</td>
  <td>76,725,540 </td>
  <td>　</td>
 </tr>
 <tr style='height:27.0pt'>
  <td>2019年10月</td>
  <td>59,898,085 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>12,532,455 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>42 </td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>47,365,630 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>1,197,962 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;color:#9C0006;
  font-weight:400;text-decoration:none;text-underline-style:none;text-line-through:
  none;'>　</td>
  <td>89,257,995 </td>
  <td>　</td>
 </tr>
 <tr style='height:27.0pt'>
  <td>2019年11月</td>
  <td>70,002,600 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>14,674,755 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>49 </td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>55,327,845 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;
  color:#9C0006;font-weight:400;text-decoration:none;text-underline-style:none;
  text-line-through:none;'>　</td>
  <td>1,400,052 </td>
  <td>　</td>
  <td>　</td>
  <td style='font-size:12.0pt;color:#9C0006;
  font-weight:400;text-decoration:none;text-underline-style:none;text-line-through:
  none;'>　</td>
  <td>103,932,750 </td>
  <td>　</td>
 </tr>
 <tr style='height:27.0pt'>
  <td>2019年12月</td>
  <td>80,464,165 </td>
  <td>　</td>
  <td>　</td>
  <td>　</td>
  <td>16,067,250 </td>
  <td>　</td>
  <td>　</td>
  <td>　</td>
  <td>54 </td>
  <td>　</td>
  <td>　</td>
  <td>64,396,915 </td>
  <td>　</td>
  <td>　</td>
  <td>　</td>
  <td>1,609,283 </td>
  <td>　</td>
  <td>　</td>
  <td>　</td>
  <td>120,000,000 </td>
  <td>　</td>
 <![if supportMisalignedColumns]>
 <tr style='display:none'>
  <td width=6 style='width:5pt'></td>
  <td width=113 style='width:85pt'></td>
  <td width=131 style='width:98pt'></td>
  <td width=103 style='width:77pt'></td>
  <td width=103 style='width:77pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=109 style='width:82pt'></td>
  <td width=76 style='width:57pt'></td>
  <td width=76 style='width:57pt'></td>
  <td width=63 style='width:47pt'></td>
  <td width=69 style='width:52pt'></td>
  <td width=54 style='width:41pt'></td>
  <td width=68 style='width:51pt'></td>
  <td width=109 style='width:82pt'></td>
  <td width=95 style='width:71pt'></td>
  <td width=87 style='width:65pt'></td>
  <td width=107 style='width:80pt'></td>
  <td width=96 style='width:72pt'></td>
  <td width=86 style='width:65pt'></td>
  <td width=86 style='width:65pt'></td>
  <td width=69 style='width:52pt'></td>
  <td width=126 style='width:95pt'></td>
  <td width=105 style='width:79pt'></td>
 </tr>
 <![endif]>
</table>
                    '''
    try:
        # else:
        #     onestring = '查询无结果，请检查任务执行情况，谢谢!'
        msg = MIMEText(sheetstring, 'html', 'utf-8')
        msg['From'] = u'<%s>' % from_addr
        msg['To'] = to_addr
        msg['Cc'] = from_addr
        msg['Subject'] = u'%s' % subject
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
