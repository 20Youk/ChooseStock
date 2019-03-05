# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:
import ConfigParser

con = ConfigParser.ConfigParser()
alert_path = '../../config/alert_test.txt'
with open(alert_path, 'r') as g:
    con.readfp(g)
    toaddr = con.get('info', 'to_addr')
    ccaddr = con.get('info', 'cc_addr')
    sql = con.get('info', 'sql')
    submit = con.get('info', 'submit')
print toaddr
