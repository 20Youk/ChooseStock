# -*- coding:utf-8 -*-
# Author: Youk.Lin
# 应用: 调用微信将对应的工资表发送给员工
import itchat
import time
import xlrd
import logging
import sys
import os

lastPath = os.path.abspath('..')
reload(sys)
sys.setdefaultencoding('gbk')
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)
# 设置日志输出格式
formatter = logging.Formatter('%(asctime)s [%(levelname)s]  %(name)s : %(message)s')
# 设置日志文件路径、告警级别过滤、输出格式
fh = logging.FileHandler(lastPath + '\\log\\logging_%s.log' % time.strftime('%Y%m%d'))
fh.setLevel(logging.WARN)
fh.setFormatter(formatter)
# 设置控制台告警级别、输出格式
ch = logging.StreamHandler()
# ch.setLevel(logging.INFO)
ch.setFormatter(formatter)
# 载入配置
logger.addHandler(fh)
logger.addHandler(ch)

if __name__ == '__main__':
    try:
        typeDict = {type(u'\u4e2d\u6587'): '%s', type(0.01): '%.2f'}
        filePath = lastPath + '\\file\\wechat_file.xls'
        rb = xlrd.open_workbook(filePath)
        rs = rb.sheet_by_index(0)
        field = rs.row_values(0, 0)
        firstRow = rs.row_values(1, 0)
        msgList = []
        for i in range(0, len(firstRow)):
            msgList.append(field[i] + ':' + typeDict[type(firstRow[i])])
        msg = '\n'.join(msgList)
        addressDict = {}
        for i in range(0, len(field)):
            addressDict[field[i]] = rs.col_values(i, start_rowx=1)
        itchat.auto_login(hotReload=True, enableCmdQR=True)
        time.sleep(2)
        for i in range(0, len(addressDict[field[0]])):
            data = []
            for item in field:
                data.append(addressDict[item][i])
            msg_1 = msg % tuple(data)
            weChat_user = itchat.search_friends(name=addressDict[field[2]][i])[0]['UserName']
            itchat.send(msg=msg_1, toUserName=weChat_user)
            logging.info(u'成功发送{0:s}的记录.....'.format(addressDict[field[2]][i]))
        print 'Done!!!'
    except (Exception, IOError), e:
        logger.error(e, exc_info=True)
