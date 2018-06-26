# -*- coding:utf8 -*-
# Author:Youk.Lin
# 微信实时保存群内图片和附件
import itchat
from itchat.content import *
import time
import os
import sys
from datetime import datetime

itchat.auto_login(enableCmdQR=False, hotReload=True)


# @itchat.msg_register([PICTURE, ATTACHMENT], isGroupChat=True)
@itchat.msg_register([ATTACHMENT], isGroupChat=True)
def gchat(msg):
    if msg['FileName'][-4:] == '.gif':
        return
    timestr = datetime.utcnow().strftime('%Y%m%d_%H%M%S.%f')
    group = msg['User']['NickName']
    filepath = '../../file/' + group
    print group + '_' + msg['ActualNickName'] + u'发了图片或附件'
    if not os.path.exists(filepath):
        os.mkdir(filepath)
    msg['Text'](filepath + '/' + msg['FileName'])
    filename = os.path.splitext(msg['FileName'])[0]
    fileext = os.path.splitext(msg['FileName'])[1]
    newname = filepath + '/' + msg['ActualNickName'] + '_' + filename + '_' + timestr + fileext
    os.rename(filepath + '/' + msg['FileName'], newname)
    with open('../../log/WeChat_%s.txt' % time.strftime('%Y%m%d'), mode='a') as log:
        log.write(u'{0:s} [Info] 微信群#{1:s}#的#{2:s}#发了图片或者附件,已保存并更名为#{3:s}# \n'.format(
                datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S.%f'), group.encode('utf8'),
                msg['ActualNickName'].encode('utf8'),
                newname.encode('utf8')))
    log.close()

if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding("gbk")
    try:
        itchat.run()
    except (KeyboardInterrupt, IOError, Exception), e:
        itchat.logout()
        with open('../../log/WeChat%s.txt' % time.strftime('%Y%m%d'), mode='a') as logFile:
            logFile.write(u'{0:s} [Error] {1:s} \n'.format(time.strftime('%Y-%m-%d %H:%M:%S'), e))
        logFile.close()
