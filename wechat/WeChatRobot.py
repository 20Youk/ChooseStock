# -*-coding:utf8-*-
# Author:Youk.Lin
import itchat
from itchat.content import *
import requests
import json


@itchat.msg_register(TEXT)
def tuling_api(msg):
    key = '610e2ce2fed44d7cb527b22728a89a0b'
    info = msg['Text'].encode('utf-8')
    print ('\n' + msg['ActualNickName'] + ':' + info)
    url = 'http://www.tuling123.com/openapi/api?key=' + key + '&info=' + info
    res = requests.get(url)
    res.encoding = 'utf-8'
    jd = json.loads(res.text)
    print ('\nTuling: ' + jd['text'])
    if jd['code'] == 100000:
        itchat.send(jd['text'], msg['FromUserName'])


if __name__ == '__main__':
    itchat.auto_login(enableCmdQR=True, hotReload=True)
    itchat.run()
