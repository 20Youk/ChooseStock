# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:通过接口查询快递进展
import requests
import json
import urllib2

def check_api(com, num):
    # url = 'http://q.kdpt.net/api'
    # data = {'id': 'XDB2gzsjbsss11ow124aNo0I_1350756306', 'com': com, 'nu': num, 'show': 'json'}
    # json_data = json.dumps(data, sort_keys=True)
    # # header = {
    # #     "Accept": "application/x-www-form-urlencoded;charset=utf-8",
    # #     "Accept-Encoding": "utf-8"
    # # }
    # req = urllib2.Request(url, json_data)
    # resp = (urllib2.urlopen(req).read().decode('utf-8'))
    # print resp.text
    url = 'http://q.kdpt.net/api?id=XDB2gzsjbsss11ow124aNo0I_1350756306&com={company}&nu={num}'
    resp = json.loads(requests.get(url.format(company=com, num=num)).text)
    print url.format(company=com, num=num)
    print resp


if __name__ == '__main__':
    company = 'auto'
    emsNum = '1072076671925'
    # emsNum = '488719320856'
    check_api(company, emsNum)
