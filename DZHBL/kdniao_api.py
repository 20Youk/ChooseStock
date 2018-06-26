# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:
import json
import urllib
import hashlib
import base64
import urllib2
import xlrd
from xlutils.copy import copy
import datetime

# 此处为快递鸟官网申请的帐号和密码
APP_id = "1333710"
APP_key = "00eb4f8d-9ef8-4f85-b563-4c83bfe9b1bd"


def encrypt(origin_data, appkey):
    """数据内容签名：把(请求内容(未编码)+AppKey)进行MD5加密，然后Base64编码"""
    m = hashlib.md5()
    m.update((origin_data+appkey).encode("utf8"))
    encodestr = m.hexdigest()
    base64_text = base64.b64encode(encodestr.encode(encoding='utf-8'))
    return base64_text


def sendpost(url, datas):
    """发送post请求"""
    postdata = urllib.urlencode(datas).encode('utf-8')
    header = {
        "Accept": "application/x-www-form-urlencoded;charset=utf-8",
        "Accept-Encoding": "utf-8"
    }
    req = urllib2.Request(url, postdata, header)
    get_data = (urllib2.urlopen(req).read().decode('utf-8'))
    return get_data


def get_company(logistic_code, appid, appkey, url):
    """获取对应快递单号的快递公司代码和名称"""
    data1 = {'LogisticCode': logistic_code}
    d1 = json.dumps(data1, sort_keys=True)
    requestdata = encrypt(d1, appkey)
    post_data = {
        'RequestData': d1,
        'EBusinessID': appid,
        'RequestType': '2002',
        'DataType': '2',
        'DataSign': requestdata.decode()}
    json_data = sendpost(url, post_data)
    sort_data = json.loads(json_data)
    return sort_data


def get_traces(logistic_code, shipper_code, appid, appkey, url):
    """查询接口支持按照运单号查询(单个查询)"""
    data1 = {'LogisticCode': logistic_code, 'ShipperCode': shipper_code}
    d1 = json.dumps(data1, sort_keys=True)
    requestdata = encrypt(d1, appkey)
    post_data = {'RequestData': d1, 'EBusinessID': appid, 'RequestType': '1002', 'DataType': '2',
                 'DataSign': requestdata.decode()}
    json_data = sendpost(url, post_data)
    sort_data = json.loads(json_data)
    return sort_data


def recognise(expresscode):
    """输出数据"""
    url = 'http://api.kdniao.cc/Ebusiness/EbusinessOrderHandle.aspx'
    # data = get_company(expresscode, APP_id, APP_key, url)
    # if not any(data['Shippers']):
    #     print("未查到该快递信息,请检查快递单号是否有误！")
    # else:
    #     print("已查到该", str(data['Shippers'][0]['ShipperName'])+"("+
    #           str(data['Shippers'][0]['ShipperCode'])+")", expresscode)
    #     trace_data = get_traces(expresscode, data['Shippers'][0]['ShipperCode'], APP_id, APP_key, url)
    #     if trace_data['Success'] == "false" or not any(trace_data['Traces']):
    #         print("未查询到该快递物流轨迹！")
    #     else:
    #         str_state = "问题件"
    #         if trace_data['State'] == '2':
    #             str_state = "在途中"
    #         if trace_data['State'] == '3':
    #             str_state = "已签收"
    #         print("目前状态： "+str_state)
    #         trace_data = trace_data['Traces']
    #         item_no = 1
    #         for item in trace_data:
    #             print(str(item_no)+":", item['AcceptTime'], item['AcceptStation'])
    #             item_no += 1
    #         print("\n")
    # return
    trace_data = get_traces(expresscode, 'EMS', APP_id, APP_key, url)
    if trace_data['Success'] == "false" or not any(trace_data['Traces']) or trace_data['State'] == '0':
        print("未查询到该快递物流轨迹！")
    else:
        str_state = [u"问题件", "", "", 0]
        if trace_data['State'] == '1':
            str_state = [u'已揽收', "", "", 1]
        if trace_data[u'State'] == '2':
            str_state = [u"在途中", "", "", 2]
        if trace_data['State'] == '3':
            str_state = [u"已签收", trace_data['Traces'][-1]['AcceptStation'].split(u'\uff1a')[1], trace_data['Traces'][-1]['AcceptTime'], 3]

        print(u"单号%s目前的状态: %s： "%(expresscode, str_state[0]))
        print str_state
        return str_state


if __name__ == '__main__':
    # code = raw_input("请输入快递单号(Esc退出)：")
    # code = code.strip()
    today = datetime.datetime.now().strftime('%Y%m%d')
    filePath = '../../file/EMS.xls'
    logPath = '../../file/log_%s.txt' % today
    try:
        logFile = open(logPath, mode='a', buffering=1)
        rb = xlrd.open_workbook(filePath)
        rs = rb.sheet_by_index(0)
        num_list = rs.col_values(0, start_rowx=1)
        code_list = rs.col_values(4, start_rowx=1)
        wb = copy(rb)
        ws = wb.get_sheet(0)
        for i in range(0, len(num_list)):
            if code_list[i] == '':
                ems_state = recognise(str(int(float(num_list[i]))))
                for j in range(0, len(ems_state)):
                    ws.write(i + 1, j + 1, ems_state[j])
                logFile.write('{0:s} : 成功获取运单{1:d}状态...\n'.format(
                        datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), int(float(num_list[i]))))
            elif int(code_list[i]) != 3:
                ems_state = recognise(str(int(float(num_list[i]))))
                for j in range(0, len(ems_state)):
                    ws.write(i + 1, j + 1, ems_state[j])
                logFile.write('{0:s} : 成功获取运单{1:d}状态...\n'.format(
                        datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), int(float(num_list[i]))))
        wb.save(filePath)
        logFile.close()
        print 'Done'
    except (urllib2.URLError, Exception, IOError), e:
        print u'程序运行错误，请检查！'
        logFile = open(logPath, mode='a', buffering=1)
        logFile.write(u'{0:s} : {1:s}\n'.format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), e))
        logFile.close()
