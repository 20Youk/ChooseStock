# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:
from PIL import Image
import pytesseract
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
import logging
import sys
import xlsxwriter
import os
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


def screen_shot():
    # 设置文件路径
    global wb0
    page_path = r'..\..\file\base\screenshot.png'
    image_path = r'..\..\file\base\code.png'
    ems_path = r'..\..\file\image\EMS_%s.png'
    excel_path = '../../file/EMS.xls'
    code_path = '../../file/RecEms.xls'
    url = 'http://www.11183.com.cn/ems/order/singleQuery_t'

    # 设置WebDriver参数
    options = Options()
    options.add_argument('--headless')
    # options.add_argument('--disable-gpu')
    driver = webdriver.Chrome(options=options)
    zero = 0
    try:
        pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe'
        # 读取已获取过的EMS单号
        if not os.path.exists(code_path):
            zero = 1
            rb0 = xlsxwriter.Workbook(code_path)
            ws0 = rb0.add_worksheet('Sheet1')
            ws0.write(0, 0, 'EMS_code')
            rec_list = []
        else:
            rb0 = xlrd.open_workbook(code_path)
            rs0 = rb0.sheet_by_index(0)
            rec_list = rs0.col_values(0, start_rowx=1)
            wb0 = copy(rb0)
            ws0 = wb0.get_sheet(0)
        # 读取EMS单号
        rb0 = xlrd.open_workbook(excel_path)
        rs = rb0.sheet_by_index(0)
        mailnum_list = rs.col_values(0, start_rowx=1)
        codelist = rs.col_values(4, start_rowx=1)
        # 循环读取运单并截图保存
        jj = len(rec_list) + 1
        for ii in range(0, len(mailnum_list)):
            if int(codelist[ii]) == 3 and mailnum_list[ii] not in rec_list:
                ws0.write(jj, 0, int(mailnum_list[ii]))
                jj += 1
                mailnum = str(int(mailnum_list[ii]))
                driver.get(url)
                driver.get_screenshot_as_file(page_path)
                driver.find_element_by_id('mailNum').send_keys(mailnum)
                element = driver.find_element_by_id('checkCode')
                left = int(element.location['x'])
                top = int(element.location['y'])
                right = int(element.location['x'] + element.size['width'])
                bottom = int(element.location['y'] + element.size['height'])
                im = Image.open(page_path)
                im = im.crop((left, top, right, bottom))
                im.save(image_path)
                # 获取验证码图片并识别
                text = pytesseract.image_to_string(im)
                driver.find_element_by_name('checkCode').send_keys(text)
                driver.find_element_by_css_selector('[class="submitbtn singlebtn"]').click()
                driver.get_screenshot_as_file(ems_path % str(mailnum))
                s = driver.find_elements_by_css_selector(css_selector='singleErrors')
                if len(s) > 0:
                    logger.warning(u'运单{0:s}的验证码识别错误,请检查'.format(mailnum))
                else:
                    logger.info(u'成功获取运单{0:s}信息并截图'.format(mailnum))
        driver.close()
        wb0.save(code_path) if zero == 0 else rb0.close()
        print 'Screen_Done'
    except (Exception, IOError), E:
        logger.error(E, exc_info=True)
        driver.close()


def encrypt(origin_data, appkey):
    """数据内容签名：把(请求内容(未编码)+AppKey)进行MD5加密，然后Base64编码
    :param appkey:
    :param origin_data:
    """
    m = hashlib.md5()
    m.update((origin_data+appkey).encode("utf8"))
    encodestr = m.hexdigest()
    base64_text = base64.b64encode(encodestr.encode(encoding='utf-8'))
    return base64_text


def sendpost(url, datas):
    """发送post请求
    :param url:
    :param datas:
    """
    postdata = urllib.urlencode(datas).encode('utf-8')
    header = {
        "Accept": "application/x-www-form-urlencoded;charset=utf-8",
        "Accept-Encoding": "utf-8"
    }
    req = urllib2.Request(url, postdata, header)
    get_data = (urllib2.urlopen(req).read().decode('utf-8'))
    return get_data


def get_traces(logistic_code, shipper_code, appid, appkey, url):
    """查询接口支持按照运单号查询(单个查询)
    :param url:
    :param appkey:
    :param appid:
    :param shipper_code:
    :param logistic_code:
    """
    data1 = {'LogisticCode': logistic_code, 'ShipperCode': shipper_code}
    d1 = json.dumps(data1, sort_keys=True)
    requestdata = encrypt(d1, appkey)
    post_data = {'RequestData': d1, 'EBusinessID': appid, 'RequestType': '1002', 'DataType': '2',
                 'DataSign': requestdata.decode()}
    json_data = sendpost(url, post_data)
    sort_data = json.loads(json_data)
    return sort_data


def recognise(expresscode):
    """输出数据
    :param expresscode:
    """
    url = 'http://api.kdniao.cc/Ebusiness/EbusinessOrderHandle.aspx'
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
            str_state = [u"已签收", trace_data['Traces'][-1]['AcceptStation'].split(u'\uff1a')[1],
                         trace_data['Traces'][-1]['AcceptTime'], 3]

        print(u"单号%s目前的状态: %s： " % (expresscode, str_state[0]))
        print str_state
        return str_state


if __name__ == '__main__':
    # code = raw_input("请输入快递单号(Esc退出)：")
    # code = code.strip()
    # 设置日志输出
    log_path = '../../log/Get_Ems' + time.strftime('%Y%m%d') + '.txt'
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

    today = datetime.datetime.now().strftime('%Y%m%d')
    filePath = '../../file/EMS.xls'
    try:
        rb = xlrd.open_workbook(filePath)
        rs1 = rb.sheet_by_index(0)
        num_list = rs1.col_values(0, start_rowx=1)
        code_list = rs1.col_values(4, start_rowx=1)
        wb = copy(rb)
        ws = wb.get_sheet(0)
        for i in range(0, len(num_list)):
            if code_list[i] == '':
                ems_state = recognise(str(int(float(num_list[i]))))
                for j in range(0, len(ems_state)):
                    ws.write(i + 1, j + 1, ems_state[j])
                logger.info(u'成功获取运单{0:d}状态...\n'.format(int(float(num_list[i]))))
            elif int(code_list[i]) != 3:
                ems_state = recognise(str(int(float(num_list[i]))))
                for j in range(0, len(ems_state)):
                    ws.write(i + 1, j + 1, ems_state[j])
                logger.info(u'成功获取运单{0:d}状态...\n'.format(int(float(num_list[i]))))
        wb.save(filePath)
        print 'EMS_API,Done!!!'
        screen_shot()
        print 'ALLDone'
    except (urllib2.URLError, Exception, IOError), e:
        print u'程序运行错误，请检查！'
        logger.warning(u'{0:s} : {1:s}\n'.format(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), e))
