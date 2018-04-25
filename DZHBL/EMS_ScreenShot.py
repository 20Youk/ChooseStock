# -*- coding:utf-8 -*-
# Author: Youk.Lin
# 应用: 自动登录ems官网获取已签收的运单截图
from PIL import Image
import pytesseract
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import xlrd
import time
import logging
import sys
from xlutils import copy
import xlsxwriter
import os

# 设置文件路径
page_path = r'..\..\file\base\screenshot.png'
image_path = r'..\..\file\base\code.png'
ems_path = r'..\..\file\image\EMS_%s.png'
log_path = '../../log/Ems_ScreenShot' + time.strftime('%Y%m%d') + '.txt'
# excel_path = '../../file/EMS.xls'
excel_path = '../../file/EMS.xls'
code_path = '../../file/RecEms.xls'
url = 'http://www.11183.com.cn/ems/order/singleQuery_t'

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
        wb0 = copy.copy(rb0)
        ws0 = wb0.get_sheet(0)
    # 读取EMS单号
    rb = xlrd.open_workbook(excel_path)
    rs = rb.sheet_by_index(0)
    mailNumList = rs.col_values(0, start_rowx=1)
    codeList = rs.col_values(4, start_rowx=1)
    # 循环读取运单并截图保存
    j = len(rec_list) + 1
    for i in range(0, len(mailNumList)):
        if int(codeList[i]) == 3 and mailNumList[i] not in rec_list:
            ws0.write(j, 0, int(mailNumList[i]))
            j += 1
            mailNum = str(int(mailNumList[i]))
            driver.get(url)
            driver.get_screenshot_as_file(page_path)
            driver.find_element_by_id('mailNum').send_keys(mailNum)
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
            driver.get_screenshot_as_file(ems_path % str(mailNum))
            s = driver.find_elements_by_css_selector(css_selector='singleErrors')
            if len(s) > 0:
                logger.warning(u'运单{0:s}的验证码识别错误,请检查'.format(mailNum))
            else:
                logger.info(u'成功获取运单{0:s}信息并截图'.format(mailNum))
    driver.close()
    wb0.save(code_path) if zero == 0 else rb0.close()
    print 'Done'
except (Exception, IOError), e:
    logger.error(e, exc_info=True)
    driver.close()
