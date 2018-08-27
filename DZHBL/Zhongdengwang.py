# -*- coding:utf-8 -*-
# Author: Youk.Lin
# 应用: 自动登录ems官网获取已签收的运单截图
from PIL import Image
import pytesseract
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
import xlrd
import time
import logging
import sys
from xlutils import copy
import xlsxwriter
import os


userCode = 'dzh001'
password = 'Dzhbl753'
# 设置文件路径
page_path = r'..\..\file\base\screenshot.png'
image_path = r'..\..\file\base\code.png'
# ems_path = r'..\..\file\image\EMS_%s.png'
log_path = '../../log/Ems_ScreenShot' + time.strftime('%Y%m%d') + '.txt'
# excel_path = '../../file/EMS.xls'
excel_path = '../../file/zhongdengwang_%s.xlsx' % time.strftime('%Y%m%d')
# code_path = '../../file/RecEms.xls'
url = 'https://www.zhongdengwang.org.cn/zhongdeng/index.shtml'

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
# options.add_argument('--headless')
# options.add_argument('--disable-gpu')
driver = webdriver.Chrome(options=options)
zero = 0
try:
    pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe'
    driver.get(url)
    driver.maximize_window()
    driver.get_screenshot_as_file(page_path)
    driver.switch_to.frame(0)
    driver.find_element_by_id('userCode').send_keys(userCode)
    passWord = driver.find_element_by_id('showpassword')
    actions = ActionChains(driver).move_to_element(passWord)
    driver.execute_script('document.getElementById("showpassword").value="%s"' % password)
    element = driver.find_element_by_id('imgId')
    left = int(element.location['x'])
    top = int(element.location['y'])
    right = int(element.location['x'] + element.size['width'])
    bottom = int(element.location['y'] + element.size['height'])
    im = Image.open(page_path)
    im = im.crop((left, top, right, bottom))
    im.save(image_path)
    # 获取验证码图片并识别
    text = pytesseract.image_to_string(im)
    print text
    driver.find_element_by_name('validateCode').send_keys(text)
    driver.find_element_by_id('login_btn').click()
    driver.close()
    print 'Done'
except (Exception, IOError), e:
    logger.error(e, exc_info=True)
    driver.close()
