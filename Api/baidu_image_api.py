# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:
from aip import AipOcr
import ConfigParser

con = ConfigParser.ConfigParser()
config_path = '../../config/config.txt'
with open(config_path, 'r') as f:
    con.readfp(f)
    APP_ID = con.get('baidu_api', 'appid')
    API_KEY = con.get('baidu_api', 'api_key')
    SECRET_KEY = con.get('baidu_api', 'sec_key')

client = AipOcr(APP_ID, API_KEY, SECRET_KEY)
""" 读取图片 """


def get_file_content(file_path):
    with open(file_path, 'rb') as fp:
        return fp.read()


image = get_file_content('../../file/example3.jpg')
""" 调用通用文字识别, 图片参数为本地图片 """
re = client.businessLicense(image)
keys = re[u'words_result'].keys()
for item in keys:
    print '%s : %s\n' % (item, re[u'words_result'][item][u'words'])
