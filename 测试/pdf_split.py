# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:
from aip import AipOcr
import ConfigParser
import PyPDF2
import cv2
import numpy as np

con = ConfigParser.ConfigParser()
config_path = '../../config/config.txt'
with open(config_path, 'r') as f:
    con.readfp(f)
    APP_ID = con.get('baidu_api', 'appid')
    API_KEY = con.get('baidu_api', 'api_key')
    SECRET_KEY = con.get('baidu_api', 'sec_key')

client = AipOcr(APP_ID, API_KEY, SECRET_KEY)
""" 读取图片 """


def pdf_split(file_path):
    input1 = PyPDF2.PdfFileReader(open(file_path, "rb"))
    pages = input1.getNumPages()
    pic_list = []
    for i in range(pages):
        page_i = input1.getPage(i)
        xobject = page_i['/Resources']['/XObject'].getObject()
        for obj in xobject:
            if xobject[obj]['/Subtype'] == '/Image':
                size = (xobject[obj]['/Width'], xobject[obj]['/Height'])
                data = xobject[obj]._data
                if xobject[obj]['/Filter'] == '/FlateDecode':
                    img = np.fromstring(data, np.uint8)
                    img = img.reshape(size[1], size[0])
                    img = 255 - img
                    img = img + 1
                    img = img[:, 220:-225]
                    cv2.imwrite("page" + "%05ui" % i + "obj" + obj[1:] + ".png", img)
                    pic_list.append("page" + "%05ui" % i + "obj" + obj[1:] + ".png")
                elif xobject[obj]['/Filter'] == '/DCTDecode':
                    img = np.fromstring(data, np.uint8)
                    cv2.imwrite(obj[1:] + '.jpg', img)
                elif xobject[obj]['/Filter'] == '/JPXDecode':
                    img = np.fromstring(data, np.uint8)
                    cv2.imwrite(obj[1:] + '.jp2', img)
    return True


def get_file_content(file_path):
    with open(file_path, 'rb') as fp:
        return fp.read()


if __name__ == '__main__':
    # pdf路径
    filepath = u'../../file/run1/发票.pdf'
    # 创建保存图片的文件夹
    # if os.path.exists(picpath):
    #     print("文件夹已存在，请重新创建新文件夹！")
    #     raise SystemExit
    # else:
    #     os.mkdir(picpath)
    pdf_split(filepath)
    # image = get_file_content('../../file/example.jpg')
    # """ 调用通用文字识别, 图片参数为本地图片 """
    # re = client.businessLicense(image)
    # keys = re[u'words_result'].keys()
    # for item in keys:
    #     print '%s : %s\n' % (item, re[u'words_result'][item][u'words'])
