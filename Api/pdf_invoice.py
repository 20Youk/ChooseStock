# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:
from aip import AipOcr
import ConfigParser
import fitz
import time
import re
import os

con = ConfigParser.ConfigParser()
config_path = '../../config/config.txt'
with open(config_path, 'r') as f:
    con.readfp(f)
    APP_ID = con.get('baidu_api', 'appid')
    API_KEY = con.get('baidu_api', 'api_key')
    SECRET_KEY = con.get('baidu_api', 'sec_key')

client = AipOcr(APP_ID, API_KEY, SECRET_KEY)
""" 读取图片 """


def pdf2pic(path, pic_path):
    # 从pdf中提取图片
    t0 = time.clock()
    # 使用正则表达式来查找图片
    check_xo = r"/Type(?= */XObject)"
    check_im = r"/Subtype(?= */Image)"
    # 打开pdf
    doc = fitz.open(path)
    # 图片计数
    imgcount = 0
    len_xref = doc._getXrefLength()

    # 打印PDF的信息
    print("文件名:{}, 页数: {}, 对象: {}".format(path, len(doc), len_xref - 1))

    # 遍历每一个对象
    for i in range(1, len_xref):
        # 定义对象字符串
        text = doc._getObjectString(i)
        is_x_object = re.search(check_xo, text)
        # 使用正则表达式查看是否是图片
        is_image = re.search(check_im, text)
        # 如果不是对象也不是图片，则continue
        if not is_x_object or not is_image:
            continue
        imgcount += 1
        # 根据索引生成图像
        pix = fitz.Pixmap(doc, i)
        # 根据pdf的路径生成图片的名称
        new_name = path.replace('\\', '_') + "_img{}.png".format(imgcount)
        new_name = new_name.replace(':', '')
        # 如果pix.n<5,可以直接存为PNG
        if pix.n < 5:
            pix.writePNG(os.path.join(pic_path, new_name))
        # 否则先转换CMYK
        else:
            pix0 = fitz.Pixmap(fitz.csRGB, pix)
            pix0.writePNG(os.path.join(pic_path, new_name))
        # 释放资源
        t1 = time.clock()
        print("运行时间:{}s".format(t1 - t0))
        print("提取了{}张图片".format(imgcount))


def get_file_content(file_path):
    with open(file_path, 'rb') as fp:
        return fp.read()


if __name__ == '__main__':
    # pdf路径
    filepath = '../../file/发票.pdf'
    picpath = '../../file/run1'
    # 创建保存图片的文件夹
    # if os.path.exists(picpath):
    #     print("文件夹已存在，请重新创建新文件夹！")
    #     raise SystemExit
    # else:
    #     os.mkdir(picpath)
    pdf2pic(filepath, picpath)
    # image = get_file_content('../../file/example.jpg')
    # """ 调用通用文字识别, 图片参数为本地图片 """
    # re = client.businessLicense(image)
    # keys = re[u'words_result'].keys()
    # for item in keys:
    #     print '%s : %s\n' % (item, re[u'words_result'][item][u'words'])
