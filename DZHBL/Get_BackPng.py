# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:
import os
import urllib
import xlrd
import time


def save_img(img_url, file_name, file_path='../../file/image/%s' % time.strftime('%Y%m%d')):
    # 保存图片到磁盘文件夹 file_path中，默认为当前脚本运行目录下的 book\img文件夹
    try:
        if not os.path.exists(file_path):
            print '文件夹', file_path, '不存在，重新建立'
            os.makedirs(file_path)
        # 获得图片后缀
        file_suffix = os.path.splitext(img_url)[1]
        # 拼接图片名（包含路径）
        filename = '{}{}{}{}'.format(file_path, os.sep, file_name.encode('utf8'), file_suffix.encode('utf8'))
        # 下载图片，并保存到文件夹中
        urllib.urlretrieve(img_url, filename=filename)
    except IOError as e:
        print '文件操作失败', e
    except Exception as e:
        print '错误 ：', e


if __name__ == '__main__':
    rb = xlrd.open_workbook('../../file/BData.xlsx')
    rs = rb.sheet_by_index(0)
    codeList = rs.col_values(0, start_rowx=1)
    hotelList = rs.col_values(1, start_rowx=1)
    urlList = rs.col_values(4, start_rowx=1)
    for i in range(0, len(codeList)):
        save_img(urlList[i], str(int(codeList[i])))
        print '完成第%d个图片下载...' % (i + 1)
    print "DONE!!!"
