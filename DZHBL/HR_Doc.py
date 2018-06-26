# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:
import docx
from win32com import client
import pythoncom
import os


def changedoc2docxbywin32(fpath, tpath):
    """把doc文件转换为docx因为用到win32com,所以仅支持windows系统
    @param fpath: 文件绝对路径,不能包含中文
    @param tpath: 文件绝对保存路径,不能包含中文"""
    pythoncom.CoInitialize()
    word = client.DispatchEx('Word.Application')  # 独立进程
    word.Visible = 0    # 不显示
    word.DisplayAlerts = 0  # 不警告
    doc = word.Documents.Open(fpath)
    doc.SaveAs(tpath, 12)  # 参数16是保存为doc,转化成docx是12
    doc.Close()
    word.Quit()
    return True


if __name__ == '__main__':
    pathFile = open('../../doc/path.cfg', mode='r')
    path0 = pathFile.read()
    srcPath = path0 + '\\source\\'
    desPath = path0 + '\\middle\\'
    docList = os.listdir(srcPath)
    for item in docList:
        allSrcPath = srcPath + item
        allDesPath = desPath + item
        changedoc2docxbywin32(allSrcPath, allDesPath)

