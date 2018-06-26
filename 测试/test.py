# -*- coding:utf-8 -*-
# Author:Lu
# 应用:
import win32api
import win32con
from tkinter import *

root = Tk()
root.title("你好吗")
root.geometry('300x300')                 # 是x 不是*


l1 = Label(root, text="请输入你的名字：")
l1.pack()  # 这里的side可以赋值为LEFT RTGHT TOP  BOTTOM
xls_text = StringVar()
xls = Entry(root, textvariable=xls_text)
xls_text.set(" ")
xls.pack()


def on_click():
    x = xls_text.get()
    string = u"\u963f\u868c:\n\n    \u6211\u89c9\u5f97\u4f60\u957f\u5f97\u5f88\u597d\u770b\n\n                      \u5c0f\u9e7f"
    win32api.MessageBox(0, string, u'这是真的', win32con.MB_OK)

Button(root, text="press", command=on_click).pack()

root.mainloop()
