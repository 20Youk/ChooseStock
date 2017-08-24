# -*- coding:utf8 -*-
from Tkinter import *
# import os
#
#
# def checkprocexit(imagename):
#     try:
#         p = os.popen('tasklist /FI "IMAGENAME eq %s"' % imagename)
#         if p.read().count(imagename) > 0:
#             return imagename + ' is exist!'
#         else:
#             return imagename + ' is not exist,please check,thanks!'
#     except Exception, e:
#         print '检测程序出错，请检查，错误信息如下:\n%s' %e
# if __name__ == '__main__':
#     s = raw_input('输入进程名字[如:wmain.exe]: ')
#     print checkprocexit(s)

top = Tk()
top.title('测试标题')
top.geometry('300x400')
li = [1, 2, 3, 4, 5]
listA = Listbox(top)
for item in li:
    listA.insert(2, item)
listA.pack(side=RIGHT)
l = Label(top, text='hello', bg='pink', font=('Arial', 12), width=8, heigh=3)
l.pack(side=LEFT)


def printtext():
    t.insert(END, 'test\n')
t = Text()
t.pack()
Button(top, text='print', command=printtext).pack()
top.mainloop()
