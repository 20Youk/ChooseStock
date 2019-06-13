# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:


def tran_number(s):
    global d
    d = 'NULL'
    try:
        d = str('%.2f' % float(s))
        return d
    except (ValueError, TypeError):
        pass
    return d

if __name__ == '__main__':
    s = '...'
    print tran_number(s)
