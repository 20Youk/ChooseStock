# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:将一串字符中的文字和数字分离


def exp_num(s):
    global s1, s2
    for i in range(0, len(s)):
        if s[i].isdigit():
            s1 = s[:i].strip()
            s2 = s[i:].strip()
            break
    return s1, s2

ss = u'中国银行股份有限公司深圳蛇口网谷支行 751057939123'
ss1, ss2 = exp_num(ss)
print ss1
print ss2
