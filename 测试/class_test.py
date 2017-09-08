# -*- coding:utf8 -*-
# __author__: Youk Lin
import time


class MyClass:
    """测试类"""
    def __init__(self, a):
        self.a = a

    def f(self, b):
        c = b + self.a
        print c
t1 = time.time()
MyClass(2).f(3)
print '代码执行消耗%.4f秒' % (t1 - time.time())