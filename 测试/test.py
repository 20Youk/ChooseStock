# -*- coding:utf-8 -*-
# Author:Lu
# 应用:
from scipy import optimize


def xnpv(rate, cashflows):
    return sum([cf/(1+rate)**((t-cashflows[0][0]).days/365.0) for (t,cf) in cashflows])


def xirr(cashflows, guess=0.1):
    try:
        return optimize.newton(lambda r: xnpv(r,cashflows),guess)
    except:
        print('Calc Wrong')

from datetime import datetime
tas = [(datetime(2010, 12, 29, 0, 0), -10000), (datetime(2012, 1, 25, 0, 0), 20), (datetime(2012, 3, 8, 0, 0), 10100)]
print xirr(tas)
print "done"
#  0.0100612640381
