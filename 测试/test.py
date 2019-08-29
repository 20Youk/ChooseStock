# -*- coding:utf-8 -*-
# Author:Lu
# 应用:
from scipy import optimize


def xnpv(rate, cashflows):
    return sum([cf / (1 + rate) ** ((t - cashflows[0][0]).days / 365.0) for (t, cf) in cashflows])


def xirr(cashflows, guess=0.1):
    try:
        return optimize.newton(lambda r: xnpv(r, cashflows), guess)
    except:
        print('Calc Wrong')


import datetime

# tas = [(datetime(2010, 12, 25, 0, 0), -10000), (datetime(2012, 1, 25, 0, 0), 20), (datetime(2012, 3, 8, 0, 0), 10100)]
tas = [(datetime.datetime(2019, 1, 23, 0, 0), -392872.81), (datetime.datetime(2019, 1, 23, 0, 0), 435878.25),
       (datetime.datetime(2019, 3, 18, 0, 0), -172111.69), (datetime.datetime(2019, 3, 29, 0, 0), 177050.35),
       (datetime.datetime(2019, 4, 18, 0, 0), -225709.86), (datetime.datetime(2019, 5, 17, 0, 0), 231430.37),
       (datetime.datetime(2019, 6, 13, 0, 0), -30000.0)]
print xirr(tas)
print "done"
#  0.0100612640381
