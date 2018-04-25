# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:日志记录
import logging
import sys


def error_func():
    b = 1 / 1
    return b

if __name__ == '__main__':
    reload(sys)
    sys.setdefaultencoding('gbk')
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger(__name__)
    # 设置日志输出格式
    formatter = logging.Formatter('%(asctime)s [%(levelname)s]  %(name)s : %(message)s')
    # 设置日志文件路径、告警级别过滤、输出格式
    fh = logging.FileHandler('../../log/logging.log')
    fh.setLevel(logging.WARN)
    fh.setFormatter(formatter)
    # 设置控制台告警级别、输出格式
    ch = logging.StreamHandler()
    # ch.setLevel(logging.INFO)
    ch.setFormatter(formatter)
    # 载入配置
    logger.addHandler(fh)
    logger.addHandler(ch)
    logger.info('test')
    try:
        error_func()
    except Exception, e:
        logger.error(e, exc_info=True)
