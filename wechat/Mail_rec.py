# -*- coding:utf8 -*-
# Author:Youk.Lin
import requests
import re
import rsa
import sys
import base64
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
# import argparse
reload(sys)
sys.setdefaultencoding('utf8')


# 打印部门人员信息
def print_tree(id, department_infos, level, _staff_infors, _f):
    prefix = '----' * level
    _text = prefix + department_infos[id]['name'] + prefix
    print _text
    _f.write(_text + '\n')
    for key, value in department_infos.items():
        if value['pid'] == id:
            print_tree(
                value['id'], department_infos, level + 1, _staff_infors, _f)
    prefix = '    ' * level
    for _staff in _staff_infors:
        if _staff['pid'] == id:
            _text = prefix + _staff['name'] + '  ' + _staff['alias']
            print _text
            _f.write(_text + '\n')


# 提取RSA算法的公钥
def get_public_key(_content):
    _regexp = r'var\s*PublicKey\s*=\s*"(\w+?)";'
    _results = re.findall(_regexp, _content)
    if _results:
        return _results[0]


# 获取ts参数
def get_ts(_content):
    _regexp = r'PublicTs\s*=\s*"([0-9]+)"'
    _results = re.findall(_regexp, _content)
    if _results:
        return _results[0]


# 计算p参数
def get_p(_public_key, _password, _ts):
    _public_key = rsa.PublicKey(int(_public_key, 16), 65537)
    res_tmp = rsa.encrypt(
        '{password}\n{ts}\n'.format(password=_password, ts=_ts), _public_key)
    return base64.b64encode(res_tmp)


def msg():
    return 'python get_tencent_exmail_contacts.py -u name@domain.com -p password'

if __name__ == "__main__":
    # description = u"获取腾讯企业邮箱通讯录"
    # parser = argparse.ArgumentParser(description=description, usage=msg())
    # parser.add_argument(
    #     "-u", "--email", required=True, dest="email", help=u"邮箱名")
    # parser.add_argument(
    #     "-p", "--password", required=True, dest="password", help=u"邮箱密码")
    # parser.add_argument(
    #     "-l", "--limit", required=False, dest="limit", default=10000, help=u"通讯录条数")
    # parser.add_argument(
    #     "-e", "--efile", required=False, dest="emailfile", default="emails.txt", help=u"邮箱保存文件")
    # parser.add_argument(
    #     "-d", "--dfile", required=False, dest="departfile", default="departments.txt", help=u"部门信息保存文件")
    # args = parser.parse_args()
    # email = args.email
    # password = args.password
    # limit = args.limit
    # emailfile = args.emailfile
    # departfile = args.departfile
    email = 'junyou.lin@gcfactoring.cn'
    password = 'Vz8youk619'
    limit = 10
    emailfile = '../../file/emails.txt'
    departfile = '../../file/departments.txt'
    session = requests.Session()

    headers = {'Connection': 'keep-alive',
               'Cache-Control': 'max-age=0',
               'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
               'Upgrade-Insecure-Requests': '1',
               'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_2) '
                             'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.116 Safari/537.36',
               'DNT': '1',
               'Accept-Encoding': 'gzip, deflate, sdch',
               'Accept-Language': 'zh-CN,zh;q=0.8,en-US;q=0.6,en;q=0.4',
               }
    resp = session.get('https://exmail.qq.com/login', headers=headers)
    content = resp.content

    public_key = get_public_key(content)

    ts = get_ts(content)

    p = get_p(public_key, password, ts)

    # print ts
    # print public_key
    # print p

    uin = email.split('@')[0]
    domain = email.split('@')[1]
    # print uin
    # print domain
    post_data = {'sid': '', 'firstlogin': False, 'domain': domain, 'aliastype': 'other', 'errtemplate': 'dm_loginpage',
                 'first_step': '', 'buy_amount': '', 'year': '', 'company_name': '', 'is_get_dp_coupon': '',
                 'starttime': int(time.time() * 1000), 'redirecturl': '', 'f': 'biz', 'uin': uin, 'p': p,
                 'delegate_url': '', 'ts': ts, 'from': '', 'ppp': '', 'chg': 0, 'loginentry': 3, 's': '',
                 'dmtype': 'bizmail', 'fun': '', 'inputuin': email, 'verifycode': ''}

    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    url = 'https://exmail.qq.com/cgi-bin/login'
    resp = session.post(url, headers=headers, data=post_data)
    regexp = r'sid=(.*?)"'
    sid = re.findall(regexp, resp.content)[0]
    url = 'https://exmail.qq.com/cgi-bin/frame_html?t=frame_html&sid={sid},' \
          '7&url=%2Fcgi-bin%2Fmail_list%3Ffolderid%3D3%26page%3D0%26topmails%3D0'
    # resp = session.get(url.format(sid=sid))
    # text = resp.text
    # 获取发件箱列表
    options = Options()
    options.add_argument('--headless')
    options.add_argument('disable-gpu')
    driver = webdriver.Chrome(options=options)
    driver.get(url.format(sid=sid))
    result = driver.find_elements_by_xpath('//*[@class="i"]/tbody/tr')
    print result
    # url = 'http://exmail.qq.com/cgi-bin/laddr_biz?action=show_party_list&sid={sid}&t=contact&view=biz'
    # resp = session.get(url.format(sid=sid))
    #
    # text = resp.text
    # regexp = r'{id:"(\S*?)", pid:"(\S*?)", name:"(\S*?)", order:"(\S*?)"}'
    # results = re.findall(regexp, text)
    # department_ids = []
    # department_infor = dict()
    # root_department = None
    # for item in results:
    #     department_ids.append(item[0])
    #     department = dict(id=item[0], pid=item[1], name=item[2], order=item[3])
    #     department_infor[item[0]] = department
    #     if item[1] == 0 or item[1] == '0':
    #         root_department = department
    #
    # regexp = r'{uin:"(\S*?)",pid:"(\S*?)",name:"(\S*?)",alias:"(\S*?)",sex:"(\S*?)",pos:"(\S*?)",tel:"(\S*?)",' \
    #          r'birth:"(\S*?)",slave_alias:"(\S*?)",department:"(\S*?)",mobile:"(\S*?)"}'
    #
    # all_emails = []
    # staff_infors = []
    # for department_id in department_ids:
    #     url = 'http://exmail.qq.com/cgi-bin/laddr_biz?' \
    #           't=memtree&limit={limit}&partyid={partyid}&action=show_party&sid={sid}'
    #     resp = session.get(url.format(limit=limit, sid=sid, partyid=department_id))
    #     text = resp.text
    #     results = re.findall(regexp, text)
    #
    #     for item in results:
    #         all_emails.append(item[3])
    #         print item[3]
    #         staff = dict(uin=item[0], pid=item[1], name=item[2], alias=item[3], sex=item[4], pos=item[
    #                      5], tel=item[6], birth=item[7], slave_alias=item[8], department=item[9], mobile=item[10])
    #         staff_infors.append(staff)
    #
    # with open(emailfile, 'w') as f:
    #     for item in all_emails:
    #         f.write(item + '\n')
    #
    # with open(departfile, 'w') as f:
    #     print_tree(root_department['id'], department_infor, 0, staff_infors, f)
    #
    # print("total email count: %i" % len(all_emails))
    # print("total department count: %i" % len(department_ids))
