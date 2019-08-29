# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:
import sys
import poplib
from email import parser
import email
import mysql.connector
import traceback
import datetime
import time

reload(sys)
sys.setdefaultencoding('utf8')

# 确定运行环境的encoding
__g_codeset = sys.getdefaultencoding()
if "ascii" == __g_codeset:
    __g_codeset = 'utf8'


def object2double(obj):
    if obj is None or obj == "":
        return 0
    else:
        return float(obj)
        # end if


def get_mail_index():
    file_1 = open('mailindex.txt', "r")
    lines = file_1.readlines()
    file_1.close()
    return int(lines[0])


def set_mail_index(index):
    f = open('mailindex.txt', 'w')
    f.write(index)
    f.close()


def utf8_to_mbs(s):
    return s.decode("utf-8").encode(__g_codeset)


def utf8_to_gbk(s):
    return s.decode("utf-8").encode('gb2312')


def mbs_to_utf8(s):
    return s.decode(__g_codeset).encode("utf-8")


def gbk_to_utf8(s):
    return s.decode('gb2312').encode("utf-8")


def query_quick(cu, sql, tuple_1):
    try:
        cu.execute(sql, tuple_1)
        rows = []
        for row in cu:
            rows.append(row)
        return rows
    except():
        print(traceback.format_exc())


# 获取信息
def query_rows(cu, sql):
    try:
        cu.execute(sql)
        rows = []
        for row in cu:
            rows.append(row)
        #
        return rows
    except():
        print(traceback.format_exc())


# 是否有新邮件
global hasNewMail
hasNewMail = True
# 全局已读的邮件数量
global globalMailReaded
globalMailReaded = get_mail_index() + 1


# 获取新邮件
def get_new_mail(conn_2, cur_2):
    try:
        global hasNewMail
        global globalMailReaded
        conn_2.commit()
        rows = query_rows(cur_2, 'SELECT count(*) AS message_count FROM hm_messages WHERE messageaccountid=1')
        message_count = rows[0][0]
        if hasNewMail:
            print('read mailindex.txt')
            globalMailReaded = get_mail_index() + 1
        # end if
        if message_count <= globalMailReaded:
            hasNewMail = False
            # print('Did not receive new mail,continue wait...')
            return None  # 没新邮件，直接返回
        # end if
        # 登陆邮箱
        host = '127.0.0.1'
        username = 'username@myserver.net'
        password = 'password'
        pop_conn = poplib.POP3(host)
        # print pop_conn.getwelcome()
        pop_conn.user(username)
        pop_conn.pass_(password)
        # Get messages from server:
        messages = [pop_conn.retr(i) for i in range(1, len(pop_conn.list()[1]) + 1)]
        # Concat message pieces:
        messages = ["\n".join(mssg[1]) for mssg in messages]
        # Parse message intom an email object:
        messages = [parser.Parser().parsestr(mssg) for mssg in messages]
        print("get new mail!")
        print pop_conn.stat()
        print('%s readed mail count is %d,all mail count is: %d' % (
            datetime.datetime.now().strftime("%y-%m-%d %H:%M:%S"), globalMailReaded, len(messages)))
        message = messages[globalMailReaded]
        j = 0
        for part in message.walk():
            j += 1
            filename = part.get_filename()
            contenttype = part.get_content_type()
            mycode = part.get_content_charset()
            # 保存附件
            if filename:
                pass
            elif contenttype == 'text/plain':  # or contenttype == 'text/html':
                # 保存正文
                data = part.get_payload(decode=True)
                content = str(data)
                if mycode == 'gb2312':
                    content = gbk_to_utf8(content)
                # end if
                content = content.replace(u'\u200d', '')
                set_mail_index(str(globalMailReaded))
                hasNewMail = True
                pop_conn.quit()
                return content, email.utils.parseaddr(message.get('from'))[1]
                # end if
                # end for
    except():
        print("search hmailserver fail,try again")
        return None
    finally:
        pass
        # end try


# end def

# 连接数据库

conn2 = mysql.connector.connect(user='root', password='password', host='127.0.0.1', database='hmailserver',
                                charset='gb2312')
cur2 = conn2.cursor()
# 只要收到电子邮件，就把这个事件记录在事件库中
# 现在就是循环查询邮箱，如果有新邮件就读取，并查询关键词库
while True:
    mailtuple = get_new_mail(conn2, cur2)
    if mailtuple is None:
        # print('Did not search MySQL,continue loop...')
        time.sleep(0.5)
        continue
    # end if
    (article, origin) = mailtuple
    # end while
