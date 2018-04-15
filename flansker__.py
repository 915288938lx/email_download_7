import poplib
import time
import os
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
import xlrd
from openpyxl import Workbook
import re
import datetime
from flanker import mime
def log_in():
    # 开始登录
    pop_conn = poplib.POP3_SSL('pop.qq.com')
    pop_conn.user('915288938@qq.com')
    pop_conn.pass_('fshsnxoporssbbja')
    mail_count = len(pop_conn.list()[1])
    # 显示邮箱状态：邮件数量，占用空间
    # print('Messages: %s. Size: %s' % pop_conn.stat())
    print('邮箱登陆成功! \n收件箱共有%s封邮件,占用空间%s字节\n' % (mail_count, pop_conn.stat()[1]))
    return mail_count, pop_conn
def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value
def retrive(pop_conn, i):
    resp_s, lines, octets = pop_conn.retr(i)
    # 解码原始邮件
    print(lines)
    msg = mime.from_string(lines)
    print(type(msg))
    print(msg)
    subject_ = msg.get('Subject')
    subject = decode_str(subject_)
    sender = msg.get('From')
    date1 = time.strptime(msg.get("Date")[0:24], '%a, %d %b %Y %H:%M:%S')  # 格式化收件时间
    send_date_str = time.strftime("%Y%m%d", date1)
    hdr, addr = parseaddr(sender)
    address = u'%s' % (addr)  # 发件人邮箱
    print('\n'+subject)
    return msg, address, subject, send_date_str

if __name__ == '__main__':
    pop_conn = log_in()[1]
    retrive(pop_conn,18)