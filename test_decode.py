import poplib
import time
import os
import base64
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
import xlrd
from openpyxl import Workbook
import re
import datetime
def log_in():
    # 开始登录
    pop_conn = poplib.POP3_SSL('pop.qq.com')
    pop_conn.user('915288938@qq.com')
    pop_conn.pass_('usrrbacstblibfjd')
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
    msg_content = b'\r\n'.join(lines).decode('utf-8',errors='replace')
    msg = Parser().parsestr(msg_content)
    # 主题有效性验证
    if msg.get('Subject'):
        subject = decode_str(msg.get('Subject'))
    else:
        subject = "主题不存在..."
    # 解析发件人邮箱地址
    hdr, addr = parseaddr(msg.get('From'))
    address = u'%s' % (addr)  # 发件人邮箱
    # 解析邮件日期
    date_ = msg.get('Date')
    return msg, address, subject, date_

if __name__ == '__main__':
    mail_count, pop_conn = log_in()
    # cao = retrive(pop_conn,1338)[2]
    # print(cao)
    # b = str(cao)
    # print(type(cao))
    # print(b)

    for i in range(203,1,-1):
        _, adress, subject, date = retrive(pop_conn,i)
        print(str(i)+adress+'  '+ str(subject) + '  ')