import poplib
import time
import os
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
import xlrd
import email
from openpyxl import Workbook
from email import policy
import re
import datetime
def log_in():
    # 开始登录
    poplib_ = poplib.POP3_SSL('pop.qq.com')
    poplib_.user('915288938@qq.com')
    poplib_.pass_('fshsnxoporssbbja')
    mail_count = len(poplib_.list()[1])
    # 显示邮箱状态：邮件数量，占用空间
    # print('Messages: %s. Size: %s' % pop_conn.stat())
    print('邮箱登陆成功! \n收件箱共有%s封邮件,占用空间%s字节\n' % (mail_count, poplib_.stat()[1]))
    return mail_count, poplib_


def retrive(pop_conn, i):
    resp_s, lines, octets = pop_conn.retr(i)
    msg_content = b'\r\n'.join(lines).decode('utf-8')
    msg = Parser().parsestr(msg_content)
    print(msg)
if __name__ == '__main__':
    poplib_ = log_in()[1]
    retrive(poplib_,20)