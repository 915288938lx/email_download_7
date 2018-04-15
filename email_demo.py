#! /usr/bin/env python
# coding=utf-8
import sys
import time
import poplib
import smtplib
from datetime import datetime
import email
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr
from email.mime.text import MIMEText
import re
import pyodbc


def remove_values_from_list(the_list, val):
    while val in the_list:
        the_list.remove(val)


# 邮件发送函数
def send_mail(mail_host, mail_user, mail_pass, mail_postfix, to_list, sub, content):
    me = "" + "" + mail_user + "@" + mail_postfix + ""
    # print me
    msg = MIMEText(content, _subtype='plain')
    msg['Subject'] = sub
    msg['From'] = me
    msg['To'] = ";".join(to_list)  # 将收件人列表以‘；’分隔
    try:
        server = smtplib.SMTP()
        server.connect(mail_host)  # 连接服务器
        server.login(mail_user, mail_pass)  # 登录操作
        server.sendmail(me, to_list, msg.as_string())
        server.close()
        return True
    except Exception, e:
        # print e
        return False
    return False


# 邮件接收函数
def accpet_mail(accpet_host, accpet_user, accpet_pass):
    mail_list = []
    # print accpet_host
    # print "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
    try:
        p = poplib.POP3(accpet_host)
        p.user(accpet_user)
        p.pass_(accpet_pass)
        (mail_count, mail_total_size) = p.stat()  # 返回一个元组:(邮件数,邮件尺寸)
        for i in range(mail_count):
            mail_map = {}
            # 邮件从1开始读取,retr方法返回一个元组:(状态信息,邮件,邮件尺寸)
            status_info, email_info, email_size = p.retr(str(i + 1))
            # print decode_header(email.message_from_string(str(email_info)))[0][0]
            # print email.message_from_string(str(email_info))
            message = {}
            last_email_item, append_flag = "", True
            # print str(decode_header(email_info)[0][0])
            remove_values_from_list(email_info, '')
            for j, email_item in enumerate(email_info):
                if j == len(email_info) - 1:
                    message["Content"] = email_item
                    break
                email_item_list = "".join(email_item.split("\r\n")).split(": ")
                if len(email_item_list) != 2:
                    if append_flag == True:
                        message[last_email_item] = message.get(last_email_item) + email_item_list[0]
                        # print email_item_list[0]
                else:
                    if email_item_list[0] in message:
                        append_flag = False
                        continue
                    message[email_item_list[0]] = email_item_list[1]
                    last_email_item = email_item_list[0]
            if "Subject" not in message or "Content" not in message:
                p.dele(str(i + 1))
                continue
            # print i+1,message.get("Subject"),message.get("Content")
            content_match_list = re.findall(r"The sending time", message.get("Content"))
            subject_match_list = re.findall(r"MAILVIEW", message.get("Subject"))
            if subject_match_list == [] or content_match_list == []:
                p.dele(str(i + 1))
                # print i+1,message.get("Subject"),message.get("Content")
                # print "2:"+str(email_info)
                continue
            # print str(i+1)
            content_list = message.get("Content").split("#")
            send_list = content_list[0].split(":")
            # print "send_time" + str(datetime.strptime(send_list[1],'%Y%m%d%H%M%S'))
            # mail_map["send_time"] =str(datetime.strptime(send_list[1],'%Y%m%d%H%M%S'))
            # mail_map["recive_time"] =  str(datetime.strptime(message.get("Date"),'%a, %d %b %Y %H:%M:%S +0800'))
            sr_list = content_list[1].split("@")
            from_list = sr_list[0].split(":")
            to_list = sr_list[1].split(":")
            # print from_list[0] + ":" + from_list[1]
            mail_map[from_list[0]] = from_list[1]
            # print to_list[0] + ":" + to_list[1]
            mail_map[to_list[0]] = to_list[1]
            # print "recive_time:"+str(datetime.strptime(message.get("Date"),'%a, %d %b %Y %H:%M:%S +0800'))
            # Mon, 30 Mar 2015 14:20:58 +0800  '%a, %d %b %Y %H:%M:%S +0800'
            # mail_map["send_time"] =  str(datetime.strptime(message.get("Date"),'%a, %d %b %Y %H:%M:%S +0800'))
            # 从字符串读取信息-->解密邮件头-->字符串取[0][0]-->以"\n"分组-->取第一个-->替换字符串"From nobody "为空-->格式化为日期格式-->得到接收时间
            # mail_map["recive_time"] = str(datetime.strptime(str(decode_header(email.message_from_string(str(email_info)))[0][0]).split("\n")[0].replace("From nobody ",""),'%a %b %d %H:%M:%S %Y'))

            mail_map["send_time"] = str(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(
                time.mktime(email.utils.parsedate(message.get('Date'))))))
            receive_time_list = message.get('Received').split(";")
            mail_map["recive_time"] = str(
                datetime.strptime(receive_time_list[len(receive_time_list) - 1].replace(" (CST)", "").lstrip(),
                                  '%a, %d %b %Y %H:%M:%S +0800'))
            # print message
            # print "send_time: %s" % mail_map.get('send_time')
            # print "recive_time: %s" %  mail_map.get('recive_time')
            # print "----------------------------------------------------------------"
            mail_list.append(mail_map)
            p.dele(str(i + 1))
        p.quit()
        return mail_list
    except poplib.error_proto, e:
        # print "Login failed:",e
        p.quit()
        return mail_list
        # sys.exit(1)
    return mail_list


# 运行当前文件时，执行sendmail和accpet_mail函数
if __name__ == "__main__":
    '''
    mailto_list,mail_send_host=['gacfiat@126.com'],"smtp.gacfiatauto.com"
    mail_send_user,mail_send_pass,mail_send_postfix = "test","GACFiat1234","gacfiatauto.com"
    NOW = datetime.now()
    now_str = datetime.strftime(NOW,'%Y%m%d%H%M%S')
    sub = "MAILVIEW"+now_str
    content = "The sending time:"+now_str+"#from_mail:JV@to_mail:126"
    send_mail(mail_send_host,mail_send_user,mail_send_pass,mail_send_postfix,mailto_list,sub,content)
    #accpet_host,accpet_user,accpet_pass = 'pop.126.com','gacfiat','fiat100?'
    accpet_host,accpet_user,accpet_pass = 'pop3.gacfiatauto.com','test','GACFiat1234' 
    accpet_mail(accpet_host,accpet_user,accpet_pass)
    '''
    DB2_HOST = '10.27.95.30'
    DB2_PORT = '50000'
    DB2_DB = 'db2s'
    DB2_USER = 'db2admin'
    DB2_PWD = 'fiat100?'
    dsn = "driver={IBM DB2 ODBC DRIVER};database=%s;hostname=%s;port=%s;protocol=tcpip;" % (DB2_DB, DB2_HOST, DB2_PORT)
    conn = pyodbc.connect(dsn + "uid=" + DB2_USER + ";pwd=" + DB2_PWD + ";");
    cursor = conn.cursor()
    cursor.execute("select * from kernel.mail_property")
    row = cursor.fetchone()
    property_list = []
    # 1 获取要发送和接收邮件的属性信息
    while row:
        (from_mail, from_protocol, from_address, from_host, auth_flag, from_port, from_user, from_password, to_mail,
         to_protocol, to_address, to_host, to_port, to_user, to_password, is_current, t2) = (
        row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12],
        row[13], row[14], row[15], row[16])
        row = cursor.fetchone()
        if is_current == "1":  # and from_mail=='126' and to_mail=="JV":
            property_list.append((from_mail, from_protocol, from_address, from_host, auth_flag, from_port, from_user,
                                  from_password, to_mail, to_protocol, to_address, to_host, to_port, to_user,
                                  to_password, is_current, t2))
    cursor.close()
    # 2 发送邮件
    for property in property_list:
        (from_mail, from_protocol, from_address, from_host, auth_flag, from_port, from_user, from_password, to_mail,
         to_protocol, to_address, to_host, to_port, to_user, to_password, is_current, t2) = property
        mailto_list = [to_address]
        mail_send_host = from_host
        mail_send_user = from_user
        mail_send_pass = from_password
        mail_send_postfix = from_host.replace(from_protocol + ".", "")
        NOW = datetime.now()
        now_str = datetime.strftime(NOW, '%Y%m%d%H%M%S')
        sub = "MAILVIEW" + now_str
        content = "The sending time:" + now_str + "#from_mail:" + from_mail + "@to_mail:" + to_mail
        # print mail_send_host
        send_resp = send_mail(mail_send_host, mail_send_user, mail_send_pass, mail_send_postfix, mailto_list, sub,
                              content)
        if send_resp:  # 如果发送失败，再尝试发送一次
            continue
        else:
            # print 'send fail,send again.'
            send_mail(mail_send_host, mail_send_user, mail_send_pass, mail_send_postfix, mailto_list, sub, content)

    # 3 接收邮件
    for property in property_list:
        (from_mail, from_protocol, from_address, from_host, auth_flag, from_port, from_user, from_password, to_mail,
         to_protocol, to_address, to_host, to_port, to_user, to_password, is_current, t2) = property
        accpet_host = to_host
        accpet_user = to_user
        accpet_pass = to_password
        starttime = datetime.now()
        # print accpet_host
        mail_list = accpet_mail(accpet_host, accpet_user, accpet_pass)
        # 4 添加邮件记录信息到DB2数据库中
        insert_sql = "insert into kernel.mail_data(from_mail,send_time,to_mail,recive_time) values "
        for i, mail in enumerate(mail_list):
            from_mail, send_time, to_mail, recive_time = mail['from_mail'], mail['send_time'], mail['to_mail'], mail[
                'recive_time']
            # print send_time,from_mail,to_mail,recive_time
            insert_sql = insert_sql + "('%s',to_date('%s','yyyy-mm-dd hh24:mi:ss'),'%s',to_date('%s','yyyy-mm-dd hh24:mi:ss'))," % (
            from_mail, send_time, to_mail, recive_time)
            if i == len(mail_list) - 1:
                insert_sql = insert_sql[:len(insert_sql) - 1]

        endtime = datetime.now()
        message = '开始时间：%s，总共耗时：%s s' % (starttime, (endtime - starttime).seconds)
        # print message
        if mail_list != None and mail_list != []:
            # print insert_sql
            cursor = conn.cursor()
            cursor.execute(insert_sql)
            cursor.commit()
            cursor.close()
            # sys.exit(1)