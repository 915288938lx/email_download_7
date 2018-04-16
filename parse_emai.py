import poplib
import email
import base64
import time
from email.header import decode_header
import re
from email.utils import parseaddr , parsedate
from email.parser import Parser
charset_collect= ['utf-8','gb2312', 'iso-8859-6', 'windows-1256', 'ibm775', 'iso-8859-4', 'windows-1257', 'ibm852', 'iso-8859-2', 'x-mac-ce', 'windows-1250', 'gb18030', 'EUC-CN', 'gb18030', 'hz-gb-2312', 'x-mac-chinesesimp', 'big5', 'x-Chinese-CNS', 'x-Chinese-Eten', 'x-mac-chinesetrad', 'cp866', 'iso-8859-5', 'koi8-r', 'koi8-u', 'x-mac-cyrillic', 'windows-1251', 'x-Europa', 'x-IA5-German', 'ibm737', 'iso-8859-7', 'x-mac-greek', 'windows-1253', 'ibm869', 'iso-8859-8-i', 'iso-8859-8', 'DOS-862', 'x-mac-hebrew', 'windows-1255', 'ASMO-708', 'DOS-720', 'x-mac-arabic', 'x-EBCDIC-Arabic', 'x-EBCDIC-CyrillicRussian', 'x-EBCDIC-CyrillicSerbianBulgarian', 'x-EBCDIC-DenmarkNorway', 'x-ebcdic-denmarknorway-euro', 'x-EBCDIC-FinlandSweden', 'x-ebcdic-finlandsweden-euro', 'x-ebcdic-finlandsweden-euro', 'x-ebcdic-france-euro', 'x-EBCDIC-Germany', 'x-ebcdic-germany-euro', 'x-EBCDIC-GreekModern', 'x-EBCDIC-Greek', 'x-EBCDIC-Hebrew', 'x-EBCDIC-Icelandic', 'x-ebcdic-icelandic-euro', 'x-ebcdic-international-euro', 'x-EBCDIC-Italy', 'x-ebcdic-italy-euro', 'x-EBCDIC-JapaneseAndKana', 'x-EBCDIC-JapaneseAndJapaneseLatin', 'x-EBCDIC-JapaneseAndUSCanada', 'x-EBCDIC-JapaneseKatakana', 'x-EBCDIC-KoreanAndKoreanExtended', 'x-EBCDIC-KoreanExtended', 'CP870', 'x-EBCDIC-SimplifiedChinese', 'X-EBCDIC-Spain', 'x-ebcdic-spain-euro', 'x-EBCDIC-Thai', 'x-EBCDIC-TraditionalChinese', 'CP1026', 'x-EBCDIC-Turkish', 'x-EBCDIC-UK', 'x-ebcdic-uk-euro', 'ebcdic-cp-us', 'x-ebcdic-cp-us-euro', 'ibm861', 'x-mac-icelandic', 'x-iscii-as', 'x-iscii-be', 'x-iscii-de', 'x-iscii-gu', 'x-iscii-ka', 'x-iscii-ma', 'x-iscii-or', 'x-iscii-pa', 'x-iscii-ta', 'x-iscii-te', 'euc-jp', 'iso-2022-jp', 'iso-2022-jp', 'csISO2022JP', 'x-mac-japanese', 'shift_jis', 'ks_c_5601-1987', 'euc-kr', 'iso-2022-kr', 'Johab', 'x-mac-korean', 'iso-8859-3', 'iso-8859-15', 'x-IA5-Norwegian', 'IBM437', 'x-IA5-Swedish', 'windows-874', 'ibm857', 'iso-8859-9', 'x-mac-turkish', 'windows-1254', 'unicode', 'unicodeFFFE', 'utf-7', 'utf-8', 'us-ascii', 'windows-1258', 'ibm850', 'x-IA5', 'iso-8859-1', 'macintosh', 'Windows-1252']

pop = poplib.POP3_SSL('pop.qq.com')
pop.user('915288938@qq.com')
pop.pass_('fshsnxoporssbbja')
list = pop.list()
lines = pop.retr(913)[1]
# print(b'\r\n'.join(lines).decode('gb2312'))
# for line in lines:
#     print(lines)
def decode_str(s):
    try:
        value, charset = decode_header(s)[0]
        if charset:
            value = value.decode(charset)
        return value, charset
    except:
        pass

def decode_email_content():
    global msg
    counter = 0
    str_lines = ''
    for charset in charset_collect:
        try:
            str_lines = b'\r\n'.join(lines).decode(charset)

            break
        except:
            counter = counter + 1
            continue
    print(str_lines)
    msg = Parser().parsestr(str_lines)



def ood_parse_date(date_):
    if parsedate(date_):
        date_tuple = parsedate(date_)
        send_date_str = time.strftime("%Y%m%d", date_tuple)
    else:
        send_date_str = ''.join(re.findall('.*?([1,2]{1}[0,9]{1}[0-9]{2})[/-]([0,1]{1}[0-9]{1})[/-]([0-3]{1}[0-9]{1}).*?', date_)[0])
    return send_date_str

# b_lines = b'\r\n'.join(lines)
# str_lines = str(b_lines,'utf-8')
# print(str_lines)
decode_email_content()
hdr, addr = parseaddr(msg.get('From'))
address = u'%s' % (addr)  # 发件人邮箱
print('adress为:%s'%address)
charset = decode_str(msg.get('Subect'))
print('subject为:%s'%decode_str(msg.get('Subject'))[0])
print('subject为:%s'%decode_str(msg.get('Subject'))[0].encode('utf-8').decode('utf-8'))
print('subject为:%s'%decode_str(msg.get('字符集为:'))[1])
date_ =msg.get('Date')
print('日期为:%s'%date_)
print('date为:%s'%ood_parse_date(date_))
# date__ = parsedate(date_)
# print(date__)
# date__str = time.strftime("%Y%m%d",date__)
# print(date__str)
# print(address)