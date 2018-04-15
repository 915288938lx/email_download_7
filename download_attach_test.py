import poplib
import email
import base64
import time
from email.header import decode_header

from openpyxl import Workbook
import xlrd
from email.header import decode_header
from email.utils import parseaddr , parsedate
from email.parser import Parser
charset_collect= ['utf-8','gb2312', 'iso-8859-6', 'windows-1256', 'ibm775', 'iso-8859-4', 'windows-1257', 'ibm852', 'iso-8859-2', 'x-mac-ce', 'windows-1250', 'gb18030', 'EUC-CN', 'gb18030', 'hz-gb-2312', 'x-mac-chinesesimp', 'big5', 'x-Chinese-CNS', 'x-Chinese-Eten', 'x-mac-chinesetrad', 'cp866', 'iso-8859-5', 'koi8-r', 'koi8-u', 'x-mac-cyrillic', 'windows-1251', 'x-Europa', 'x-IA5-German', 'ibm737', 'iso-8859-7', 'x-mac-greek', 'windows-1253', 'ibm869', 'iso-8859-8-i', 'iso-8859-8', 'DOS-862', 'x-mac-hebrew', 'windows-1255', 'ASMO-708', 'DOS-720', 'x-mac-arabic', 'x-EBCDIC-Arabic', 'x-EBCDIC-CyrillicRussian', 'x-EBCDIC-CyrillicSerbianBulgarian', 'x-EBCDIC-DenmarkNorway', 'x-ebcdic-denmarknorway-euro', 'x-EBCDIC-FinlandSweden', 'x-ebcdic-finlandsweden-euro', 'x-ebcdic-finlandsweden-euro', 'x-ebcdic-france-euro', 'x-EBCDIC-Germany', 'x-ebcdic-germany-euro', 'x-EBCDIC-GreekModern', 'x-EBCDIC-Greek', 'x-EBCDIC-Hebrew', 'x-EBCDIC-Icelandic', 'x-ebcdic-icelandic-euro', 'x-ebcdic-international-euro', 'x-EBCDIC-Italy', 'x-ebcdic-italy-euro', 'x-EBCDIC-JapaneseAndKana', 'x-EBCDIC-JapaneseAndJapaneseLatin', 'x-EBCDIC-JapaneseAndUSCanada', 'x-EBCDIC-JapaneseKatakana', 'x-EBCDIC-KoreanAndKoreanExtended', 'x-EBCDIC-KoreanExtended', 'CP870', 'x-EBCDIC-SimplifiedChinese', 'X-EBCDIC-Spain', 'x-ebcdic-spain-euro', 'x-EBCDIC-Thai', 'x-EBCDIC-TraditionalChinese', 'CP1026', 'x-EBCDIC-Turkish', 'x-EBCDIC-UK', 'x-ebcdic-uk-euro', 'ebcdic-cp-us', 'x-ebcdic-cp-us-euro', 'ibm861', 'x-mac-icelandic', 'x-iscii-as', 'x-iscii-be', 'x-iscii-de', 'x-iscii-gu', 'x-iscii-ka', 'x-iscii-ma', 'x-iscii-or', 'x-iscii-pa', 'x-iscii-ta', 'x-iscii-te', 'euc-jp', 'iso-2022-jp', 'iso-2022-jp', 'csISO2022JP', 'x-mac-japanese', 'shift_jis', 'ks_c_5601-1987', 'euc-kr', 'iso-2022-kr', 'Johab', 'x-mac-korean', 'iso-8859-3', 'iso-8859-15', 'x-IA5-Norwegian', 'IBM437', 'x-IA5-Swedish', 'windows-874', 'ibm857', 'iso-8859-9', 'x-mac-turkish', 'windows-1254', 'unicode', 'unicodeFFFE', 'utf-7', 'utf-8', 'us-ascii', 'windows-1258', 'ibm850', 'x-IA5', 'iso-8859-1', 'macintosh', 'Windows-1252']
str_list = ['日期', '单位净值', '市值', '银行存款', '存出保证金']


def login():
    global lines
    pop = poplib.POP3_SSL('pop.163.com')
    pop.user('lx915288938@163.com')
    pop.pass_('lx12345')
    mail_count = len(pop.list()[1])
    print(mail_count)
    lines = pop.retr(21)[1]



def decode_str(s):
    try:
        value, charset = decode_header(s)[0]
        if charset:
            value = value.decode(charset)
        return value
    except:
        pass

def decode_email_content():
    global msg
    counter = 0
    str_lines = 0
    for charset in charset_collect:
        try:
            str_lines = b'\r\n'.join(lines).decode(charset)

            break
        except:
            counter = counter + 1
            continue
    msg = Parser().parsestr(str_lines)

def download_attachment(msg):
    # 下载附件
    for part in msg.walk():
        filename = part.get_filename()
        if filename:
            file_name_ = decode_str(filename)
            data = part.get_payload(decode=True)
            write_file = open(file_name_, 'wb')
            write_file.write(data)
            write_file.close()
            print('附件%s已下载...' % file_name_)
            open_workbook = xlrd.open_workbook(file_name_)
            open_sheet = open_workbook.sheet_by_index(0)
            sheet = open_sheet
            rows = sheet.nrows  # 行数
            cols = sheet.ncols  # 列数
            # 每个要查找值的行数与列数
            row_col = []
            # 将row_col的列表再装配到row_col_list中
            row_col_list = []
            for r in range(0, rows):
                for c in range(0, cols):
                    got_cell_value = sheet.cell(r, c).value
                    for str_ in str_list:
                        if str_ in got_cell_value:
                            str_row = r
                            str_col = c
                            print(str_row)
                            print(str_col)
        else:
            pass


if __name__ == '__main__':
    login()
    decode_email_content()
    hdr, addr = parseaddr(msg.get('From'))
    download_attachment(msg)
    address = u'%s' % (addr)  # 发件人邮箱
    print(address)
    print('===============任务完成===============\n')





