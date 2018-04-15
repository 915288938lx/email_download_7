import re

address_list = [
'zctgsjfs@tg.gtja.com',
'duanjiushuang@chinastock.com.cn',
'yangfan1@swhysc.com',
'zszqwbfw@stocke.com.cn',
'dfjjwb@orientsec.com.cn',
'cpbbfs@gfund.com',
'FAreport@citics.com',
'yywbfa@cmschina.com.cn',
'js@cifutures.com.cn',
'cpbbfs@gfund.com',
'admin@service.pingan.com',
'crtliangy@crctrust.com',
'yezs@cifutures.com.cn',
'zggzhd@htfc.com']
new_list = []
for add in address_list:
    new_list.append(re.split('@',add)[1])
print(new_list)
a = 'liu.d.915288@qq.com.net'
len = len(address_list)
b = re.split('@',a)
print(b)
if 'stocke.com.cn' in new_list:
    print('True')