import time
date_467 = 'Sat, 2 Mar 2013 6:47:09 +0800'
date_1041 = 'Tue, 5 May 2015 17:37:07 +0800'
from email.utils import parsedate, parsedate_to_datetime
# date 467
date_ = parsedate(date_1041)
date__str = time.strftime("%Y%m%d",date_)
print(date__str)
# date1 = time.strptime(date_467[0:24], '%a, %d %b %Y %H:%M:%S')  # 格式化收件时间
# send_date_str = time.strftime("%Y%m%d", date1)
# print(send_date_str)

# date 1041
# date2 = time.strptime(date_1041[0:24], '%a, %d %b %Y %H:%M:%S')  # 格式化收件时间 <class 'time.struct_time'>
# send_date_str_2 = time.strftime("%Y%m%d", date2)
# print(send_date_str_2)
