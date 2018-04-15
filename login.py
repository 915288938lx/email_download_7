import datetime
import time
detestr2 = '2017-01-01'
date2 = datetime.datetime.strptime(detestr2,'%Y-%m-%d')
print(type(date2))

today = time.localtime(time.time())


time.gmtime(today)

print(today)

n = (today-date2).days

print(n)