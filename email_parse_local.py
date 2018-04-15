import re
a = '估值日期：2018-0124'
# b = re.findall('.*?([1,2]{1}[0,9]{1}[0-9]{2})[/-]([0,1]{1}[0-9]{1})[/-]([0-3]{1}[0-9]{1}).*?',a)
# print(type(b))
# print(len(b))
# print(b[0])
# print(''.join(b[0]))

sp = re.split('\.','woede.exe.txt')
print(sp[len(sp) - 1])

pattern_2 = '.*?([1,2]{1}[0,9]{1}[0-9]{2}[0,1]{1}[0-9]{1}[0-3]{1}[0-9]{1}).*?'
print(re.findall(pattern_2,a))