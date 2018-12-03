import re
with open('wa.txt') as f:
    lines = f.readlines()
    a = 1
    for line in lines:
        a = a+1
        if a % 2 == 0:
            # print('奇数行')
            # print(line)
            # print(re.split('\|',line))
            xx = re.split('\│',line)
            print(xx)

        else:
            # print('偶数行')
            # print(line)
            # print('')
            pass