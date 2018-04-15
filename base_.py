a=b'\tcharset="gb2312"'
x = b'Received: from mail.message.cmbchina.com (unknown [61.144.248.21])'
b = a.decode('utf-8')
x_= x.decode('utf-8')
print(b)
print(x_)