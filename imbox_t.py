import datetime
from imbox import Imbox
# #登录邮箱
# with Imbox('imap.qq.com', username='915288938@qq.com', password='zlfdjwztlgepbefj', ssl=True, ssl_context=None, starttls=False) as imbox:
#     folders = imbox.folders()
#     all_messages = imbox.messages()
#     for id, m in all_messages:
#         subjects = m.subject
#         date = m.date
#         print(subjects)
#         print(date)


with Imbox('imap.qq.com', username='915288938@qq.com', password='zlfdjwztlgepbefj', ssl=True, ssl_context=None, starttls=False) as imbox:
    spe_date = imbox.messages(folder='INBOX')
    for id, m in spe_date:
        try:
            x = m.subject
            print('\n'+x)
            print(m.attachments[0]['filename'])

        except:
            pass
