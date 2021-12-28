import os
assert os.name == 'nt' # only works for Windows OS

import win32com.client as win32

outlook = win32.Dispatch('outlook.application')

template = '''
Dear {},
<br><br>
'''

ending = '''
<br><br>
Best regards,<br>
<b>Russell</b>
'''

subject = 'Sample Email'
name = '<what do you want to put after Dear>'
address_list = ['youremail@address1.com', 'youremail@address2.com']
cc_address_list = ['yourccemail@address1.com', 'yourccemail@address2.com']

def main(message=''):
    assert len(address_list) >= 1
    mail = outlook.CreateItem(0)
    mail.To = ';'.join(address_list)
    mail.Subject = subject
    mail.HTMLBody = template.format(name) + message + ending
    if cc_address_list:
        mail.CC = ';'.join(cc_address_list)
    mail.Send()

main('''
Some message here in <b>HTML</b> mode.<br><br>
Some thank you message!
''')