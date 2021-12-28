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
message = '''
Some message here in <b>HTML</b> mode.<br><br>
Some thank you message!
'''

def send_email(address_list, cc_address_list, subject, name, message=''):
    assert subject
    assert name
    assert len(address_list) >= 1
    mail = outlook.CreateItem(0)
    mail.To = ';'.join(address_list)
    mail.Subject = subject
    mail.HTMLBody = template.format(name) + message + ending
    if cc_address_list:
        mail.CC = ';'.join(cc_address_list)
    mail.Send()

# Send the same email to multiple recipients
send_email(address_list, cc_address_list, subject, name, message)

# Send email with different names to multiple recipients, assuming no CC
names = ["Name1", "Name2"]
assert len(names) == len(address_list)
for name, add in zip(names, address_list):
    send_email([add], [], subject, name, message)