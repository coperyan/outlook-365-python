import os
from outlook_mail import *

#Credentials
username = 'email@email.com'
password = 'password123'
email = 'shared@email.com' #You can send emails via a shared mailbox or your normal email

#Recipients
to_list = [
'test123@email.com'
]

cc_list = [
'cc123@email.com'
]

#Creating subject / body strings
subject = 'Testing, attention please'
body = 'Hello, \r\n\r\nThis is a test email. \r\n\r\nThanks, \r\nRyan'

#File path and name
filepath = r'C:\Users\Administrator\Desktop\File.CSV'
filename = os.path.basename(filepath) #or you can create this manually, ex. 'File.CSV'

#Creating outlook class, sending mail
om = OutlookMail(username = username, password = password, email = email)
om.add_recipients(to_recipients = to_list, cc_recipients = cc_list)
om.add_message(subject = subject, body = body)
om.add_attachment(file = filepath, name = filename)
om.send_mail()
