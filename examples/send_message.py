import os

from Outlook365 import OutlookMessage

user = 'emailacct@email.com'
pwd = 'password'
to_emails = ['toemail@email.com']
subject = 'Test Email'
body = 'Hi,\n\nThis is a test email.'

msg = OutlookMessage(
    username = user,
    password = pwd,
    email = None, #defaults to username email
    to_recipients = to_emails,
    subject = subject,
    body = body
)