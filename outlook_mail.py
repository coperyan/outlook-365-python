from exchangelib import DELEGATE, Account, Credentials, Message, FileAttachment, Mailbox, Configuration

#Defining outlook class
class OutlookMail:

    #Initialization
    def __init__(self,username,password,email):
        self.username = username
        self.password = password
        self.email = email
        self.subject = None
        self.account = None
        self.message = None
        self.body = None
        self.myfile = None
        self.to_recipients = []
        self.cc_recipients = []
        self.bcc_recipients = []

    #Adding recipients function
    def add_recipients(self, to_recipients, cc_recipients = None, bcc_recipients = None):
        if cc_recipients is not None:
                self.cc_recipients.extend(cc_recipients)

        if bcc_recipients is not None:
                self.bcc_recipients.extend(bcc_recipients)

        self.to_recipients.extend(to_recipients)

    #Adding message subject & body function
    def add_message(self, subject, body):
        self.email = self.email
        self.subject = subject
        self.body = body

        self.mail_credentials = Credentials(username = self.username
                                , password = self.password)

        self.mail_config = Configuration(server = 'outlook.office365.com', credentials = self.mail_credentials)

        self.account = Account(primary_smtp_address = self.email
                                , config = self.mail_config
                                , autodiscover = False
                                , access_type = DELEGATE)
        try:
                if len(self.bcc_recipients) == 0 and len(self.cc_recipients) == 0 and len(self.to_recipients) != 0:
                        self.message = Message(account = self.account
                                                , folder = self.account.sent
                                                , subject = self.subject
                                                , body = self.body
                                                , to_recipients = self.to_recipients)

                elif len(self.cc_recipients) == 0:
                        self.message = Message(account = self.account
                                , folder = self.account.sent
                                , subject = self.subject
                                , body = self.body
                                , to_recipients = self.to_recipients
                                , bcc_recipients = self.bcc_recipients)

                elif len(self.bcc_recipients) == 0:
                        self.message = Message(account = self.account
                                , folder = self.account.sent
                                , subject = self.subject
                                , body = self.body
                                , to_recipients = self.to_recipients
                                , cc_recipients = self.cc_recipients)
        except MissingEmail:
                print('Missing recipients email addresses.')


    #Adding attachment
    def add_attachment(self, file, name):
        with open(file, 'rb') as f:
                content = f.read()
        self.myfile = FileAttachment(name = name, content = content)
        self.message.attach(self.myfile)

    #Sending mail
    def send_mail(self):
        self.message.send_and_save()
