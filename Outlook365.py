import os
from exchangelib import (
    Credentials,
    Configuration,
    Account,
    DELEGATE,
    FileAttachment,
    Message,
)
from bs4 import BeautifulSoup

SERVER = "outlook.office365.com"

# Documentation
# https://ecederstrand.github.io/exchangelib/


class OutlookException(Exception):
    """Exception obj"""

    def __init__(self, error_msg):
        self.error_msg = error_msg

    def __str__(self):
        return self.error_msg


class Outlook:
    def __init__(self, username, password, email=None):
        """Standard outlook class
        Aauths to Account object
        """
        self.username = username
        self.password = password
        self.email = username if not email else email
        self.authenticate()

    def authenticate(self):
        """Authenticates to O365 account
        Username must be xxxxx@xxxx.com
        Email can be same as username or shared mailbox, etc.
        """
        try:
            self.mail_credentials = Credentials(
                username=self.username, password=self.password
            )
            self.mail_config = Configuration(
                server=SERVER, credentials=self.mail_credentials
            )
            self.account = Account(
                primary_smtp_address=self.email,
                config=self.mail_config,
                autodiscover=False,
                access_type=DELEGATE,
            )
        except Exception as e:
            raise OutlookException(e)


class OutlookMessage(Outlook):
    """Outlook message object
    Pass user, pass, email for authentication
        (email can be blank and will default to user)
    Add to/cc/bcc recipients as lists (can be empty)
    Subject = string
    Body = string OR HTMLBODY object from exchangelib
    file_attachments = list with full filepath
    img_attachments - use for inline references in email via HTMLBody
    importance = string
    """

    def __init__(
        self,
        username,
        password,
        email=None,
        to_recipients=[],
        cc_recipients=[],
        bcc_recipients=[],
        subject=None,
        body=None,
        file_attachments=[],
        img_attachments=[],
        importance="Normal",
    ):
        """ """
        super(OutlookMessage, self).__init__(
            username, password, username if not email else email
        )
        self.to_recipients = to_recipients
        self.cc_recipients = cc_recipients
        self.bcc_recipients = bcc_recipients
        self.subject = subject
        self.body = body
        self.file_attachments = file_attachments
        self.img_attachments = img_attachments
        self.importance = (
            importance if importance in ["High", "Low", "Normal"] else "Normal"
        )
        self.msg_item = None
        self.build_message()
        self.add_attachments()
        self.send_it()

    def build_message(self):
        """Function called by object initialization
        Checks for recipients, creates message object
        """
        if (
            len(self.to_recipients) > 0
            or len(self.cc_recipients) > 0
            or len(self.bcc_recipients) > 0
        ):
            self.msg_item = Message(
                account=self.account,
                folder=self.account.sent,
                subject=self.subject,
                body=self.body,
                to_recipients=self.to_recipients,
                cc_recipients=self.cc_recipients,
                bcc_recipients=self.bcc_recipients,
                importance=self.importance,
            )
        else:
            raise OutlookException("Must add recipients..")

    def add_attachments(self):
        """
        Iteration for file / inline image attachments
        """
        for attachment in self.file_attachments:
            with open(attachment, "rb") as f:
                content = f.read()
            iter_file = FileAttachment(
                name=os.path.basename(attachment), content=content
            )
            self.msg_item.attach(iter_file)

        for attachment in self.img_attachments:
            with open(attachment, "rb") as f:
                iter_file = FileAttachment(
                    name=os.path.basename(attachment),
                    content=f.read(),
                    is_inline=True,
                    content_id=os.path.basename(attachment),
                )
            self.msg_item.attach(iter_file)

    def send_it(self):
        """
        Sends / saves message
        """
        self.msg_item.send_and_save()


class OutlookMailboxItem:
    """Outlook Mailbox Items (effectively message item)
    TBD: Add notes
    """

    def __init__(self, msg):
        self.msg = msg
        self.attachments = msg.attachments
        self.saved_files = []
        self.get_metadata()

    def get_metadata(self):
        """ """
        self.id = self.msg.id
        self.subject = self.msg.subject
        self.to_recipients = (
            []
            if not self.msg.to_recipients
            else [x.email_address for x in self.msg.to_recipients]
        )
        self.cc_recipients = (
            []
            if not self.msg.cc_recipients
            else [x.email_address for x in self.msg.cc_recipients]
        )
        try:
            soup = BeautifulSoup(self.msg.body, features="lxml")
            soup = soup.find_all("p")
            self.body = ""
            for line in soup:
                self.body + line.text + "\n"
        except:
            self.body = ""

    def save_attachments(self, path, contains=None):
        """ """
        if contains:
            attachments = [x for x in self.attachments if contains in x.name]
        else:
            attachments = self.attachments

        buffer_size = 1024
        for attachment in attachments:
            name = attachment.name
            save_path = os.path.join(path, name)
            if os.path.exists(save_path):
                save_path = os.path.join(
                    path, name.replace(".xlsx", f"_{(len(os.listdir(path))+1)}.xlsx")
                )
            with open(save_path, "wb") as f, attachment.fp as fp:
                buffer = fp.read(buffer_size)
                while buffer:
                    f.write(buffer)
                    buffer = fp.read(buffer_size)
            print(f"Saved {save_path}")
            self.saved_files.append(save_path)

    def mark_as_read(self):
        """ """
        self.msg.is_read = True
        self.msg.save()


class OutlookMailboxSearch(Outlook):
    """Class to search a mailbox for items
    Pass standard auth, mailbox string
    Subject contains = string search
    Has attachments boolean
    Is read boolean
    Limit integer
    Check item_results for list of OutlookMailboxItems
    """

    def __init__(
        self,
        username,
        password,
        email=None,
        mailbox=None,
        subject_contains=None,
        has_attachments=None,
        is_read=None,
        limit=100,
    ):
        super(OutlookMailboxSearch, self).__init__(
            username, password, username if not email else email
        )
        self.folder = (
            self.account.inbox
            if mailbox == "inbox"
            else self.account.sent
            if mailbox == "sent"
            else self.account.inbox
        )
        self.timezone = self.account.default_timezone
        self.search(has_attachments, is_read, subject_contains, limit)

    def search(self, has_attachments, is_read, subject_contains, limit):
        """ """
        self.results = list(
            self.folder.filter(
                has_attachments=has_attachments,
                is_read=is_read,
                subject__icontains=subject_contains,
            ).order_by("-datetime_received")
        )[:limit]

        self.item_results = [OutlookMailboxItem(m) for m in self.results]
