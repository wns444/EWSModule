import os
from datetime import timedelta

from exchangelib import Q, Configuration, Account, DELEGATE
from exchangelib import Message, Mailbox, FileAttachment
from exchangelib import Credentials
from exchangelib import UTC_NOW
import exchangelib
from exchangelib.errors import DoesNotExist
from exchangelib.fields import MailboxListField
from exchangelib.folders import FolderQuerySet
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
from exchangelib.protocol import NoVerifyHTTPAdapter
from exchangelib.protocol import BaseProtocol

BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter

class EWSWorker:
    def __init__(self, username: str, password: str, server_endpoint: str, smtp_address: str):
        self.Credentials: Credentials = Credentials(username=username, password=password)
        self.Config: Configuration = Configuration(credentials=self.Credentials,service_endpoint=server_endpoint)
        self.Account: Account = Account(
            primary_smtp_address=smtp_address,
            config=self.Config,
            access_type=DELEGATE
        )

    def send_message(
        self,
        recipients: list[str], 
        cc_recipients: list[str] | None = None, 
        bcc_recipients: list[str] | None = None,
        subject: str | None = None, 
        body: str | None = None, 
        html_body: str | None = None,
        file_paths_Attachments: list[str] | None = None,
        inline_paths_attachments: list | None = None,
    ):
       
        to_recipients = list()
        to_recipients.extend([Mailbox(email_address=recipient) for recipient in recipients])

        to_ccrecipients = list()
        to_ccrecipients.extend([Mailbox(email_address=cc_recipient) for cc_recipient in cc_recipients])

        to_bccrecipients = list()
        to_bccrecipients.extend([Mailbox(email_address=bcc_recipient) for bcc_recipient in bcc_recipients])

        msg_obj = Message(
            account = self.Account,
            folder = self.Account.sent,
            to_recipients = to_recipients,
            cc_recipients = to_ccrecipients,
            bcc_recipients = to_bccrecipients
        )

        if subject:
            msg_obj.subject = str(subject)
        else:
            msg_obj.subject = 'Без темы'

        if body:
            msg_obj.body = str(body)
        elif html_body:
            msg_obj.body = html_body
        else:
            msg_obj.body = ''
        if isinstance(file_paths_Attachments, list):
            for path_attachment in file_paths_Attachments:
                attachment_content = open(path_attachment,'rb').read()
                attach_to_send = FileAttachment(name =os.path.basename(path_attachment),content = attachment_content, is_inline = False)
                msg_obj.attach(attach_to_send)

        for path_attachment in inline_paths_attachments:
            attachment_content = open(path_attachment,'rb').read()
            attach_to_send = FileAttachment(name =os.path.basename(path_attachment),content = attachment_content, is_inline = True,content_id=os.path.basename(path_attachment))
            msg_obj.attach(attach_to_send)
       
        msg_obj.send_and_save()

    def forward_message(
        self, 
        message: Message,
        subject: str | None = None,
        body: str | None = None,
        recipients: list[str] | None = None,
        cc_recipients:list[str] | None = None, 
        bcc_recipients: list[str] | None = None
    ):
        to_recipients = list()
        if isinstance(recipients, list):
            to_recipients.extend([Mailbox(email_address=recipient) for recipient in recipients])

        to_ccrecipients = list()
        if isinstance(cc_recipients, list):
            to_ccrecipients.extend([Mailbox(email_address=cc_recipient) for cc_recipient in cc_recipients])

        to_bccrecipients = list()
        if isinstance(bcc_recipients, list):
            to_bccrecipients.extend([Mailbox(email_address=bcc_recipient) for bcc_recipient in bcc_recipients])
        
        if subject:
            subject_to_reply = subject
        else:
            subject_to_reply = f"FW: {message.subject}"
            
        message.forward(subject=subject_to_reply, to_recipients=to_recipients, cc_recipients=to_ccrecipients, bcc_recipients=to_bccrecipients, body=body)


    def reply_message(
        self, 
        message: Message,
        subject: str | None = None,
        body: str | None = None,
        recipients: list[str] | None = None,
        cc_recipients:list[str] | None = None, 
        bcc_recipients: list[str] | None = None, 
    ):
        
        to_recipients = list()
        if isinstance(recipients, list):
            to_recipients.extend([Mailbox(email_address=recipient) for recipient in recipients])
        if isinstance(message.to_recipients, MailboxListField):
            to_recipients.extend([recipient for recipient in message.to_recipients])

        to_ccrecipients = list()
        if isinstance(cc_recipients, list):
            to_ccrecipients.extend([Mailbox(email_address=cc_recipient) for cc_recipient in cc_recipients])
        if isinstance(message.cc_recipients, MailboxListField):
            to_ccrecipients.extend([cc_recipient for cc_recipient in message.cc_recipients])

        to_bccrecipients = list()
        if isinstance(bcc_recipients, list):
            to_bccrecipients.extend([Mailbox(email_address=bcc_recipient) for bcc_recipient in bcc_recipients])
        

        if subject:
            subject_to_reply = subject
        else:
            subject_to_reply = f"RE: {message.subject}"

        message.reply(to_recipients=to_recipients, subject=subject_to_reply,cc_recipients=to_ccrecipients, bcc_recipients=to_bccrecipients, body=body)
    
    def get_message_byID(self, msg_id: str, folder_name: str | None = None) -> Message:

        folder_messages: FolderQuerySet = self.Account.inbox
        
        if folder_name:
            folder_messages: FolderQuerySet = self.Account.root / 'Корневой уровень хранилища' / folder_name 

        message = folder_messages.get(message_id=msg_id) 

        if isinstance(message, Message):
            return message
        else:
            raise DoesNotExist(f"Сообщение с ID {msg_id} не найдено в папке {folder_name}.")

    def get_messages(
        self,
        subject_contains: str | list | None = None,
        body_contains: str | list | None = None,
        senders_emails: str | list | None = None,
        is_read: bool | None = None,
        days_sience: int | None = None,
        folder_name: str | None = None
    ) -> exchangelib.queryset.QuerySet:

        folder_messages = self.Account.inbox
        if folder_name:
            folder_messages = self.Account.root / 'Корневой уровень хранилища' / folder_name
        
        # Базовые фильтры
        query = Q()
        
        if isinstance(days_sience, int):
            query &= Q(datetime_received__gte=UTC_NOW() - timedelta(days=days_sience))
        
        if isinstance(is_read, bool):
            query &= Q(is_read=is_read)

        if isinstance(body_contains, str):
            query &= Q(body__contains=body_contains)
        elif isinstance(body_contains, list):
            for body in body_contains:
                query &= Q(body__contains=body)

        if isinstance(subject_contains, str):
            query &= Q(subject__contains=subject_contains)
        elif isinstance(subject_contains, list):
            for subj in subject_contains:
                query &= Q(subject__contains=subj)

        if isinstance(senders_emails, str):
            sender_mailbox = Mailbox(email_address = senders_emails)
            query &= Q(sender=sender_mailbox.email_address)
        elif isinstance(senders_emails, list):
            for email in senders_emails:
                sender_mailbox = Mailbox(email_address = email)
                query &= Q(sender=sender_mailbox.email_address)

        return folder_messages.filter(query)