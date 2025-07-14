# EWSWorker

A Python library for working with Exchange Web Services (EWS) using exchangelib.

## Features

- Send emails with attachments
- Forward and reply to messages
- Search messages by various criteria
- Retrieve messages by ID
- Support for HTML emails and inline attachments
- Work with different mailbox folders

## Installation

```bash
pip install EWSModule
```

Usage

```python
from EWSModule import EWSWorker

# Initialize the worker
worker = EWSWorker(
    username="your_username",
    password="your_password",
    server_endpoint="https://outlook.office365.com/EWS/Exchange.asmx",
    smtp_address="your_email@example.com"
)
```
Sending Emails
```python
# Simple email
worker.send_message(
    recipients=["recipient1@example.com", "recipient2@example.com"],
    subject="Test Email",
    body="This is a test message"
)

# Email with attachments
worker.send_message(
    recipients=["recipient@example.com"],
    subject="Email with attachments",
    body="Please see attached files",
    file_paths_Attachments=["/path/to/file1.pdf", "/path/to/file2.jpg"]
)

# HTML email with inline images
worker.send_message(
    recipients=["recipient@example.com"],
    subject="HTML Email",
    html_body="<h1>Hello</h1><p>This is HTML content <img src='cid:image1.png'></p>",
    inline_paths_attachments=["/path/to/image1.png"]
)
Working with Messages
python
# Get message by ID
message = worker.get_message_byID("AAMkAD...", folder_name="Inbox")

# Forward a message
worker.forward_message(
    message=message,
    recipients=["forward_to@example.com"],
    subject="FW: Important message",
    body="Please review this message"
)

# Reply to a message
worker.reply_message(
    message=message,
    body="Thanks for your message. Here's my response..."
)

# Search for messages
messages = worker.get_messages(
    subject_contains="Important",
    senders_emails=["manager@example.com"],
    days_sience=7,
    is_read=False,
    folder_name="Inbox"
)

for msg in messages:
    print(f"Subject: {msg.subject}, From: {msg.sender.email_address}")
```
## Configuration
The EWSWorker requires the following initialization parameters:

username: Your Exchange username

password: Your Exchange password

server_endpoint: EWS endpoint URL (e.g., "https://outlook.office365.com/EWS/Exchange.asmx")

smtp_address: Your primary SMTP address

## Error Handling
The library raises standard exchangelib exceptions, including:

DoesNotExist when a message isn't found

Authentication errors for invalid credentials

Connection errors for network issues

## Dependencies
exchangelib==5.5.0