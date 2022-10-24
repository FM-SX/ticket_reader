from imbox import Imbox
from datetime import datetime
import pandas as pd 

username = open('login/username', 'r').read()
password = open('login/pass', 'r').read()
host = 'imap.gmail.com'

mail = Imbox(host, username=username, password=password, ssl=True)
messages = mail.messages(raw='has:attachment')

for (uid, message) in messages:
    message.subject
    message.sent_from
    message.sent_to
    message.date
    print(message.attachments)
    break