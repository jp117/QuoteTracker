from __future__ import print_function
import httplib2
import os
import datetime

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import send_email
import QuoteSpreadsheet

import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import mimetypes

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

import auth

def messageform(sender, recepient, subject, emailbody):
    return sendInst.create_message(sender, recepient, subject, emailbody)


SCOPES = 'https://mail.google.com/'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Gmail API Python Quickstart'
authInst = auth.auth(SCOPES,CLIENT_SECRET_FILE,APPLICATION_NAME)
credentials = authInst.get_credentials()

http = credentials.authorize(httplib2.Http())
service = discovery.build('gmail', 'v1', http=http)

sendDate = datetime.datetime.now().strftime('%m/%d/%Y')
sendInst = send_email.send_email(service)
sender = "john@atlasswitch.com"
subject = "Quotes to Follow Up " + sendDate


#Go through each employee and send
emailbody = QuoteSpreadsheet.emailSubBodySalesman("Harris")
message = messageform(sender, "harris@atlasswitch.com", subject, emailbody)
sendInst.send_message('me', message)

emailbody = QuoteSpreadsheet.emailSubBodySalesman("Steve")
message = messageform(sender, "steve@atlasswitch.com", subject, emailbody)
sendInst.send_message('me', message)

emailbody = QuoteSpreadsheet.emailSubBodySalesman("Matt")
message = messageform(sender, "matthew@atlasswitch.com", subject, emailbody)
sendInst.send_message('me', message)

#if new salesman, add similar to above

emailbody = QuoteSpreadsheet.emailSubBodyExec("John")
message = messageform(sender, "john@atlasswitch.com", subject, emailbody)
sendInst.send_message('me', message)

emailbody = QuoteSpreadsheet.emailSubBodyExec("Gina")
message = messageform(sender, "gina@atlasswitch.com", subject, emailbody)
sendInst.send_message('me', message)

#if new exec, add similar to above