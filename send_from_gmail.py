#!/usr/bin/env python
# -*- coding: utf-8 -*-

import smtplib
import json
from email import encoders, utils
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart

with open('secret.json') as sf:
    data = json.load(sf)


def create_message(from_addr, to_addr, subject, body, mine, attach_file):
    """
    Mailのメッセージを構築する
    """
    msg = MIMEMultipart()
    msg["Subject"] = subject
    msg["From"] = from_addr
    msg["To"] = ','.join(to_addr)
    msg["Date"] = utils.formatdate()

    body = MIMEText(body)
    msg.attach(body)

    attachment = MIMEBase(mine['type'], mine['subtype'])

    file = open(attach_file['path'], 'rb')
    attachment.set_payload(file.read())
    file.close()
    encoders.encode_base64(attachment)
    attachment.add_header('Content-Disposition', 'attachment', filename=attach_file['name'])
    msg.attach(attachment)
    return msg


def send_gmail(from_addr, to_addr, msg):
    """
    mailを送信する
    """
    smtp = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    smtp.ehlo()
    smtp.login(data['gmail_id'], data['gmail_secret'])
    smtp.sendmail(from_addr, to_addr, msg.as_string())
    smtp.quit()


if __name__ == '__main__':
    print("このコードはインポートして使ってね。")
