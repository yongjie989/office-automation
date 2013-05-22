# -*- coding: cp950 -*-

import smtplib
import os
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email.Utils import COMMASPACE, formatdate
from email import Encoders


def send_mail(send_from, send_to, subject, text, files=[], server="localhost"):
    assert type(send_to)==list
    assert type(files)==list

    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach( MIMEText(text) )

    for f in files:
        part = MIMEBase('application', "octet-stream")
        part.set_payload( open(f,"rb").read() )
        Encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(f))
        msg.attach(part)

    smtp = smtplib.SMTP(server)
    smtp.set_debuglevel(1)
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.close()

send_mail(
    send_from = 'Yong Jie Huang <yongjie989@gmail.com>',
    send_to = ['user1@example.com'],
    subject = 'Subject: daily report',
    text = 'Here is email body.',
    files = ['c:\\automail\\daily_report.ppt'],
    server = 'smtp.example.com'
    )
