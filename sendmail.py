# -*- coding: utf-8 -*-
"""
Created on Fri May  1 20:00:16 2020

@author: Aditya Bhattacharya
"""
import pandas as pd
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime import message

SenderAddress = "toaditya1999@gmail.com"
password = "9635325858"

e = pd.read_excel("sendmail.xlsx")
emails = e['Emails'].values
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(SenderAddress, password)
msg = "Hello this is a email from python"
subject = "Test mail"
body = "Subject: {}\n\n{}".format(subject,msg)
for email in emails:
    filename = e['file'].values
    with open(filename, "rb") as attachment:
    # Add file as application/octet-stream
    # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

# Encode file in ASCII characters to send by email    
    encoders.encode_base64(part)

# Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
)

# Add attachment to message and convert message to string
    message.attach(part)
    text = message.as_string()
    server.sendmail(SenderAddress, email, text)
    print("mail sent")
server.quit()

    