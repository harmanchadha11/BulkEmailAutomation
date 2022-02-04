import pandas as p
import configparser as cf
import constants as cnst
import smtplib as smtp
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

try:
    confParser = cf.ConfigParser()
    confParser.read(cnst.CONFIG_FILE_NAME)
    user_details = confParser[cnst.SECTION_USER]
    admin_email = user_details[cnst.ADMIN_EMAIL]
    admin_password = user_details[cnst.ADMIN_PASSWORD]

    data = p.read_excel(cnst.EXCEL_FILE_NAME)
    names = list(data.get(cnst.NAME))
    emails = list(data.get(cnst.EMAIL))


    mail_server = smtp.SMTP('smtp.gmail.com',587)
    mail_server.starttls()
    mail_server.login(admin_email,admin_password)
    sender = admin_email
    receipents = emails
    message = MIMEMultipart("alternative")
    message['Subject'] = 'Testing message'
    message['from'] = sender
    html = cnst.HTML
    full_message = MIMEText(html,"html")
    message.attach(full_message)
    mail_server.send(sender,emails,message.as_string())

    print("messages has been sent to everyone")



except Exception as e:
    print(e)
