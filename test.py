import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time
from email.mime.base import MIMEBase
from email import encoders

email_from = 'oliver@startupwriter.org'
email_password = '123@123Sw'

server = smtplib.SMTP('smtp.hostinger.com', 587)
server.connect('smtp.hostinger.com', 587)  # sometimes this is needed to establish connection first
server.ehlo()
server.starttls()
server.ehlo()
server.login(email_from, email_password)

body = "Hello - my name is " +" and I'm the client manager at Startup Writer.\n\nOur company is specialized in website content writing, especially blogs,\nnewsletters, press releases, and much more. We have built a simplified\nsystem that can help businesses like yours create website content in\nimproved quality and affordable pricing.\n\nI figured this might be of interest to you given the fact that it's hard to find the\nright talent that can help your brand talk with its customers.\n\nI'd love to get your feedback even if you're not looking for a writing team right\nnow. Do you have 20 minutes next week? I'm open Monday at 2 or 3 PM ET if \neither may work.\n\n" + ". \nwww.startupwriter.org"
msg = MIMEMultipart()
msg['From'] = email_from 
msg['To'] = 'clintonkjkj@gmail.com'
msg['Subject'] = body
msg['Reply-To'] = "sophia@startupwriter.org"
msg.attach(MIMEText(body, 'plain'))
text = msg.as_string()
try:
    server.sendmail(email_from, 'clintonkjkj@gmail.com', text)
    time.sleep(2)
    print('Mail has been sent to: ', 'clintonkjkj@gmail.com')
except smtplib.SMTPRecipientsRefused:
    print("invalid Email address:", 'clintonkjkj@gmail.com')
except TypeError:
    print("No Email found at cell:")
except smtplib.SMTPSenderRefused:
    print("Sender refused Email address:", 'clintonkjkj@gmail.com')
server.quit()
