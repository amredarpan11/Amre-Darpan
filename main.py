# coding=utf-8
from typing import List, Any

import pandas as p
import smtplib as sm
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
#from email.mime.base import MIMEBase
#from email import encoders

#read the email from file
data = p.read_excel("student.xlsx")
print(type(data))
print(data.__str__())
mail_col = data.get("mail")
list_of_emails = list(mail_col)
print(list_of_emails)


#object of smtp
server = sm.SMTP("smtp.gmail.com" , 587)
server.starttls()
   # login
server.login("amredarpan11@gmail.com","Darpan@030718")
from_ = "amredarpan11@gmail.com"
to_ = list_of_emails
message = MIMEMultipart("alternative")
message['subject'] = "test message"
message["from"] = "amredarpan11@gmail.com"

body='hi,i am sending file!'
message.attach(MIMEText(body, 'plain'))


#filename='document.txt'
#attachment=open(filename, 'rb')

#part=MIMEBase('application','octet-stream')
#part.set_payload((attachment).read())
#encoders.encode_base64(part)
#part.add_header('content-disposition' ,"attachment; filename=" +filename)
#send the message
server.sendmail(from_, to_, message.as_string())


