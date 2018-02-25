import pandas as pd
import smtplib
import os.path as op
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders


def send_mail(send_from, send_to, subject, message, files=[], server="smtp.gmail.com", port=587, use_tls=True):

    msg = MIMEMultipart()
    msg['From'] = send_from
    msg['To'] = COMMASPACE.join(send_to)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(message))

    for path in files:
        part = MIMEBase('application', "octet-stream")
        with open(path, 'rb') as file:
            part.set_payload(file.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        'attachment; filename="{}"'.format(op.basename(path)))
        msg.attach(part)

    smtp = smtplib.SMTP(server, port)
    if use_tls:
        smtp.starttls()
    smtp.login('YourEmailAddress@gmail.com', 'Passwd')
    smtp.sendmail(send_from, send_to, msg.as_string())
    smtp.quit()

def writeText(txt):
    from PIL import Image, ImageFont, ImageDraw
    i=Image.open("/home/roadburner/Desktop/Certificate_of_participation.png")
    font_type=ImageFont.truetype('/usr/share/fonts/truetype/tlwg/TlwgTypo-Oblique.ttf',65)
    draw=ImageDraw.Draw(i)
    draw.text(xy=(1150,693),text=txt,fill=(0,0,0),font=font_type)
    draw.text(xy=(1150,785),text="Robo Sumo",fill=(0,0,0),font=font_type)
    i.save('/home/roadburner/Desktop/temp.jpg')


a=pd.read_excel(open('/home/roadburner/Desktop/contacts.xlsx','rb'),sheet_name="Sheet1", header=None)

for i in range(0,len(a)):
    writeText(a[0][i])
    send_mail('YourEmailAddress@gmail.com',a[1][i], 'Do Not Reply', 'XXX', files=['/home/roadburner/Desktop/temp.jpg'])
