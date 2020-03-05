import smtplib
import xlrd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
file_location = "C:\\Users\\Aritra\\Desktop\\Certificates\\CSe.xlsx"
e=input('Enter No. of Pages')
for i in range(int(e)):

        workbook = xlrd.open_workbook(file_location)
        sheet = workbook.sheet_by_name('CSE Responses')
        x = []
        rownum=2
        for value in sheet.col_values(4):
            if isinstance(value, str):
                x.append(value)
        
        email_user = 'srmistwalkathon@gmail.com'
        email_password = 'certificate1234'
        email_send = x[i+1]
        subject = 'Walkathon 2K19 - Participation(Corrected)'
    
        msg = MIMEMultipart()
        msg['From'] = 'srmistwalkathon@gmail.com'
        msg['To'] = x[i+1]
        msg['Subject'] = 'Walkathon 2K19 - Participation'

        body = 'Hi Participant\nThank You for participating in Walkathon 2019 , a venture of SRM Institute of Science and Technology Ramapuram as a part of Texus.\nWe have attached a certificate as a gesture of gratitude!\n\nHope you have a nice day.\n\n The last email sent to you might have the wrong certificate attached to it. It was a malfunction in our mailbot. We are sorry for the inconvienience caused to you from our end. We have attached the corrent once this time.'
        msg.attach(MIMEText(body,'plain'))

        filename='page-'+str(i+1)+'.pdf'
        attachment  =open(filename,'rb')

        part = MIMEBase('application','octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',"attachment; filename= "+filename)

        msg.attach(part)
        text = msg.as_string()
        server = smtplib.SMTP('smtp.gmail.com',587)
        server.starttls()
        server.login(email_user,email_password)
        server.sendmail(email_user,email_send,text)
        server.quit()
        print('Sent to email '+x[i+1])
    
