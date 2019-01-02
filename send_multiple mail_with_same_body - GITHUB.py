
import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import time

#creating loop for picking mail one by one


for x in range(3,221):  #FROM WHICH ROW TO WHICH ROW EMAIL ID'S ARE PRESENT # in this case 3 to 220 (221-1) 
    wb = openpyxl.load_workbook('workbook.xlsx')# FILE NAME 
    # sheet = wb.get_sheet_by_name('Sheet1') #SHEET NAME#
    sheet = wb['Sheet1']

    print(x) #PRINTING ID NAME FOR INFORMATION
    mail_1 = sheet.cell(row=x, column=1).value # FATCHING EMAIL ID'S ONE BY ONE
    email_send = mail_1
    print(mail_1)
    email_user = 'USER_NAME' #PUT YOUR GMAIL USER NAME HERE
    email_password = 'PASSWORD' #PUT YOUR PASSWORD HERE


    subject = 'SUBJECT OF MAIL' #PUT SUBJECT OF MAIL HERE

    msg = MIMEMultipart()
    msg['From'] = email_user
    msg['To'] = email_send
    msg['Subject'] = subject


    body = """
                <HTML>

                 #PASTE YOUR BODY TEXT IN HTML FORMAT HERE
                 # IF YPU DON'T KNOW HTML , DON'T WORRY.
                 # TYPE IN MS-WORD DOCUMENT AND SAVE THAT AS HTML DOC
                 #OPEN THAT HTML FILE IN NOTEPAD AND COPY PASTE IT IN BETWEEN HTML TAGS 
                       
</HTML>
                """

    


    msg.attach(MIMEText(body,'html'))

    filename = 'file_name to be attached'
    attachment = open(filename, 'rb')

    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= " + filename)

    msg.attach(part)
    text = msg.as_string()
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(email_user, email_password)
    server.sendmail(email_user, email_send, text)
    time.sleep(10) # PROVIDING REST OF 10SEC BETWEEN TO MAILS
server.quit()
