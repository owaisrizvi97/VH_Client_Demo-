
import imaplib
import base64
import os
import email
import smtplib
import pandas as pd
import pyodbc 
from datetime import datetime


def sendNotification(df):

        
    import smtplib, ssl
    
    companyEmail = pd.unique(df['Email'])
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    port = 465  # For SSL
    smtp_server = "smtp.gmail.com"
    sender_email = "vectorhealthcompliance@gmail.com"  # Enter your address
    
    password = '459139403341895266'

    
    
    context = ssl.create_default_context()
    for email in companyEmail:
    
        print('Sending Email to ' , email)
        grouped = df.groupby('Email').get_group(email)
        str_msg = "recieved {}  files ".format(grouped['Title'].count())
        
        print(grouped)
        email2 = email.split('@')
        message = """\
        Subject: VH-Files recieved from """+email2[0]+"""

        We have recieved your files: \n""" + str(grouped['Title'].to_string()).split('   ')[1] +'      '+str(grouped['File Size'].to_string()).split('   ')[1] +'      '+str(datetime.now())
    
        with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
            server.login(sender_email, password)
        
            server.sendmail(sender_email, email, message)






def check(date):
    conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=DESKTOP-P707RUQ\SQLEXPRESS;'
                      'Database=VH_Automation;'
                      'Trusted_Connection=yes;')
    cur = conn.cursor()
    cur.execute("select * from [dbo].[automation] where [DateUploaded] = (?) ", (date) )
    rows = cur.fetchall()

    conn.commit()
    conn.close()
    print(len(rows))
    if len(rows) > 0:
        return False
    
    return True


def insert(dict):
    emailedBy, DateTime, Filename, InsertedBy = dict['EmailedBy'],dict['Date'],dict['File'],dict['RunBy']
    file_size = dict['File Size']
    ext = dict['Extension']
    company = dict["Email"].split('.')
    client,domain =dict['Company'],dict['Domain']
    
    conn = pyodbc.connect('Driver={SQL Server};'
                        'Server=DESKTOP-P707RUQ\SQLEXPRESS;'
                        'Database=VH_Automation;'
                        'Trusted_Connection=yes;')
    cur = conn.cursor()
    cur.execute("INSERT INTO  [dbo].[automation]([EmailedBy],[FileName],[RunBy],[Source],[Client],[Domain],[Company],[FileExtension],[FileSize],[DateUploaded]) VALUES(?,?,?,?,?,?,?,?,?,?)",( emailedBy ,Filename,InsertedBy , 'Email',client,domain,company[0],ext,file_size,dict['Date'] ) )
    conn.commit()
    conn.close()
   


email_user = 'vectorhealthcompliance@gmail.com'
email_pass = '459139403341895266'
mail = imaplib.IMAP4_SSL('imap.gmail.com',993)

mail.login(email_user, email_pass)
allMail =  mail.select('Inbox')
print(allMail)
# type, data = mail.search(None, 'ALL') 
# type, data = mail.search(None, 'UNSEEN') 
type, data = mail.search(None, 'ALL') 

mail_ids = data[0]
summary=[]



for num in data[0].split():
    typ, data = mail.fetch(num, '(RFC822)' )
    raw_email = data[0][1]
# converts byte literal to string removing b''
    raw_email_string = raw_email.decode('utf-8')
    email_message = email.message_from_string(raw_email_string)
    print(email_message['Date'])
    From = email_message['From']
    Date = email_message['Date']
    print("from:  " ,From)
    emailedby=From.split('<')
    try: 
        email_c = emailedby[1].split('>')
        client_domain = email_c[0].split('@')
                
        dirName = From.split('@')
        if check(Date):
            for part in email_message.walk():
                
                
                if part.get_content_maintype() == 'multipart':

                    continue
                if part.get('Content-Disposition') is None:
                    continue
                
                fileName = part.get_filename()
                ext = fileName.split('.')
                ext = ext[1]
                
                
                
                if  email_c[0]=='pharmaaclient@gmail.com':
                    if bool(fileName):
                        folder =client_domain[0].split('.')
                        filePath = os.path.join('C:/Users/User/Desktop/pythonProject/VH_Automation/DriveReader/VectorHealth/VH_SpendView_Server/'+ 'VH_Pharma_Client' , fileName)
                        if not os.path.isfile(filePath) :
                            fp = open(filePath, 'wb')
                            fp.write(part.get_payload(decode=True))
                            fp.close()
                        file_size = os.path.getsize(filePath)
                else:
                    if bool(fileName):
                        folder =client_domain[0].split('.')
                        filePath = os.path.join('C:/Users/User/Desktop/pythonProject/VH_Automation/DriveReader/VectorHealth/VH_SpendView_Server/'+ 'VH_Pharma_Client' , fileName)
                        if not os.path.isfile(filePath) :
                            fp = open(filePath, 'wb')
                            fp.write(part.get_payload(decode=True))
                            fp.close()
                        file_size = os.path.getsize(filePath)

                    
                Title = fileName.split('.')
                ext = Title[1]
                Title = Title[0]
                summary.append({ "EmailedBy":emailedby[0],"Email": email_c[0], "Domain": client_domain[1] , "Company" : client_domain[0],"Date":Date  ,"File":fileName,"Title": Title,"Extension" : ext,"RunBy":"Ali","Extension":ext,"File Size": file_size })
                
                dirName = dirName[1].split('.')
                dirName = dirName[0]
    
    except:
        print('exception')
        
    
        # try:
        #     # Create target Directory
        #     os.mkdir(dirName)
        #     print("Directory " , dirName ,  " Created ") 
        # except FileExistsError:
        #     print("Directory " , dirName ,  " already exists")
        
        # if bool(fileName):
        #     filePath = os.path.join('C:/Users/User/Desktop/pythonProject/VH_Automation/EmailReader/'+dirName, fileName)
        #     if not os.path.isfile(filePath) :
        #         fp = open(filePath, 'wb')
        #         fp.write(part.get_payload(decode=True))
        #         fp.close()
            
            # print('Downloaded "{file}" from email titled "{subject}" with UID .'.format(file=fileName, subject=subject))


df = pd.DataFrame(summary)
for data in summary:
    insert(data)
df.to_excel("EmailData.xlsx")



import internalEmail
if df.empty:
    print('DataFrame is empty!')
else:
    sendNotification(df)
    # internalEmail.sendInternalNoti(df)
