from pydrive.auth import GoogleAuth
import os
from pydrive.drive import GoogleDrive
import pandas as pd
import time
import json
from datetime import datetime
import pyodbc 
import shutil
import imaplib
import base64
import os
import email
import smtplib


def sendNotification(str_msg):
        
    import smtplib, ssl
    from email.mime.multipart import MIMEMultipart

    from email.mime.text import MIMEText
    
    port = 465  # For SSL
    smtp_server = "smtp.gmail.com"
    sender_email = "vectorhealthtest@gmail.com"  # Enter your address
    receiver_email = 'vectorhealthtest@gmail.com'  # Enter receiver address
    password = '459139403341895266'
   
    cc = "alisiddiqui1997@hotmail.com"
    

    # str_msg ="Comapny \t File Size\t  Filename \t\t  \n"
    # company= df['Company']
    # filename= df['Title']+df['Extension']
    # email = df['Email']
    # file_size=df['File Size']
    # for i in range(len(company)):
    #     str_msg = str_msg + str(company[i])+ '  \t '+str(file_size[i]) + '  \t '+ str(filename[i])+' \t\t\t  ' +" \n"
    # print(str_msg)
    
    context = ssl.create_default_context()
    
        
    message = """\
        
        Subject: VH-Files (recieved)

        {}.""".format(str_msg)
        
    
    

    
    msg = MIMEMultipart('alternative')
    msg['Subject'] = "VectorHealth"
    msg['To'] = receiver_email
    msg['Cc'] = cc
    body = MIMEText(message)
    msg.attach(body)
    with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
        server.login(sender_email, password)
        
        server.sendmail(sender_email, receiver_email, msg.as_string())






def check_file_download(title,cdate):
    print(title,cdate)
    conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=DESKTOP-P707RUQ\SQLEXPRESS;'
                      'Database=VH_Automation;'
                      'Trusted_Connection=yes;')
    cur = conn.cursor()
    cur.execute("select * from [dbo].[automation] where  [FileName]=(?)", (title ) )
    rows = cur.fetchall()
    cur.execute("select * from [dbo].[automation] where  [DateUploaded]=(?)", (cdate ) )
    rows2 = cur.fetchall()
    conn.commit()
    conn.close()
    
    
    if(len(rows) == len(rows2)):
        if len(rows) > 0:
            return False
    
    return True




def readFiles(parent):
    Files = []
    file_list = drive.ListFile({'q': "'%s' in parents and trashed=false" % parent}).GetList()
    for f in file_list:
        if f['mimeType']=='application/vnd.google-apps.folder': # if folder
            Files+= readFiles(f['id'])
            
        else:
            
            Files.append({ "id":f["id"], "parentId":f['parents'][0]['id'] ,"title":f["title"],"email":f["owners"][0]["emailAddress"], "status: ": 0, "CreatedDate":f["createdDate"] })            
                
    return Files


def get_email():
    email_user = 'vectorhealthtest@gmail.com'
    email_pass = '459139403341895266'
    mail = imaplib.IMAP4_SSL('imap.gmail.com',993)
    mail.login(email_user, email_pass)
    allMail =  mail.select('Inbox')
    type, data = mail.search(None, 'ALL')
    mail_ids = data[0]
    summary_msg = "EMAIL_SUMMARY\nTitle \t\t\t\t\t\t Date \t\t\t\t\t\t Status\n"
    for num in data[0].split():
        typ, data = mail.fetch(num, '(RFC822)' )
        raw_email = data[0][1]
    # converts byte literal to string removing b''
        raw_email_string = raw_email.decode('utf-8')
        email_message = email.message_from_string(raw_email_string)
        Date = email_message['Date']
        for part in email_message.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue 
            fileName = part.get_filename()
            summary_msg = summary_msg + fileName + '\t\t\t\t'+ Date + '\t\t\t\t' + str(not check_file_download(fileName,Date))+'\n'
    sendNotification(summary_msg)
                




gauth = GoogleAuth()
gauth.LocalWebserverAuth()
drive = GoogleDrive(gauth)
AllFiles = readFiles('root')
files_in_drive = []
for data in AllFiles:
    files_in_drive.append({'Title':data['title'], 'Date':data['CreatedDate'] })

df = pd.DataFrame(files_in_drive)
status =[]
summary_msg = "DRIVE_SUMMARY\nTitle \t\t\t\t\t\t Date \t\t\t\t\t\t Status\n"
for i in range(len(df)):
    title,date=df.iloc[i]
    summary_msg = summary_msg + title + '\t\t\t\t'+ date + '\t\t\t\t' + str(not check_file_download(title,date))+'\n'
print(summary_msg)

sendNotification(summary_msg)



get_email()