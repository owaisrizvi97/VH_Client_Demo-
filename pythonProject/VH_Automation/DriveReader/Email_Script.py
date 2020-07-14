import imaplib
import base64
import os
import email
import smtplib
import pandas as pd
import pyodbc 

def sendNotification(df):
        
    import smtplib, ssl
    companyEmail = pd.unique(df['Email'])
    
    port = 465  # For SSL
    smtp_server = "smtp.gmail.com"
    sender_email = "vectorhealthtest@gmail.com"  # Enter your address
    receiver_email = companyEmail  # Enter receiver address
    password = '459139403341895266'

    
    
    context = ssl.create_default_context()
    for email in companyEmail:
        grouped = df.groupby('Email').get_group(email)
        str = "recieved {}  files ".format(grouped['Title'].count())
        message = """\
        Subject: VH-Files (recieved)

        We have {}.""".format(str)
    
        with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
            server.login(sender_email, password)
        
            server.sendmail(sender_email, email, message)




def insert(dict):
    emailedBy, DateTime, Filename, InsertedBy = dict['EmailedBy'],dict['Date'],dict['File'],dict['RunBy']
    client,domain =dict['Client'],dict['Domain']
    conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=DESKTOP-P707RUQ\SQLEXPRESS;'
                      'Database=VH_Automation;'
                      'Trusted_Connection=yes;')
    cur = conn.cursor()
    cur.execute("INSERT INTO  [dbo].[automation]([EmailedBy],[FileName],[RunBy],[Source],[Client],[Domain]) VALUES(?,?,?,?,?,?)",( emailedBy ,Filename,InsertedBy , 'Email',client,domain) )
    conn.commit()
    conn.close()


email_user = 'vectorhealthtest@gmail.com'
email_pass = '459139403341895266'
mail = imaplib.IMAP4_SSL('imap.gmail.com',993)

mail.login(email_user, email_pass)
allMail =  mail.select('Inbox')
print(allMail)
type, data = mail.search(None, 'ALL')
mail_ids = data[0]
summary=[]



for num in data[0].split():
    typ, data = mail.fetch(num, '(RFC822)' )
    raw_email = data[0][1]
# converts byte literal to string removing b''
    raw_email_string = raw_email.decode('utf-8')
    email_message = email.message_from_string(raw_email_string)
    From = email_message['From']
    Date = email_message['Date']
    for part in email_message.walk():
        # this part comes from the snipped I don't understand yet... 
        
        if part.get_content_maintype() == 'multipart':

            continue
        if part.get('Content-Disposition') is None:
            continue
        
        fileName = part.get_filename()

        emailedby=From.split('<')
        email_c = emailedby[1].split('>')
        client_domain = email_c[0].split('@')
        summary.append({ "EmailedBy":emailedby[0],"email": email_c[0], "Domain": client_domain[1] , "Client" : client_domain[0],"Date":Date  ,"File":fileName,"RunBy":"Ali" })
        # dirName = From.split('<')
        dirName = From.split('@')
        dirName = dirName[1].split('.')
        dirName = dirName[0]
        try:
            # Create target Directory
            os.mkdir(dirName)
            print("Directory " , dirName ,  " Created ") 
        except FileExistsError:
            print("Directory " , dirName ,  " already exists")
        if bool(fileName):
            filePath = os.path.join('C:/Users/User/Desktop/pythonProject/VH_Automation/EmailReader/'+dirName, fileName)
            if not os.path.isfile(filePath) :
                fp = open(filePath, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()
            
            # print('Downloaded "{file}" from email titled "{subject}" with UID .'.format(file=fileName, subject=subject))


df = pd.DataFrame(summary)
for data in summary:
    
    insert(data)
df.to_excel("EmailData.xlsx")



# sendNotification(df)
       



