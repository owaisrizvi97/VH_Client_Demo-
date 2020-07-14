import pandas as pd
import datetime
def sendInternalNoti(df):
        
    import smtplib, ssl
    from email.mime.multipart import MIMEMultipart

    from email.mime.text import MIMEText
    
    port = 465  # For SSL
    smtp_server = "smtp.gmail.com"
    sender_email = "vectorhealthtest@gmail.com"  # Enter your address
    receiver_email = 'vh.pharma.client@gmail.com'  # Enter receiver address
    password = '459139403341895266'
   
    cc = "alisiddiqui1997@hotmail.com"
    

    str_msg ="Comapny \t File Size\t   TimeStamp\t\t\t\t\t\t Filename \n"
    company= df['Company']
    filename= df['Title']+df['Extension']
    email = df['Email']
    file_size=df['File Size']
    
    for i in range(len(company)):
        str_msg = str_msg + str(company[i])+ '  \t '+str(file_size[i]) + '  \t '+ str(datetime.datetime.today()) +' \t\t\t  '+str(filename[i]) +" \n"
    print(str_msg)
    
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
print('Internal Email sent')


