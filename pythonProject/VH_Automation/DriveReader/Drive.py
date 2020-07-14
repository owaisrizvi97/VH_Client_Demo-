from pydrive.auth import GoogleAuth
import os
from pydrive.drive import GoogleDrive
import pandas as pd
import time
import json
from datetime import datetime
import pyodbc 
import shutil

gauth = GoogleAuth()
gauth.LocalWebserverAuth()
drive = GoogleDrive(gauth)

def copyFiles(src, dest):
    shutil.copy2(src, dest) 

def readFolder(parent):
    FolderNames= []
    emailFolderInfo=[]
    file_list = drive.ListFile({'q': "'%s' in parents and trashed=false" % parent}).GetList()
    for f in file_list:
        if f['mimeType']=='application/vnd.google-apps.folder': # if folder
            FolderNames.append(  {"id":f["id"],"title":f["title"] ," email":  f["owners"][0]["emailAddress"]})
            
    return FolderNames

def readFiles(parent):
    Files = []
    file_list = drive.ListFile({'q': "'%s' in parents and trashed=false" % parent}).GetList()
    for f in file_list:
        if f['mimeType']=='application/vnd.google-apps.folder': # if folder
            Files+= readFiles(f['id'])
            
        else:
            
            Files.append({ "id":f["id"], "parentId":f['parents'][0]['id'] ,"title":f["title"],"email":f["owners"][0]["emailAddress"], "status: ": 0, "CreatedDate":f["createdDate"] })            
                
    return Files
def downloadFiles(fileListId, fileListTitle, directoryName):
    from pathlib import Path
    from random import randint
    if Path(files['title']).is_file():
        return False,0,0,0
        # file6 = drive.CreateFile({'id':  fileListId})
        # name=os.path.splitext(files['title'])
        # filename = name[0]+"_"+str(randint(1,900))+name[1]
        # file6.GetContentFile(filename)
        
    else:
        try:
            # Create target Directory
            os.mkdir(directoryName)
            print("Directory " , directoryName ,  " Created ") 
        except FileExistsError:
            print("Directory " , directoryName ,  " already exists")
        '''if bool(fileListTitle):
            filePath = os.path.join('D:/task1/'+directoryName, fileListTitle)
            if not os.path.isfile(filePath) :
                fp = open(filePath, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()'''
        file6 = drive.CreateFile({'id':  fileListId})
        file6.GetContentFile(fileListTitle)
        
        
    
    name=os.path.splitext(files['title'])
    
    
    return True,os.path.getsize(files['title']),name[0],name[1]


# def trash(fileListId):
#     fileA = drive.CreateFile({'id':  fileListId})
#     fileA.Trash()

#     fileA.FetchMetadata(fields='permissions,labels,mimeType')
#     print(fileA['permissions'])
    # The permissions are also available as file1['permissions']:
    

def insert(dict):
    emailedBy, Filename, InsertedBy,Source = dict['Email'],dict['Title']+dict['Extension'],'Ali','Drive'
    client,domain = dict['Client'],dict['Domain']
    conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=DESKTOP-P707RUQ\SQLEXPRESS;'
                      'Database=VH_Automation;'
                      'Trusted_Connection=yes;')
    cur = conn.cursor()
    cur.execute("INSERT INTO  [dbo].[automation]([EmailedBy],[FileName],[RunBy],[Source],[Client],[Domain]) VALUES(?,?,?,?,?,?)",( emailedBy, Filename, InsertedBy,Source,client,domain) )
    conn.commit()
    conn.close()

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




summary=[]
AllFiles = readFiles('root')
Folders = readFolder('root')

for files in AllFiles:
    
    client_name = files['email'].split('@')
    #print('Printing the names', client_name[0])
    dirName = client_name[0]
    for parent in Folders:
        if parent['id'] == files['parentId']:
            company = parent['title']
        else:
            continue
    status,size,fileName,extension, = downloadFiles(files['id'],files['title'],dirName)   
    if status:
        summary.append({'Company': company, 'Title':fileName,'Extension': extension, 'Email' : files['email'],'Client':client_name[0],'Domain':client_name[1] ,'File Size': size ,'Uploaded (by Company)':files['CreatedDate'] ,'status': status,   })
    # trash(files['id'])
    else:
        continue


for data in summary:
    insert(data)

df = pd.DataFrame(summary)

df.to_excel("Summary.xlsx")

# sendNotification(df)
path = 'c:/Users/User/Desktop/pythonProject/VH_Automation/DriveReader/'
for data in summary:
   fname = data['Title'] + data['Extension']
   dest = path + data['Client'] 
   src = path + fname
   copyFiles(src, dest)