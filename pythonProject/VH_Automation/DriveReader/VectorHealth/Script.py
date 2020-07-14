from pydrive.auth import GoogleAuth
import os
from pydrive.drive import GoogleDrive
import pandas as pd
import time
import json
from datetime import datetime
import pyodbc 
import shutil

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
def downloadFiles(fileListId, fileListTitle, directoryName,CreatedDate):
    from pathlib import Path
    from random import randint
    
    chk = check_file_download(fileListTitle,CreatedDate)
    name=os.path.splitext(files['title'])
    if chk:
        path = 'C:/Users/User/Desktop/pythonProject/VH_Automation/DriveReader/VectorHealth/VH_SpendView_Server/' + directoryName + '/' 
        if os.path.isfile(path+fileListTitle):
        
            file6 = drive.CreateFile({'id': fileListId })
            file6.GetContentFile(fileListTitle)
            fnumb=0
            
            
            filename_toCopy = path+fileListTitle
            
            while(os.path.isfile(filename_toCopy)):
                filename_toCopy = fileListTitle.split('.')
                last_name = filename_toCopy[1] 
                filename_toCopy =filename_toCopy[0]
                filename_toCopy = filename_toCopy.split('(')
                filename_toCopy = filename_toCopy[0] +'(' + str(fnumb)+')'
            print(filename_toCopy)
            shutil.copy2(fileListTitle, path+filename_toCopy +'.'+last_name)
            
            return True,os.path.getsize(files['title']),filename_toCopy,name[1]
        else:
            file6 = drive.CreateFile({'id': fileListId })
            file6.GetContentFile(fileListTitle)
            filename_toCopy = fileListTitle.split('.')
            CreatedDate = CreatedDate.split(':')
            print('Saving file: ')
            print(path+filename_toCopy[0]+CreatedDate[0]+filename_toCopy[1])
            
            shutil.copy2(fileListTitle, path+filename_toCopy[0]+"_"+CreatedDate[0]+'.'+filename_toCopy[1])
            return True,os.path.getsize(files['title']),name[0],name[1]

        # file6 = drive.CreateFile({'id':  fileListId})
        # name=os.path.splitext(files['title'])
        # filename = name[0]+"_"+str(randint(1,900))+name[1]
        # file6.GetContentFile(filename)
        
    else:
        # try:
        #     # Create target Directory
        #     os.mkdir('/Company/'+directoryName)
        #     print("Directory " , directoryName ,  " Created ") 
        # except FileExistsError:
        #     print("Directory " , directoryName ,  " already exists")
        
        path = 'C:/Users/User/Desktop/pythonProject/VH_Automation/DriveReader/VectorHealth/VH_SpendView_Server/' + directoryName + '/'
        print(path)
        file6 = drive.CreateFile({'id': fileListId })
        file6.GetContentFile(fileListTitle)
        shutil.copy2(fileListTitle, path+fileListTitle)
        # elif(company == 'Regeneron'):
        #     path = 'C:/Users/User/Desktop/pythonProject/VH_Automation/DriveReader/VectorHealth/Companies/Regeneron/'
        #     file6 = drive.CreateFile({'id': fileListId })
        #     file6.GetContentFile(fileListTitle)
        #     shutil.copy2(fileListTitle, path+'/'+fileListTitle)
        # else:
        #     path = 'C:/Users/User/Desktop/pythonProject/VH_Automation/DriveReader/VectorHealth/Companies/Oakwood/'
        #     file6 = drive.CreateFile({'id': fileListId })
        #     file6.GetContentFile(fileListTitle)
        #     shutil.copy2(fileListTitle, path+'/'+fileListTitle)
            
    
        
    
    return False,os.path.getsize(files['title']),name[0],name[1]

def delFile(fileListTitle):
    os.remove(fileListTitle )

# def trash(fileListId):
#     fileA = drive.CreateFile({'id':  fileListId})
#     fileA.Trash()

#     fileA.FetchMetadata(fields='permissions,labels,mimeType')
#     print(fileA['permissions'])
    # The permissions are also available as file1['permissions']:
    

def insert(dict):
    emailedBy, Filename, InsertedBy,Source = dict['Email'],dict['Title']+dict['Extension'],'Ali','Drive'
    client,domain = dict['Client'],dict['Domain']
    company = dict['Company']
    ext =dict['Extension']
    uploadedC_date = dict['Uploaded (by Company)']
    file_size = dict['File Size']

    print('INSERT CHECK:')
    chk = check_file_download(Filename,uploadedC_date)
    
    print(chk)
    
    if chk:
        conn = pyodbc.connect('Driver={SQL Server};'
                        'Server=DESKTOP-P707RUQ\SQLEXPRESS;'
                        'Database=VH_Automation;'
                        'Trusted_Connection=yes;')
        cur = conn.cursor()
        cur.execute("INSERT INTO  [dbo].[automation]([EmailedBy],[FileName],[RunBy],[Source],[Client],[Domain],[Company],[FileExtension],[FileSize],[DateUploaded]) VALUES(?,?,?,?,?,?,?,?,?,?)",( emailedBy, Filename, InsertedBy,Source,client,domain, company,ext,file_size, uploadedC_date ) )
        conn.commit()
        conn.close()

def sendNotification(df):
        
    import smtplib, ssl
    
    companyEmail = pd.unique(df['Email'])
    print(companyEmail)
    print(df['Email'])
    port = 465  # For SSL
    smtp_server = "smtp.gmail.com"
    sender_email = "vectorhealthcompliance@gmail.com"  # Enter your address
    
    password = '459139403341895266'

    
    
    context = ssl.create_default_context()
    for email in companyEmail:
        print('Sending Email to ' , email)
        grouped = df.groupby('Email').get_group(email)
        # titles =str(grouped['Title'])
        title_list = []
        for title in grouped['Title']:
            
            if len(title)>2:
                title_list.append(title)
        
        size_list = []
        count = []
        i=1
        timestamp=[]
        email2 = email.split('@')
        for sizeList in grouped['File Size']:
            
            if (sizeList)>2:
                size_list.append(sizeList)
                count.append(i)
                timestamp.append(str(datetime.now()))
                i = i +1

        
        

        df = pd.DataFrame([size_list,timestamp, title_list], index = ['','','' ],columns=count )
        df= df.T
       

        message = """\
        Subject: VH-Files recieved from """+email2[0]+"""

        We have recieved your files: \n""" + str(df)
        
        with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
            server.login(sender_email, password)
        
            server.sendmail(sender_email, email, message)




gauth = GoogleAuth()
gauth.LocalWebserverAuth()
drive = GoogleDrive(gauth)
check_sum = []
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
            status,size,fileName,extension, = downloadFiles(files['id'],files['title'],company,files['CreatedDate'])   
            delFile(files['title'])
            if status:
                dict_to_add={'Company': company, 'Title':fileName,'Extension': extension, 'Email' : files['email'],'Client':client_name[0],'Domain':client_name[1] ,'File Size': size ,'Uploaded (by Company)':files['CreatedDate'] ,'status': status,   } 
                insert(dict_to_add)
                summary.append(dict_to_add)
        
        else:
            continue
    # status,size,fileName,extension, = downloadFiles(files['id'],files['title'],company,files['CreatedDate'])   
    # delFile(files['title'])
    # if status:
    #     dict_to_add={'Company': company, 'Title':fileName,'Extension': extension, 'Email' : files['email'],'Client':client_name[0],'Domain':client_name[1] ,'File Size': size ,'Uploaded (by Company)':files['CreatedDate'] ,'status': status,   } 
    #     insert(dict_to_add)
    #     summary.append(dict_to_add)
        
    # trash(files['id'])
    

    
import internalEmail
df = pd.DataFrame(summary)

sendNotification(df)
# internalEmail.sendInternalNoti(df)
df.to_excel("Summary.xlsx")


# path = 'c:/Users/User/Desktop/pythonProject/VH_Automation/DriveReader/'
# for data in summary:
#    fname = data['Title'] + data['Extension']
#    dest = path + data['Client'] + fname
#    src = path + fname
#    copyFiles(src, dest)

# import internalEmail
# if df.empty:
#     print('DataFrame is empty!')
# else:
#     internalEmail.sendInternalNoti(df)