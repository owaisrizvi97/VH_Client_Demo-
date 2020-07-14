
from django.shortcuts import render, redirect
import os
import pandas as pd
def home(request):
    return render(request,'home.html')

def scriptDrive(request):
    print(os.getcwd())
    
    os.system("python Script.py")
    
    from_dest = {"From":"DRIVE"}
    
    return render(request,"displayDrive.html",from_dest)


def scriptMail(request):
    print(os.getcwd())
    from_dest = {"From":"MAIL"}
    os.system("python Email_Script.py")
    return render(request,"displayDrive.html",from_dest)