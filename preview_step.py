import win32com.client
import os


def ready_html_email(name,area,step):
    path = os.path.join(os.getcwd(),"steps",f"{step}.txt")
    body = open(path,"r").read().format(name=name,area=area)
    outlook = win32com.client.Dispatch("Outlook.Application")
    accounts = outlook.Session.Accounts
    mail = outlook.CreateItem(0)
    for acc in accounts:
        if acc.DisplayName == "janindu@bubltown.co.uk":
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, acc))
            break
    mail.Subject = "test"
    signature = open("signature.txt","r").read().format(name=name,email="janindu@bubltown.co.uk")
    
    mail.HTMLBody = body
    mail.HTMLBody += signature

    mail.To = ";".join(["janidul03@gmail.com"])
    mail.Display(True)
    
try:
    step = input("Enter step to preview: ")
except FileNotFoundError:
    print("File does not exsist")
ready_html_email("Janindu","wallignton",step)