import os
import win32com.client
from multiprocessing import Process
import time
import pandas as pd
from datetime import datetime
import random
import csv

class EmailSender:
    def __init__(self,batch_no,email_address,subject:str,step,cc_recipients=None, bcc_recipients=None) -> None:
        self.file_path = None
        self.batch_no = batch_no
        self.email_address = email_address
        self.cc_recipients = cc_recipients
        self.bcc_recipients = bcc_recipients
        self.subject = subject       
        self.step = step
        self.path= os.getcwd()
    
    def prepare_batch_file(self):
        df =  pd.read_excel(self.get_file_location())
        return df
    
    def prepare_signature_html_file(self,name):
        path = os.path.join(self.path,"signature.txt")
        signature = open(path,"r").read().format(name)
        return signature
    
    def prepare_text_html_file(self):
        pass
    
    def prepare_subject(self,area):
        """passing subject if area should be added to subject it should be added by replacing {} to area"""
        try:
            if "{}" in self.subject:
                self.subject = self.subject.replace("{}",area)
                return self.subject
        except:
            return self.subject
    
    def get_file_location(self):
        self.file_path = os.path.join(self.path,"batches",f"{self.batch_no}.xlsx")
        return self.file_path

    def send_email_from_account(self, subject, body, to_recipients):
        outlook = win32com.client.Dispatch("Outlook.Application")
        accounts = outlook.Session.Accounts
        mail = outlook.CreateItem(0)

        for acc in accounts:
            if acc.DisplayName == self.email_address:
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, acc))
                break

        mail.Subject = self.subject
        mail.Body  = body
        mail.To = ";".join(to_recipients)

        if self.cc_recipients:
            mail.CC = ";".join(self.cc_recipients)

        if self.bcc_recipients:
            mail.BCC = ";".join(self.bcc_recipients)

        mail.Send()
    
    def preparing_template(self):
        path = os.path.join(self.path,"steps",f"{self.step}.txt")
        template_text = open(path,"r").read()
        return template_text
    
    def read_html_file(self):
        path = os.path.join(self.path,"steps",f"{self.step}.html")
        with open(path, 'r') as file:
            html_content = file.read()
        return html_content

    def randmoizer(self,start,stop):
        return random.randrange(start,stop)
    
    def sending_emails(self):
        template_text = self.preparing_template()
        df = self.prepare_batch_file()
        for index, row in df.iterrows():
            name = row['Name']  
            to = [str(row['Email']).strip()]  
            area = row['Area'] 
            body = template_text.format(name,area)
            subject = self.prepare_subject(area)

            self.send_email_from_account(body=body,to_recipients=to,subject=subject)
            time_now = datetime.now().strftime('%H:%M:%S')
            date_now= datetime.now().date()
            print(f"Email:{to} Time:{time_now} sending from:{self.email_address}")
            
            path = os.path.join(self.path,"batch_output",f"{self.batch_no}_output.csv")
            with open(path,'a', newline='') as csv_file:
                csv_writer = csv.writer(csv_file)
                csv_writer.writerow((to[0],self.email_address,time_now,self.step,self.batch_no,self.subject,date_now))
            csv_file.close()
            
            sleep=self.randmoizer(20,30)
            time.sleep(sleep)
        

        
    