import os
import win32com.client
from multiprocessing import Process
import time
import pandas as pd
from datetime import datetime
import random
import csv

#to do figure out including the area in to the subject DONE

class EmailSender:
    def __init__(self,name,batch_no,email_address,subject:str,step,cc_recipients=None, bcc_recipients=None) -> None:
        self.file_path = None
        self.batch_no = batch_no
        self.email_address = email_address
        self.cc_recipients = cc_recipients
        self.bcc_recipients = bcc_recipients
        self.subject = subject       
        self.step = step
        self.path= os.getcwd()
        self.name = name
    
    def prepare_batch_file(self):
        """
        reads the excel batch for relavent for each batch number
        
        Returns:
            df: a pandas dataframe
        """
        df =  pd.read_excel(self.get_file_location())
        return df
    
    def prepare_signature_html_file(self):
        """opens the signature text file and changes the placeholder values
        
        Returns:
            str: text in html format for the email signature with place holder values for name and
            senders email address filled
        """
        path = os.path.join(self.path,"signature.txt")
        signature = open(path,"r").read().format(name=self.name,email=self.email_address)
        return signature
    
    def prepare_subject(self,area):
        """passing subject if area should be added to subject it should be added by replacing {} to area"""

        if "{}" in self.subject:
            subject = self.subject
            subject = subject.replace("{}",area)
            return subject
        else:
            return self.subject
    
    def get_file_location(self):
        """get the file location of the batch file according to the 
        batch no
        Returns:
            str: file location 
        """
        self.file_path = os.path.join(self.path,"batches",f"{self.batch_no}.xlsx")
        return self.file_path


    
    def preparing_template(self):
        """preparing the text template file

        Returns:
            str: str of the text template
        """
        path = os.path.join(self.path,"steps",f"{self.step}.txt")
        template_text = open(path,"r").read()
        return template_text
    
    def randmoizer(self,start,stop):
        """get a random number between start and stop 
        
        Args:
            start (int): start number
            stop (int): stop number

        Returns:
            int:a random number with in the range of start and stop
        """
        return random.randrange(start,stop)
    
    def send_email_from_account(self, subject, body, to_recipients,signature):
        outlook = win32com.client.Dispatch("Outlook.Application")
        accounts = outlook.Session.Accounts
        mail = outlook.CreateItem(0)
        for acc in accounts:
            if acc.DisplayName == self.email_address:
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, acc))
                break
            
        mail.Subject = subject
        mail.HTMLBody = body
        mail.HTMLBody += signature
        
        mail.To = ";".join(to_recipients)
        if self.cc_recipients:
            mail.CC = ";".join(self.cc_recipients)
        if self.bcc_recipients:
            mail.BCC = ";".join(self.bcc_recipients)
        mail.Send()
    
    def make_directories(self,batch_no):
        path = os.path.join(self.path,"batch_output",batch_no)
        if  not os.path.exists(path):
            try:
                os.makedirs(path)
            except OSError as e:
                print(f"Error creating directory {path}: {e}")
        
    def sending_emails(self):
        template_text = self.preparing_template()
        df = self.prepare_batch_file()
        for index, row in df.iterrows():
            name = row['Name']  
            to = [str(row['Email']).strip()]  
            area = row['Area'] 
            body = template_text.format(name=name,area=area)
            subject = self.prepare_subject(area)
            signature = self.prepare_signature_html_file()

            self.send_email_from_account(signature=signature,body=body,to_recipients=to,subject=subject)
            time_now = datetime.now().strftime('%H:%M:%S')
            date_now= datetime.now().date()
            print(f"Email:{to} Time:{time_now} sending from:{self.email_address}")
            
            self.make_directories(self.batch_no)
            
            path = os.path.join(self.path,"batch_output",self.batch_no,f"step_{self.step}_output.csv")
            with open(path,'a', newline='') as csv_file:
                csv_writer = csv.writer(csv_file)
                csv_writer.writerow((to[0],self.email_address,time_now,self.step,self.batch_no,self.subject,date_now))
            csv_file.close()
            
            sleep=self.randmoizer(20,30)
            time.sleep(sleep)
        

    