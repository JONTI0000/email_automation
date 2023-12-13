from multiprocessing import Process
from main import EmailSender
import re
import os
path = os.getcwd()

def validate_email(email):
    pattern = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    return re.match(pattern, email)

def get_valid_email():
    while True:
        email = input("Enter the Email: ").strip()
        if validate_email(email):
            return email
        else:
            print("Invalid email. Please enter a valid email.")

def get_valid_step():
    global path
    dirs = os.listdir(os.path.join(path,"steps"))
    dirs = [dir.replace(".txt","") for dir in dirs ]
    while True:
        step = input("Enter the Step: ").strip()
        if step in dirs:
            return step
        else:
            print("Invalid Step,Enter a valid Step")

def get_valid_batch():
    global path
    dirs = os.listdir(os.path.join(path,"batches"))
    dirs = [dir.replace(".xlsx","") for dir in dirs ]
    while True:
        batch_no = input("Enter batch no: ").strip()
        if batch_no in dirs:
            return batch_no
        else:
            print("Invalid Step,Enter a valid Step")

def get_valid_email_count():
     while True:
        count = input("How many Email addresses: ").strip()
        try:
            count = int(count)
            return count
        except:
            print("Not a valid count")
        

def getting_input():
    count = get_valid_email_count()
    instances = []
    input_count = 0
    while input_count < count:
        if input_count < count:
            email = get_valid_email()
            step = get_valid_step()
            subject = input("Enter the Subject: ").strip()
            batch_no = get_valid_batch()
            instance= EmailSender(batch_no=batch_no, email_address=email, subject=subject, step=step)
            instances.append(instance)
            input_count += 1
    return instances

"""
if __name__ == '__main__':
   processes = [Process(target=inst.sending_emails) for inst in instances]
   
   for process in processes:
       process.start()
   
   for process in processes:
       process.join()
   
   print("All functions completed")"""