import pandas as pd
import schedule
import time
from datetime import datetime
from multiprocessing import Process
from pre_run import getting_input
from main import EmailSender
from datetime import datetime


# Function to execute at a specific date and time
def execute_functions(date, time_value, name, email_address, subject, step, batch_no):
    instances= []
    start_time = datetime.now()
    now = start_time.strftime('%H:%M:%S')
    instance= EmailSender(name=name,batch_no=batch_no, email_address=email_address, subject=subject, step=step)
    instances.append(instance)
    processes = [Process(target=inst.sending_emails) for inst in instances]

    for process in processes:
        process.start()
    for process in processes:
        process.join()
    end_time = datetime.now()
    end = end_time.strftime('%H:%M:%S')
    time_difference = end_time - start_time
    print(f"Time difference: {time_difference}")
    print("All functions completed")

    

# Function to read the CSV file and schedule the execution of functions
def read_csv_and_schedule(csv_path):
    df = pd.read_csv(csv_path)

    for index, row in df.iterrows():
        date = row['date']
        time_value = row['time']
        name = row['name']
        email_address = row['email address']
        subject = row['subject']
        step = row['step']
        batch_no = row['batch no']

        # Convert date and time columns to a datetime object
        entry_datetime = datetime.strptime(f"{date} {time_value}", "%m/%d/%Y %H:%M")

        # Schedule the execution of functions at the specified date and time
        schedule.every().day.at(entry_datetime.strftime('%H:%M')).do(
            execute_functions, date, time_value, name, email_address, subject, step, batch_no
        )

# Replace 'your_csv_file.csv' with the actual path to your CSV file
csv_file_path = 'shedule.csv'
read_csv_and_schedule(csv_file_path)

# Main loop to keep the script running and check for scheduled tasks
while True:
    schedule.run_pending()
    time.sleep(1)
