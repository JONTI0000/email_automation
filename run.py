from multiprocessing import Process
from pre_run import getting_input
from main import EmailSender
from datetime import datetime
import schedule
from sheduler import getting_input,load_data,finding_closest_date,get_the_time
start_time = datetime.now()
now = start_time.strftime('%H:%M:%S')

"""df = load_data()
date = finding_closest_date(df)
time = get_the_time(df,date)



if __name__ == '__main__':
    instances = getting_input()
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
    
    
    """

import schedule
import time
from datetime import datetime

def my_task():
    if __name__ == '__main__':
        instances = getting_input()
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

def schedule_task():
    df = load_data()
    date = finding_closest_date(df)
    times = get_the_time(df,date)
    time_dt = datetime.strptime(times,"%H:%M")

    # Schedule the task to run at a specific day and time
    scheduled_date = datetime.combine(date,time_dt.time())
    schedule.every().day.at(scheduled_date.strftime("%H:%M")).do(my_task)

    # Loop to continuously check the current date and time
    while True:
        # Check if the scheduled date and time is reached
        if datetime.now() >= scheduled_date:
            # Run the task
            my_task()
            break  # Exit the loop once the task is executed
        
        # Sleep for a short duration before checking again
        time.sleep(60)  # Sleep for 60 seconds (adjust as needed)

# Start scheduling the task
schedule_task()
