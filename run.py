from multiprocessing import Process
from pre_run import getting_input
from main import EmailSender
from datetime import datetime

start_time = datetime.now()
now = start_time.strftime('%H:%M:%S')

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
    
    
    
    