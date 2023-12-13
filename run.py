from multiprocessing import Process
from pre_run import getting_input
from main import EmailSender


if __name__ == '__main__':
    instances = getting_input()
    processes = [Process(target=inst.sending_emails) for inst in instances]
   
    for process in processes:
        process.start()
    
    for process in processes:
        process.join()
    
    print("All functions completed")
    
    
    