import Issue_Status as i
import schedule
import time
from datetime import datetime

def schedule_job():
    print('\n \n \n \n')
    print('Scheduled at 6 PM: Please check with Murali before closing this...')
    
    def job():
        print('Start of TD Java Testing Status Generation')
        i.issues_Status_Mail()
        print('End  of TD Java Testing Status Generation')        
        now = datetime.now()
        dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
        print('status sending completed... at {}'.format(dt_string))
        print('\n \n *********************************************** \n \n ')
        
    schedule.every().day.at("18:00").do(job)
    #job()

    
    while True:
        schedule.run_pending()
        time.sleep(60)
        
if __name__ == "__main__":
    schedule_job()

