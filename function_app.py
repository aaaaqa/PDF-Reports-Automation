import logging
from datetime import datetime
import azure.functions as func
from main import final 

app = func.FunctionApp()

@app.function_name(name="myTimer")
@app.timer_trigger(schedule="0 0 6 * * *", arg_name="myTimer", run_on_startup=True, use_monitor=False)
def func_AutomationPDF(myTimer: func.TimerRequest) -> None:
    timestamp = datetime.today().strftime('Y-%m-%d %H:%M:%S')
    if myTimer.past_due:
        logging.info('The timer is past due!')
    final()
    logging.info('Python timer trigger function executed at %s', timestamp)

#@app.schedule(schedule="0 0 6 * * *", arg_name="myTimer", run_on_startup=True, use_monitor=False) 
#def timer_trigger(myTimer: func.TimerRequest) -> None:
#    if myTimer.past_due:
#        logging.info('The timer is past due!')
#    logging.info('Python timer trigger function executed.')