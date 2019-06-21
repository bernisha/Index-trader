# -*- coding: utf-8 -*-
"""
Created on Thu May  9 09:34:47 2019

@author: BLala
"""

#import pyautogui 
import time as tm
from time import sleep  
from pre_flow_calc_fx_cc import pre_flow_calcFx_cc as clr_cash  
from write_excel  import create_BPMcashfile as cash_bpm 
from write_excel import clear_cash_fx_drop as cash_drop


import shutil
import os
    

from datetime import datetime, timedelta
startDate = (datetime.today()).strftime("%d-%m-%Y %H-%M-%S")


source = 'C:\\IndexTrader\\auto_cash_BPM_load\\'
dest1 = 'C:\\IndexTrader\\required_inputs\\'
files = os.listdir(source)
files=['flows.csv']
for f in files:
        print(f)
        shutil.copy(source+f, dest1)


run_time = tm.time()  
#pyautogui.keyDown('win')
#pyautogui.press('r')
#pyautogui.keyUp('win')
#pyautogui.typewrite("C:\\IndexTrader\\auto_cash_BPM_load\\flows.csv\n") #cmd to open command prompt
##pyautogui.PAUSE = 10.0
#pyautogui.moveTo(611, 429)
#
#pyautogui.click()
#pyautogui.hotkey('alt', 'f','a','o') 
#pyautogui.typewrite("C:\\IndexTrader\\required_inputs\\flows.csv\n") #cmd to open command prompt
#pyautogui.press('y')
#pyautogui.hotkey('alt', 'f','c') 

with open(str("C:/IndexTrader/auto_cash_BPM_load/Logs/log_cc_"+str(startDate)+'.txt'),"w",newline='\r\n') as fle:
    
    fle.write("Inflows file updated \n")
    # Create batch file
    try:
        clr_cash(response='yes',automatic=False,orders=False,testing=False,clear_cash=True)
        fle.write("Cash calculated in batch cash calc \n")
    # Create cash file to drop to Listener folder
        cash_bpm(clear_cash=True)
        fle.write("Cash  file created\n")
    # Drop cash file to listner folder
        cash_drop()
        fle.write("Drop cash file tp Listener \n")
        msg="Cash cleared in BPM"
    except:
        msg="Error in clearing of cash BPM"


    fle.write("--- %s seconds ---" % (tm.time() - run_time))
    fle.write(str("\n"+msg))
    fle.close()

