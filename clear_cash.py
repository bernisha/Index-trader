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
import configparser
from pathlib import Path

base_path = Path(__file__).parent
file_path = (base_path / "../config/config_clearcash.ini").resolve()
print(file_path)

Config = configparser.ConfigParser()
Config
#<ConfigParser.ConfigParser instance at 0x00BA9B20>
Config.read(file_path)
#Config.sections()

#Config.options('section1')[1]



from datetime import datetime, timedelta
startDate = (datetime.today()).strftime("%d-%m-%Y %H-%M-%S")


source = Config.get('section0',Config.options('section0')[0]) # path_to_flows_input
dest1 = Config.get('section0',Config.options('section0')[1]) # path_to_flows_output

config_path= Config.get('section0',Config.options('section0')[3]) 
config_path_cash=Config.get('section0',Config.options('section0')[4]) 
output_batch=Config.get('section0',Config.options('section0')[5]) 
IT_fold=Config.get('section0',Config.options('section0')[7]) 

user_dict=Config.get('section1',Config.options('section1')[1]) 
fnd_dict=Config.get('section1',Config.options('section1')[2]) 
csh_lmt_file=Config.get('section1',Config.options('section1')[3]) 
fnd_excp_list=Config.get('section1',Config.options('section1')[4]).split(",") 
vba_bin=Config.get('section0',Config.options('section0')[6])
files = os.listdir(source)
files=[Config.get('section1',Config.options('section1')[0])]

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

with open(str(Config.get('section0',Config.options('section0')[2])+'log_cc_'+str(startDate)+'.txt'),"w",newline='\r\n') as fle:
        
    fle.write("Inflows file updated \n")
    # Create batch file
    try:
        clr_cash(config_path,config_path_cash, output_batch,user_dict,fnd_dict,csh_lmt_file,fnd_excp_list,vba_bin,req_input_direc=dest1,
                 flows=files[0],response='yes',automatic=False,orders=True,testing=False,clear_cash=True)
        fle.write("Cash calculated in batch cash calc \n")
    # Create cash file to drop to Listener folder
        cash_bpm(fnd_excp=fnd_excp_list,clear_cash=True,user_file=str(dest1+user_dict),dir_out=output_batch)
        fle.write("Cash  file created\n")
    # Drop cash file to listner folder
        cash_drop(dir_imp=output_batch,lis_fld=IT_fold,user_file=str(dest1+user_dict))
        fle.write("Drop cash file tp Listener \n")
        msg="Cash cleared in BPM"
    except:
        msg="Error in clearing of cash BPM"


    fle.write("--- %s seconds ---" % (tm.time() - run_time))
    fle.write(str("\n"+msg))
    fle.close()

