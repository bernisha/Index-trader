# -*- coding: utf-8 -*-
"""
Created on Wed Jun  5 09:18:25 2019

@author: BLala
"""

import pyautogui as paut


paut.position() 

paut.size()

paut.locateOnScreen('IndexTrader.lnk')  


import pyautogui
screenWidth, screenHeight = pyautogui.size()
currentMouseX, currentMouseY = pyautogui.position()


pyautogui.moveTo(1551, 224)
pyautogui.click()

pyautogui.keyDown('alt')
pyautogui.press(' ')
pyautogui.press('n')
pyautogui.keyUp('alt')

for i in range(25):
    print(i)
    pyautogui.hotkey('alt',' ','n','alt')
    pyautogui.hotkey('win','r')
    


pyautogui.hotkey('command', 'l')

pyautogui.keyDown('command')
pyautogui.press('d')
pyautogui.keyUp('command')


#pyautogui.keyDown('command')
#pyautogui.press('d')
#pyautogui.keyUp('command')


pyautogui.moveTo(1480, 250)
pyautogui.doubleClick()
pyautogui.moveTo(1480, 250)


pyautogui.PAUSE = 10.0

pyautogui.keyDown('win')
pyautogui.press('r')
pyautogui.keyUp('win')



pyautogui.click()
pyautogui.moveRel(None, 10)  # move mouse 10 pixels down
pyautogui.doubleClick()
pyautogui.moveTo(500, 500, duration=2, tween=pyautogui.easeInOutQuad)  # use tweening/easing function to move mouse over 2 seconds.
pyautogui.typewrite('Hello world!', interval=0.25)  # type with quarter-second pause in between each key
pyautogui.press('esc')
pyautogui.keyDown('shift')
pyautogui.press(['left', 'left', 'left', 'left', 'left', 'left'])
pyautogui.keyUp('shift')
pyautogui.hotkey('ctrl', 'c')



import pyautogui 
from time import sleep    #import sleep - for the delay</p><p>
pyautogui.hotkey('win', 'r')  #windows_key + Run</p><p>
#pyautogui.typewrite("cmd\n") #cmd to open command prompt
pyautogui.keyDown('win')
pyautogui.press('r')
pyautogui.keyUp('win')
pyautogui.typewrite("C:\\IndexTrader\\auto_cash_BPM_load\\flows.csv\n") #cmd to open command prompt
pyautogui.PAUSE = 10.0
pyautogui.moveTo(611, 429)

pyautogui.click()
pyautogui.hotkey('alt', 'f','a','o') 
pyautogui.typewrite("C:\\IndexTrader\\required_inputs\\flows.csv\n") #cmd to open command prompt
pyautogui.press('y')
pyautogui.hotkey('alt', 'f','c') 

pyautogui.hotkey('win', 'r')  #windows_key + Run</p><p>

pyautogui.keyDown('win')
pyautogui.press('r')
pyautogui.keyUp('win')

pyautogui.typewrite("S:\DFM\Tools\Indexation_trading_tools\IndexTrader\GUI\IndexTrader.exe\n") #cmd to open command prompt
pyautogui.moveTo(611, 429)
#pyautogui.PAUSE = 10.0
pyautogui.click()

sleep(0.500)       #500 milisecond delay (depends on your computer speed)</p><p><br>#write the code then press enter('\n') thus pc will auto-lock</p><p>pyautogui.typewrite("rundll32.exe user32.dll, LockWorkStation\n")


