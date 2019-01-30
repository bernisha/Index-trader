# -*- coding: utf-8 -*-
"""
Created on Thu Aug  9 11:17:11 2018

@author: blala
"""


def input_tloader(termi_nate_cnt=5):
    cnt=0
    loop=True
    while loop:
        d1a = input ("1. Load trades into Decalog: 1) Yes. 2) No.[Y/N]?: ")
        
        if d1a=="Y":
            print ("Trade load in progress",end='', flush=True)
            break
        elif d1a=="N":
            print ("Please exit batch trade",end='', flush=True)
            break
        else:
            cnt=cnt+1
            #print(cnt)
            print("Invalid input, please select the correct option")
            if cnt==termi_nate_cnt:
                print("You have run out of options, default option selected")
                d1a='N'
                #loop=False
                #print(loop)
                break
    
    if loop:
        x = [d1a]

    return x    

input_tloader()    