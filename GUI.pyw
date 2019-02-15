# -*- coding: utf-8 -*-
"""
Created on Thu Jan 31 14:27:41 2019

@author: BLala
"""

#from tkinter import Tk, Label, Button
#
#
#
#
#class OMCSIdxTrd:
#    def __init__(self, master):
#        self.master = master
#        master.title("OMCS IndexTrader")
#
#        self.label = Label(master, text="Welcome to Indexation Trading hub!")
#        self.label.pack()
#
#        self.greet_button = Button(master, text="Futures Report", command=self.greet)
#        self.greet_button.pack()
#        
#
#        self.close_button = Button(master, text="Close", command=root.quit)
#        self.close_button.pack()
#
#    def greet(self):
#        print("Greetings!")
#
#root = Tk()
#my_gui = OMCSIdxTrd(root)
#root.mainloop()

import tkinter
import tkinter.messagebox
from tkinter import ttk

import sys
import tkinter
import tkinter.messagebox
    #import tkinter.ttk
    
from PIL import ImageTk, Image
    #from write_excel import input_fx as inp
    
from futures_calc_fx import fut_calc_func as fut_calc
from write_excel  import tloader_fmt_futures as load_fut
from write_excel  import create_BPMcashfile as cash_bpm  
from pre_flow_calc_fx import pre_flow_calcFx as batch_calc_fx
from write_excel import tloader_fmt_equity as tloader_equity_or_fut


class OMGCS_Index_gui:

 
    def __init__(self, window):
        self.window = window
        window.geometry("500x500+500+300")
        window.title("OMCS IndexTrader")
        
        self.label = tkinter.Label(window, text="Welcome to OMCS Indexation hub!",font=("Helvetica", 14))
        self.label.pack()
        
        self.icon =  ImageTk.PhotoImage(Image.open("c:/IndexTrader/images/index.jpg").resize((60,60),Image.ANTIALIAS))  
        label2 = tkinter.Label(window, image = self.icon)
        label2.pack()
   
        self.omig =  ImageTk.PhotoImage(Image.open("c:/IndexTrader/images/omig.jpg").resize((98,42),Image.ANTIALIAS))  
        #self.omig.zoom(2,2)
        label_o = tkinter.Label(window, image = self.omig)
        label_o.place(relx=0.12, rely=0.93, anchor="c")
        
        self.cs =  ImageTk.PhotoImage(Image.open("c:/IndexTrader/images/CS_Indexation.jpg").resize((98,42),Image.ANTIALIAS))  
        #self.omig.zoom(2,2)
        label_p = tkinter.Label(window, image = self.cs)
        label_p.place(relx=0.89, rely=0.93, anchor="c")
       
     
        self.text_btn = tkinter.Button(window, text = "1. Generate Futures Report!", command = self.fut_report) # create a button to call a function called 'say_hi'
        self.text_btn.pack()

        self.text_btnF = tkinter.Button(window, text = "2. Load Futures into Decalog!", command = self.load_fut) # create a button to call a function called 'say_hi'
        self.text_btnF.pack()

        self.text_btnB = tkinter.Button(window, text = "3. Generate Batch Cash calc!", command = self.batch_report) # create a button to call a function called 'say_hi'
        self.text_btnB.pack()

        self.text_btnC = tkinter.Button(window, text = "4. Create BPM cash & Futures file!", command = self.cashforBPM) # create a button to call a function called 'say_hi'
        self.text_btnC.pack()

        
        self.text_btnO = tkinter.Button(window, text = "5. Drop Post-Opt Files to Folder!", command = self.cashforBPM) # create a button to call a function called 'say_hi'
        self.text_btnO.pack()

        self.text_btnP = tkinter.Button(window, text = "6. Load Batch trades into Decalog!", command = self.load_trades) # create a button to call a function called 'say_hi'
        self.text_btnP.pack()

        
        
        
        self.progress = ttk.Progressbar(window, orient="horizontal",
                                        length=200, mode="determinate")
        self.progress.place(relx=0.5, rely=0.8, anchor="c")
        
        self.bytes = 0
        self.maxbytes = 0    
        
        self.close_btn = tkinter.Button(window, text = "Close", command = self.on_closing)# closing the 'window' when you click the button
        #self.close_btn.pack()
        self.close_btn.place(relx=0.5, rely=0.9, anchor="c")
    
    def start(self):
        self.progress["value"] = 0
        self.maxbytes = 5000
        self.progress["maximum"] = 5000
        self.fut_report()
        self.batch_report()
        self.load_fut()
        self.cashforBPM()
        self.load_trades()
        
        
# Create the futures report
    def fut_report(self):
        self.progress["value"] = 0
        #tkinter.messagebox.showinfo("Are flows & cash limits up to date: 1) Yes. 2) No.[Y/N]?:")
        response = tkinter.messagebox.askquestion("Flows", "Are flows & cash limits up to date?")
      #  print(response)
        lbl=tkinter.Label(window, text = " \n \n \n" ,fg='grey', font=("Helvetica", 14), width =50)
        lbl.place(relx=0.5, rely=0.6, anchor="c")
# If user clicks 'Yes' then it returns 1 else it returns 0
        if response == 'yes':
          #   def read_bytes(self):
              '''simulate reading 500 bytes; update progress bar'''
              self.bytes += 500
              self.progress["value"] = self.bytes
              if self.bytes < self.maxbytes:
            # read more bytes after 100 ms
                 self.after(100, self.fut_report)

              lbl=tkinter.Label(window, text = "Futures report generated!", fg='green', font=("Helvetica", 14), bg='white')
              lbl.place(relx=0.5, rely=0.6, anchor="c")
        
        else:
            lbl=tkinter.Label(window, text = "Please Update Flows", fg='red', font=("Helvetica", 14), bg='white')
            lbl.place(relx=0.5, rely=0.6, anchor="c")
        if response=='yes':
            fut_calc(response)
            #print("yes")
          
# Create the load for futures
            
    def load_fut(self):
        self.progress["value"] = 0
        '''simulate reading 500 bytes; update progress bar'''
        self.bytes += 500
        self.progress["value"] = self.bytes
        lbl=tkinter.Label(window, text = " \n \n \n" ,fg='grey', font=("Helvetica", 14), width =50)
        lbl.place(relx=0.5, rely=0.6, anchor="c")
        g=load_fut()
        
        if self.bytes < self.maxbytes:
    # read more bytes after 100 ms
        
             self.after(100, self.load_fut)
        lbl=tkinter.Label(window, text = g, fg='green', font=("Helvetica", 10), bg='white')
        lbl.place(relx=0.5, rely=0.6, anchor="c")
        #print("yes")
        
        
## Create the Batch Report
    def batch_report(self):
        self.progress["value"] = 0
        #tkinter.messagebox.showinfo("Are flows & cash limits up to date: 1) Yes. 2) No.[Y/N]?:")
        response = tkinter.messagebox.askquestion("Flows", "Are flows & cash limits up to date?")
       # print(response)
        lbl=tkinter.Label(window, text = " \n \n \n" ,fg='grey', font=("Helvetica", 14), width =50)
        lbl.place(relx=0.5, rely=0.7, anchor="c")
# If user clicks 'Yes' then it returns 1 else it returns 0
        if response == 'yes':
          #   def read_bytes(self):
              '''simulate reading 500 bytes; update progress bar'''
              self.bytes += 500
              self.progress["value"] = self.bytes
              if self.bytes < self.maxbytes:
            # read more bytes after 100 ms
                 self.after(100, self.batch_report)

              lbl=tkinter.Label(window, text = "Batch cash calc generated!" ,fg='green', font=("Helvetica", 14), bg='white')
              lbl.place(relx=0.5, rely=0.6, anchor="c")
        else:
            lbl=tkinter.Label(window, text = "Please Update Flows",fg='red', font=("Helvetica", 14), bg='white')
            lbl.place(relx=0.5, rely=0.6, anchor="c")
        if response=='yes':
            # runfile('C:/IndexTrader/code/pre_flow_calc.py', wdir='C:/IndexTrader/code')               
            batch_calc_fx(response)
            # print("yes")
       
        #tkinter.Label(window, text = "Futures report in progress").pack()
       # runfile('C:/IndexTrader/code/futures_calc.py', wdir='C:/IndexTrader/code')

# Create cash file
    def cashforBPM(self):
        self.progress["value"] = 0
        '''simulate reading 500 bytes; update progress bar'''
        lbl=tkinter.Label(window, text = " \n \n \n" ,fg='grey', font=("Helvetica", 10), width =50)
        lbl.place(relx=0.5, rely=0.6, anchor="c")

        b=cash_bpm()
    #    tkinter.Label(window, text = b).pack()
        self.bytes += 500
        self.progress["value"] = self.bytes
        if self.bytes < self.maxbytes:
    # read more bytes after 100 ms
             self.after(100, self.cash_bpm)
        lbl=tkinter.Label(window, text = b,fg='green', font=("Helvetica", 10), bg='white')
        lbl.place(relx=0.5, rely=0.6, anchor="c")     

# Load trades
    def load_trades(self):
        self.progress["value"] = 0
        '''simulate reading 500 bytes; update progress bar'''
        lbl=tkinter.Label(window, text = " \n \n \n" ,fg='grey', font=("Helvetica", 10), width =50)
        lbl.place(relx=0.5, rely=0.6, anchor="c")

        d=tloader_equity_or_fut("y")
    #    tkinter.Label(window, text = b).pack()
        self.bytes += 500
        self.progress["value"] = self.bytes
        if self.bytes < self.maxbytes:
    # read more bytes after 100 ms
             self.after(100, self.tloader_equity_or_fut)
        lbl=tkinter.Label(window, text = d,fg='green', font=("Helvetica", 10), bg='white')
        lbl.place(relx=0.5, rely=0.6, anchor="c") 
             
    def on_closing(self):
        import os
        if tkinter.messagebox.askokcancel("Quit", "Do you want to quit?"):
            window.destroy()
            #sys.modules[__name__].__dict__.clear()
            os._exit(00)
            
window = tkinter.Tk()
#window.title("GUI")

geeks_bro = OMGCS_Index_gui(window)


 

#close_btn = tkinter.Button(window, text = "Close", command = window.destroy)# closing the 'window' when you click the button
        #self.close_btn.pack()
#close_btn.place(relx=0.5, rely=0.9, anchor="c")
 
#def on_closing():
#    if tkinter.messagebox.askokcancel("Quit", "Do you want to quit?"):
#        window.destroy()
#
#window.protocol("WM_DELETE_WINDOW", on_closing)
    
window.mainloop()
#window.withdraw()

#sys.exit()



#import tkinter
#import tkinter.messagebox
#
#window = tkinter.Tk()
#window.title("GUI")
#
## creating a simple alert box
#tkinter.messagebox.showinfo("Alert Message", "This is just a alert message!")
## creating a question to get the response from the user [Yes or No Question]
#response = tkinter.messagebox.askquestion("Simple Question", "Do you love Python?")
## If user clicks 'Yes' then it returns 1 else it returns 0
#if response == 1:
#    tkinter.Label(window, text = "You love Python!").pack()
#else:
#    tkinter.Label(window, text = "You don't love Python!").pack()
#
#window.mainloop()
#
#from tkinter import *
#from PIL import ImageTk, Image
#root = Tk()
#
#canv = Canvas(root, width=80, height=80, bg='white')
#canv.grid(row=2, column=3)
#
#img = ImageTk.PhotoImage(Image.open("bll.jpg"))  # PIL solution
#canv.create_image(20, 20, anchor=NW, image=img)
#
#mainloop()
#
#photo= tkinter.PhotoImage(file = "c:/IndexTrader/images/index.png")
#
#from PIL import ImageTk, Image
#img = ImageTk.PhotoImage(Image.open("c:/IndexTrader/images/index.jpg"))  
#l=tkinter.Label(image=img)
#l.pack()
#
#mainloop()













