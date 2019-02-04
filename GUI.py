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
#import tkinter.ttk

from PIL import ImageTk, Image
#from write_excel import input_fx as inp
from tkinter import ttk

from futures_calc_fx import fut_calc_func as fut_calc

class OMGCS_Index_gui:

    def __init__(self, window):
        self.window = window
        window.geometry("400x400+500+300")
        window.title("OMCS IndexTrader")
        
        self.label = tkinter.Label(window, text="Welcome to Indexation Trading hub!")
        self.label.pack()
        
        self.icon =  ImageTk.PhotoImage(Image.open("c:/IndexTrader/images/index.jpg").resize((50,50),Image.ANTIALIAS))  
        label2 = tkinter.Label(window, image = self.icon)
        label2.pack()
   
     
        self.text_btn = tkinter.Button(window, text = "1. Generate Futures Report!", command = self.fut_report) # create a button to call a function called 'say_hi'
        self.text_btn.pack()

        self.text_btnB = tkinter.Button(window, text = "2. Generate Batch Cash calc!", command = self.batch_report) # create a button to call a function called 'say_hi'
        self.text_btnB.pack()

        self.progress = ttk.Progressbar(window, orient="horizontal",
                                        length=200, mode="determinate")
        self.progress.place(relx=0.5, rely=0.8, anchor="c")
        
        self.bytes = 0
        self.maxbytes = 0    
        
        self.close_btn = tkinter.Button(window, text = "Close", command = self.window.destroy)# closing the 'window' when you click the button
        #self.close_btn.pack()
        self.close_btn.place(relx=0.5, rely=0.9, anchor="c")
    
    def start(self):
        self.progress["value"] = 0
        self.maxbytes = 5000
        self.progress["maximum"] = 5000
        self.fut_report()
        self.batch_report()

    def fut_report(self):
        #tkinter.messagebox.showinfo("Are flows & cash limits up to date: 1) Yes. 2) No.[Y/N]?:")
        response = tkinter.messagebox.askquestion("Flows", "Are flows & cash limits up to date?")
        print(response)
# If user clicks 'Yes' then it returns 1 else it returns 0
        if response == 'yes':
          #   def read_bytes(self):
              '''simulate reading 500 bytes; update progress bar'''
              self.bytes += 500
              self.progress["value"] = self.bytes
              if self.bytes < self.maxbytes:
            # read more bytes after 100 ms
                 self.after(100, self.fut_report)

              tkinter.Label(window, text = "Futures report generated!").pack()
        else:
            tkinter.Label(window, text = "Please Update Flows").pack()
        if response=='yes':
            fut_calc(response)

    def batch_report(self):
        #tkinter.messagebox.showinfo("Are flows & cash limits up to date: 1) Yes. 2) No.[Y/N]?:")
        response = tkinter.messagebox.askquestion("Flows", "Are flows & cash limits up to date?")
        print(response)
# If user clicks 'Yes' then it returns 1 else it returns 0
        if response == 'yes':
          #   def read_bytes(self):
              '''simulate reading 500 bytes; update progress bar'''
              self.bytes += 500
              self.progress["value"] = self.bytes
              if self.bytes < self.maxbytes:
            # read more bytes after 100 ms
                 self.after(100, self.batch_report)

              tkinter.Label(window, text = "Batch cash calc generated!").pack()
        else:
            tkinter.Label(window, text = "Please Update Flows").pack()
        if response=='yes':
            fut_calc(response)                 
       
        #tkinter.Label(window, text = "Futures report in progress").pack()
       # runfile('C:/IndexTrader/code/futures_calc.py', wdir='C:/IndexTrader/code')
        

window = tkinter.Tk()
#window.title("GUI")

geeks_bro = OMGCS_Index_gui(window)

#close_btn = tkinter.Button(window, text = "Close", command = window.destroy)# closing the 'window' when you click the button
        #self.close_btn.pack()
#close_btn.place(relx=0.5, rely=0.9, anchor="c")
 
window.mainloop()


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













