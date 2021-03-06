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

import os
import sys
import tkinter
import tkinter.messagebox
from tkinter import ttk
import webbrowser

#import sys
#import tkinter
#import tkinter.messagebox
    #import tkinter.ttk
    
from PIL import ImageTk, Image
    #from write_excel import input_fx as inp
    
from futures_calc_fx import fut_calc_func as fut_calc
from write_excel  import tloader_fmt_futures as load_fut
from write_excel  import create_BPMcashfile as cash_bpm  
from pre_flow_calc_fx import pre_flow_calcFx as batch_calc_fx
from write_excel import tloader_fmt_equity as tloader_equity_or_fut
from write_excel import BPM_output_loads as list_lds

def resource_path(relative_path):
     if hasattr(sys, '_MEIPASS'):
         return os.path.join(sys._MEIPASS, relative_path)
     return os.path.join(os.path.abspath("."), relative_path)

#resource_path("c:/IndexTrader/images/index.jpg")

class OMGCS_Index_gui:

 
    def __init__(self, window):
        
        self.window = window
        window.geometry("500x600+500+300")
        window.title("OMCS IndexTrader")
        window.resizable(0, 0)
        
        #self.y_pos =0.785
        #self.x_pos =0.5
        self.y_pos =0.19
        self.x_pos =0.60
        
        self.pos_x=0.6
        self.pos_y=0.4
        
        
        self.wl=200
        
        self.label = tkinter.Label(window, text="OMCS INDEXATION TRADING HUB",font=("Courier 14 bold"))
        self.label.place(relx=0.50, rely=0.05, anchor="c")
        
        if getattr(sys, 'frozen', False):
            # we are running in a bundle
            #bundle_dir = sys._MEIPASS
            bundle_dir='C:/IndexTrader/code'
        else:
            # we are running in a normal Python environment
            #bundle_dir = os.path.dirname(os.path.abspath(__file__))
            bundle_dir='C:/IndexTrader/code'
		
        
        #     self.icon =  ImageTk.PhotoImage(Image.open("c:/IndexTrader/images/index.jpg").resize((60,60),Image.ANTIALIAS))  
        self.icon =  ImageTk.PhotoImage(Image.open(bundle_dir + "/images/index.jpg").resize((50,50),Image.ANTIALIAS))  
        label2 = tkinter.Label(window, image = self.icon)
        label2.grid(row=1, column=5, rowspan=2,columnspan=1, sticky='ew', padx=10, pady=2)
   
        self.icon1 =  ImageTk.PhotoImage(Image.open(bundle_dir + "/images/index.jpg").resize((50,50),Image.ANTIALIAS))  
        label2_1 = tkinter.Label(window, image = self.icon1)
        label2_1.place(relx=0.1, rely=0.05, anchor="c")
   
        self.omig =  ImageTk.PhotoImage(Image.open(bundle_dir + "/images/omig.jpg").resize((98,42),Image.ANTIALIAS))  
        #self.omig.zoom(2,2)
        label_o = tkinter.Label(window, image = self.omig)
        label_o.place(relx=0.12, rely=0.93, anchor="c")
        
        self.cs =  ImageTk.PhotoImage(Image.open(bundle_dir + "/images/CS_Indexation.jpg").resize((98,42),Image.ANTIALIAS))  
        #self.omig.zoom(2,2)
        label_p = tkinter.Label(window, image = self.cs)
        label_p.place(relx=0.89, rely=0.93, anchor="c")
        
    
# Futures Trading
    
        self.labelframe = tkinter.LabelFrame(window, text=" Futures Trading ",  bd=3,relief=tkinter.RIDGE, font="Courier 12 bold")
        self.labelframe.grid(column=0, row=4,columnspan=4, sticky='ew',padx=5, pady=5,ipadx=0, ipady=0)
       # self.top = tkinter.Label(self.labelframe, text="")
       # self.top.pack()
     
        self.text_btn = tkinter.Button(self.labelframe, text = "     Generate Futures Report    ", wraplength= self.wl, command = self.fut_report) # create a button to call a function called 'say_hi'
        self.text_btn.grid(column=0, row=5, sticky='ew',padx=5, pady=5)
        
        top = tkinter.Label(self.labelframe , text="ggggggggggggggggggggg \n   \n  ",  fg='SystemButtonFace', bg='SystemButtonFace')
        top.grid(column=3, row=4,rowspan=3, columnspan=1, sticky='ew', padx=25, pady=5)

    
        self.text_btnF = tkinter.Button(self.labelframe, text = "    Load Futures into Decalog     ",wraplength= self.wl, command = self.load_fut) # create a button to call a function called 'say_hi'
        self.text_btnF.grid(column=0, row=6, sticky='ew',padx=5, pady=5)

# Batch Trading
        
        self.labelframeB = tkinter.LabelFrame(window, text="Batch Equity Trading",  bd=3,relief=tkinter.RIDGE, font="Courier 12 bold")
        self.labelframeB.grid(column=0, row=7,columnspan=4, sticky='ew',padx=5, pady=5,ipadx=0, ipady=0)
       #

        self.text_btnB = tkinter.Button(self.labelframeB, text = "Generate Batch Cash Calc", wraplength= self.wl,command = self.batch_report) # create a button to call a function called 'say_hi'
        self.text_btnB.grid(column=0, row=8, sticky='ew',padx=5, pady=5)

        self.text_btnC = tkinter.Button(self.labelframeB, text = "Create BPM Cash & Futures file",wraplength= self.wl, command = self.cashforBPM) # create a button to call a function called 'say_hi'
        self.text_btnC.grid(column=0, row=9, sticky='ew',padx=5, pady=5)

        
        self.text_btnO = tkinter.Button(self.labelframeB, text = "Drop Post-Opt Files to Folder",wraplength= self.wl, command = self.lst_lod) # create a button to call a function called 'say_hi'
        self.text_btnO.grid(column=0, row=10, sticky='ew',padx=5, pady=5)
        self.flag= True

      #  self.bot = tkinter.Label(window, text=" ",  fg='SystemButtonFace', bg='SystemButtonFace')
      #  self.bot.grid(column=2, row=7,rowspan=5, columnspan=1, sticky='ew', padx=0, pady=5)
  
      
        #if self.flag:
        self.text_btnP = tkinter.Button(self.labelframeB, text = "Load Batch trades into Decalog", wraplength= self.wl,command = self.load_trades) # create a button to call a function called 'say_hi'
        self.text_btnP.grid(column=0, row=11, sticky='ew',padx=5, pady=5)


# Download Frame
        self.labelframeC = tkinter.LabelFrame(window, text="Select trades to upload:", bd=0,bg='SystemButtonFace', fg='SystemButtonFace',  font=("Courier 9 bold"))
        self.labelframeC.grid(column=0, row=13,columnspan=2, rowspan= 5, sticky='ew',padx=5, pady=5)
        
        dem = tkinter.Label(self.labelframeC, text="ggggg\ngggg\nggggggg\nggg\ngg\n",  fg='SystemButtonFace', bg='SystemButtonFace')
        dem.grid(column=1, row=13,rowspan=5, padx=0, pady=0)
      
        self.labelframeT = tkinter.LabelFrame(window, text="Download Templates",  bd=3,relief=tkinter.RIDGE, font="Courier 10 bold", fg= "dark blue", highlightcolor='pink')
        self.labelframeT.grid(column=3, row=19,padx=0, pady=0,ipadx=0, ipady=0)
        
     #   self.labelframeT = tkinter.LabelFrame(window, text="Download Template",  bd=3,relief=tkinter.RIDGE, font="Courier 12 bold")
     #   self.labelframeT.grid(column=3, row=20,columnspan=1, sticky='ew',padx=10, pady=10,ipadx=0, ipady=0)
        tem = tkinter.Label(self.labelframeT, text="cash_flow_file",  fg='blue',  cursor="hand2")
        tem.grid(column=4, row=19,rowspan=1, padx=40, pady=0)
        tem.bind("<Button-1>", self.callback)
        
        
        
        
        
        
       # self.flag=False
        
        
        self.progress = ttk.Progressbar(window, orient="horizontal",
                                        length=200, mode="determinate")
        self.progress.place(relx=0.5, rely=0.90, anchor="c")
        
        self.bytes = 0
        self.maxbytes = 0    
        
        self.close_btn = tkinter.Button(window, text = "Close", command = self.on_closing)# closing the 'window' when you click the button
        #self.close_btn.pack()
        self.close_btn.place(relx=0.5, rely=0.96, anchor="c")
    
    def start(self):
        self.progress["value"] = 0
        self.maxbytes = 5000
        self.progress["maximum"] = 5000
        self.fut_report()
        self.batch_report()
        self.load_fut()
        self.cashforBPM()
        self.lst_lod()
        self.load_trades()
 
# Link the cash flow File
    def callback(self, event):
        file_path=r'file://za.investment.int/dfs/dbshared/DFM/Tools/Indexation_trading_tools/IndexTrader/Templates/flows.csv'
        webbrowser.open_new(file_path)
        
        
# Create the futures report
    def fut_report(self):
        self.progress["value"] = 0
        #tkinter.messagebox.showinfo("Are flows & cash limits up to date: 1) Yes. 2) No.[Y/N]?:")
        response = tkinter.messagebox.askquestion("Flows", "Are flows & cash limits up to date?")
      #  print(response)
        lbl=tkinter.Label(window, text = "Futures report generated          \n  \n" ,font=("Helvetica", 10), fg='SystemButtonFace', bg='SystemButtonFace')
        lbl.place(relx=self.x_pos, rely=self.y_pos, anchor="c")
      #  lbl.grid(column=1, row=7, columnspan=1,padx=5, pady=5, sticky='e')
# If user clicks 'Yes' then it returns 1 else it returns 0
        if response == 'yes':
          #   def read_bytes(self):
              '''simulate reading 500 bytes; update progress bar'''
              self.bytes += 500
              self.progress["value"] = self.bytes
              if self.bytes < self.maxbytes:
            # read more bytes after 100 ms
                 self.after(100, self.fut_report)

              lbl=tkinter.Label(window, text = "Futures report generated    ", fg='green', font=("Helvetica", 10), bg='white')
              lbl.place(relx=self.x_pos, rely=self.y_pos, anchor="c")
        
        else:
            lbl=tkinter.Label(window, text = "Please Update Flows", fg='red', font=("Helvetica", 10), bg='white')
            lbl.place(relx=self.x_pos, rely=self.y_pos, anchor="c")
        if response=='yes':
            fut_calc(response)
            #print("yes")
          
# Create the load for futures
            
    def load_fut(self):
        self.progress["value"] = 0
        '''simulate reading 500 bytes; update progress bar'''
        self.bytes += 500
        self.progress["value"] = self.bytes
        lbl=tkinter.Label(window, text = "Futures report generated          \n  \n" , font=("Helvetica", 10),fg='SystemButtonFace', bg='SystemButtonFace')
        lbl.place(relx=self.x_pos, rely=self.y_pos, anchor="c")
        g=load_fut()
        
        if self.bytes < self.maxbytes:
    # read more bytes after 100 ms
        
             self.after(100, self.load_fut)
        lbl=tkinter.Label(window, text = g, fg='green', font=("Helvetica", 10), bg='white')
        lbl.place(relx=self.x_pos, rely=self.y_pos, anchor="c")
        #print("yes")
        
        
## Create the Batch Report
    def batch_report(self):
        self.progress["value"] = 0
        #tkinter.messagebox.showinfo("Are flows & cash limits up to date: 1) Yes. 2) No.[Y/N]?:")
        lbl=tkinter.Label(window, text = " \n \n gggggggggggggggggggggggggg" ,fg='SystemButtonFace', font=("Helvetica", 10), bg='SystemButtonFace')
        lbl.place(relx=self.pos_x, rely=self.pos_y, anchor="c")
      #  lbl.grid(column=2, row=9, sticky='ew',padx=5, pady=5)

        response = tkinter.messagebox.askquestion("Flows", "Are flows & cash limits up to date?")
       # print(response)
       
# If user clicks 'Yes' then it returns 1 else it returns 0
        if response == 'yes':
          #   def read_bytes(self):
              '''simulate reading 500 bytes; update progress bar'''
              self.bytes += 500
              self.progress["value"] = self.bytes
              if self.bytes < self.maxbytes:
            # read more bytes after 100 ms
                 self.after(100, self.batch_report)

              lbl=tkinter.Label(window, text = "Batch cash calc \n generated!" ,fg='green', font=("Helvetica", 10), bg='white')
              lbl.place(relx=self.pos_x, rely=self.pos_y, anchor="c")
              #lbl.grid(column=2, row=9, sticky='ew',padx=5, pady=5)
        else:
            lbl=tkinter.Label(window, text = "Please Update Flows",fg='red', font=("Helvetica", 10), bg='white')
            lbl.place(relx=self.pos_x, rely=self.pos_y, anchor="c")
            #lbl.grid(column=2, row=9, sticky='ew',padx=5, pady=5)
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
        lbl=tkinter.Label(window, text = " \n \n gggggggggggggggggggggggggg" ,fg='SystemButtonFace', bg= 'SystemButtonFace',font=("Helvetica", 10))
        lbl.place(relx=self.pos_x, rely=self.pos_y, anchor="c")
        #lbl.grid(column=2, row=10, rowspan = 1, sticky='ew',padx=5, pady=5)

        b=cash_bpm()
    #    tkinter.Label(window, text = b).pack()
        self.bytes += 500
        self.progress["value"] = self.bytes
        if self.bytes < self.maxbytes:
    # read more bytes after 100 ms
             self.after(100, self.cash_bpm)
        lbl=tkinter.Label(window, text = b,fg='green', font=("Helvetica", 10), bg='white')
        lbl.place(relx=self.pos_x, rely=self.pos_y, anchor="c")     
        #lbl.grid(column=2, row=10, sticky='ew', rowspan = 1, padx=5, pady=5)
        
# Create files for listener 
        
    def lst_lod(self):
        self.progress["value"] = 0
        '''simulate reading 500 bytes; update progress bar'''
        lbl=tkinter.Label(window, text = " \n \n gggggggggggggggggggggggggg" ,fg='SystemButtonFace', bg= 'SystemButtonFace',font=("Helvetica", 10))
        lbl.place(relx=self.pos_x, rely=self.pos_y, anchor="c")
        #lbl.grid(column=2, row=10, rowspan = 1, sticky='ew',padx=5, pady=5)

        b=list_lds()
    #    tkinter.Label(window, text = b).pack()
        self.bytes += 500
        self.progress["value"] = self.bytes
        if self.bytes < self.maxbytes:
    # read more bytes after 100 ms
             self.after(100, self.list_lds)
        lbl=tkinter.Label(window, text = b,fg='green', font=("Helvetica", 10), bg='white')
        lbl.place(relx=self.pos_x, rely=self.pos_y, anchor="c")     
        #lbl.grid(column=2, row=10, sticky='ew', rowspan = 1, padx=5, pady=5)

# Load trades
    def load_trades(self):
       # self.flag= True
        self.progress["value"] = 0
        '''simulate reading 500 bytes; update progress bar'''
        lbl=tkinter.Label(window, text = " \n \n gggggggggggggggggggggggggg" ,fg='SystemButtonFace', bg='SystemButtonFace', font=("Helvetica", 10))
        lbl.place(relx=self.pos_x, rely=self.pos_y, anchor="c")
        def select_fx():
            global d
            lbl=tkinter.Label(window, text = " \n \n \n" ,fg='grey', font=("Helvetica", 10), width =20)
            lbl.place(relx=self.pos_x, rely=self.pos_y, anchor="c")
       
            selection = var.get()
           # print(var)
            print(selection)
#            if selection==1:
#                print("y")
#            else:
#                print("n")
            d=tloader_equity_or_fut(selection)
            print(d)
            lbl=tkinter.Label(window, text = d,fg='green', font=("Helvetica", 10), bg='white')
            lbl.place(relx=self.pos_x, rely=self.pos_y, anchor="c")
             
          #  return d 
       
    #    tkinter.Label(window, text = b).pack()
        self.bytes += 500
        self.progress["value"] = self.bytes
        if self.bytes < self.maxbytes:
    # read more bytes after 100 ms
             self.after(100, self.tloader_equity_or_fut, True)
        var = tkinter.IntVar()
       # if self.flag:
        st=13
        labelframeC = tkinter.LabelFrame(window, text="Select trades to upload:",  bd=3,relief=tkinter.RIDGE, font=("Courier 10 bold"))
        labelframeC.grid(column=0, row=st,columnspan=2, sticky='ew',padx=5, pady=5)
        #tkinter.Label(labelframeC, text = "Select trades to upload:",fg='black', font=("Helvetica", 8)).grid(column=1, row=st+1)
        tkinter.Radiobutton(labelframeC, text = "Equities only", variable = var, value = 1,font=("Helvetica", 8)).grid(column=1, row=st+2)
        tkinter.Radiobutton(labelframeC, text = "Futures only", variable = var, value = 2,font=("Helvetica", 8)).grid(column=1, row=st+3)
        tkinter.Radiobutton(labelframeC, text = "Both Equities & Futures", variable = var, value = 3,font=("Helvetica", 8)).grid(column=1, row=st+4,padx=20)
        tkinter.Button(labelframeC, text = "OK", command = select_fx).grid(column=1, row=st+5)
       # self.flag=False
      #  else:
           # self.after(100, self.tloader_equity_or_fut, False)
        print(self.flag)
        #return self.flag
        #print(kgl)
        
     
  #  print("The flag" +str(self.flag))
             
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


#
#import tkinter
#from tkinter import *
#
#def function():
#    selection = var.get()
#
#    if  selection == 1:
#        # Default
#        print(selection)
#
#    elif selection == 2:
#        # User-defined
#        print("No")
#
#    else:#selection==0
#        #No choice
#        print("What")
#
#    master.quit()
#
#master = Tk()
#var = IntVar()
#Label(master, text = "Select OCR language").grid(row=0, sticky=W)
#Radiobutton(master, text = "default", variable = var, value = 1).grid(row=1, sticky=W)
#Radiobutton(master, text = "user-defined", variable = var, value = 2).grid(row=2, sticky=W)
#Button(master, text = "OK", command = function).grid(row=3, sticky=W)
#mainloop()
#
#




