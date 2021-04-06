# -*- coding: utf-8 -*-
"""
Created on Wed Mar 31 12:11:06 2021

@author: ranja
"""



#importing all relevant modules from Tkinter, pandas, os and numpy
import os
import io
from tkinter import * 
from tkinter.ttk import *
import pandas as pd
import numpy as np
import xlsxwriter
  
#get ask oppenfile
from tkinter.filedialog import askopenfile
 
#set up a Tk winddow  
root = Tk()
root.geometry('200x200')

#working directory is the current wd, if not chosen
folder_selected = os.getcwd()
 
#function to set a working directory
def set_directory():
    #select a working directory:
    global folder_selected
    folder_selected = filedialog.askdirectory()
    os.chdir(folder_selected)

#function to save the file
def save_file(name, dataframe):
     #save the file as an excel
        # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(name + '.xlsx', engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    dataframe.to_excel(writer, sheet_name="Sheet 1", index=None)
    writer.save()
    print("An excel file has been saved in the following directory: " + folder_selected)
  
  
#function to open files and put them into a dataframe
#the files that will open are csv JSON and JSON-Stat
def open_file():
    global e
    global df
    
    #csv file
    file = askopenfile(mode ='rb+', filetypes =[('CSV Files', '*.csv'),('JSON Files', '*.json')])
    df = pd.read_csv(file)
    
    if file is not None:
        print(df.info())
    
    
        # Ask user for the name of the file they would like to save:
        root2 = Tk()
        root2.geometry('600x300')
        root2.title("Save File")
       # root2.title('Your file will now be exported to Excel to the following directory '+ folder_selected + ', please insert a name for it below:')

        e = Entry(root2)
        e.pack()
        e.focus_set()

        b = Button(root2,text='Save File',command=lambda:save_file(e.get(), df))
        b.pack(side='bottom')
        root.mainloop()
       

#set working directory button
set_dir_btn = Button(root, text="Set working directory", command=lambda:set_directory())
set_dir_btn.pack(side = TOP, pady = 10)

open_btn = Button(root, text ='Open File', command = lambda:open_file())
open_btn.pack(side = TOP, pady = 10)

exit_btn = Button(root, text = "Exit", command = root.destroy)
exit_btn.pack(side = TOP, pady = 10)  
  
root.mainloop()

    
