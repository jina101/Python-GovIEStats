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
  
#function to set a working directory
def set_directory():
    #select a working directory:
    global folder_selected
    folder_selected = filedialog.askdirectory()
    os.chdir(folder_selected)
    
    if folder_selected is None:
        folder_selected = os.getcwd()
        
    
#function to open files and put them into a dataframe
#the files that will open are csv JSON and JSON-Stat
def open_file():
    #csv file
    file = askopenfile(mode ='rb+', filetypes =[('CSV Files', '*.csv'),('JSON Files', '*.json')])
    df = pd.read_csv(file)
    
    if file is not None:
        print(df.describe())
    
        #save the file as an excel
        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter('myexcel.xlsx', engine='xlsxwriter')

        # Convert the dataframe to an XlsxWriter Excel object.
        df.to_excel(writer, sheet_name='Sheet1', index=None)
        writer.save()

#set working directory button
set_dir_btn = Button(root, text="Set working directory", command=lambda:set_directory())
set_dir_btn.pack(side = TOP, pady = 10)

open_btn = Button(root, text ='Open File', command = lambda:open_file())
open_btn.pack(side = TOP, pady = 10)

exit_btn = Button(root, text = "Exit", command = root.destroy)
exit_btn.pack(side = TOP, pady = 10)  
  
root.mainloop()

    
