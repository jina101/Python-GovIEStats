# -*- coding: utf-8 -*-
"""
Created on Wed Mar 31 12:11:06 2021
@author: ranja
"""



#importing all relevant modules from Tkinter, pandas, os and numpy, xlsxwriter, json etc
import os
import io
from tkinter import * 
from tkinter.ttk import *
import pandas as pd
import numpy as np
import xlsxwriter
import json as jsn
from pyjstat import pyjstat
from sys import exit
import tkinter.filedialog


!git add "govie_descriptive_stats.py"
!git commit -m "My commit"
!git push origin master
  
#get ask oppenfile
from tkinter.filedialog import askopenfile
 
#set up a Tk winddow  
root = Tk()
root.geometry('500x150')
root.title("Open a CSV, JSON, JSON-Stat File or Set Working Directory")

#working directory is the current wd, if not chosen
folder_selected = os.getcwd()
 
#function to set a working directory
def set_directory():
    #select a working directory:
    global folder_selected
    try:
        folder = filedialog.askdirectory()
        os.chdir(folder)
        print("The current working directory is " + folder)
        folder_selected = folder
    except Exception:
        os.chdir(folder_selected)
        print("The current working directory is " + folder_selected)

#function to save the file
def save_file(name, dataframe):
     #save the file as an excel
        # Create a Pandas Excel writer using XlsxWriter as the engine.
    try:
        writer = pd.ExcelWriter(name + '.xlsx', engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
        dataframe.to_excel(writer, sheet_name="Sheet 1", index=None)
        writer.save()
        print("An excel file has been saved in the following directory: " + folder_selected)
    
    except Exception as error:
         print(f'Unable to name file. An error occured: <{error}> Please exit and try again.')
         exit()
  
#function + gui for saving a file
def save_file_dialogue():
      global e
      global df
     
      
    # Ask user for the name of the file they would like to save:
      root2 = Tk()
      root2.geometry('600x300')
      root2.title("Save File")
      #root2.title('Your file will now be exported to Excel to the following directory '+ folder_selected + ', please insert a name for it below:')
     
      e = Entry(root2)
      e.pack()
      e.focus_set()

      b = Button(root2,text='Save File',command=lambda:save_file(e.get(), df))
      b.pack(side = TOP, pady = 10)
      eX = Button(root2, text = "Exit", command = root2.destroy)
      eX.pack(side = TOP, pady = 10) 
          
  
#function to open files and put them into a dataframe
#the files that will open are csv JSON and JSON-Stat
def open_file():
    global e
    global df
    
    #csv file
    try:
        file = tkinter.filedialog.askopenfilename(initialdir = "/", filetypes =(('CSV Files', '*.csv'),('Text Files', '*.txt'),('JSON Files', '*.json')))

        if file is not None and file.lower().endswith(('.json', '.txt')):
            
                #json-stat files
            try:
                f = open(file)
                filepath = f.read()
                data = pyjstat.Dataset.read(filepath)
                df = data.write('dataframe')
                print("Please see descriptive details below:")
                print(df.info())
                print(df.head(5))
                save_file_dialogue()
            except:
                pass
            
            #json files   
            try:
                f = open(file)
                filepath = f.read()
                data = jsn.loads(filepath)
                df = pd.json_normalize(data['features'])
                print("Please see descriptive details below:")
                print(df.info())
                print(df.head(5))
                save_file_dialogue()
            except Exception as err:
                print(f'Cannot open file. An error has occurred: <{err}>')
           
     
        #read a csv
        else:
            try:

                df = pd.read_csv(file, encoding='utf-8')
                print("Please see descriptive details below:")
                print(df.info())
                print(df.head(5))
                save_file_dialogue()    
            except Exception as e:
                print(f'Cannot open file. An error has occurred: <{e}>')
                exit()


    
    except(OSError,FileNotFoundError):
        print(f'Unable to find or open <{file}> Please exit and try again.')
        exit()
    except Exception as error:
        print(f'An error has occurred: <{error}> Please exit and try again.')
        exit()
    
    
        
       

#set working directory button
set_dir_btn = Button(root, text="Set working directory", command=lambda:set_directory())
set_dir_btn.pack(side = TOP, pady = 10)

open_btn = Button(root, text ='Open File', command = lambda:open_file())
open_btn.pack(side = TOP, pady = 10)

exit_btn = Button(root, text = "Exit", command = root.destroy)
exit_btn.pack(side = TOP, pady = 10)  
  
root.mainloop()