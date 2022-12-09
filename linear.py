import tkinter
from tkinter import *
from tkinter.ttk import *
import pandas as pd
import numpy as np
import warnings
import matplotlib.pyplot as plt
from wordcloud import WordCloud
from matplotlib import rcParams
from PIL import Image
from PIL import ImageTk,Image
import sqlite3
from openpyxl import Workbook
from openpyxl import load_workbook


window = Tk()
window.title("Check attendance")
window.geometry("600x400")


#connect to excel file
wb = Workbook()
wb = load_workbook("hello.xlsx")
#print from excel file function
def query():
    sheet_name = str(combo.get())
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        list_all = ' '
        for value in sheet_to_df_map[combo.get()].values:
            for i in range (0,len(value)):
                list_all = list_all + "{:15}".format(str(value[i]))
                #list_all = f'{list_all + str((value[i]))}'
            list_all = f'{list_all}\n'

        a.configure(text=list_all)
#====================================================================

#label
a = tkinter.Label(window,text="",fg="blue",font=("Ariel",10),justify=LEFT)
a.place(x= 250,y=130,height=100,width=200)


# set name

#===================================================================

#set color
canvas = Canvas(window,bg="red",width=200,height=600)
canvas.place(x=0,y=0)
canvas_1 = Canvas(window,bg="orange",width=400,height=100)
canvas_1.place(x=202,y=0)
name = tkinter.Label(window,bg="orange",text="Check attendance",font=("Ariel",30))
name.place(x=250,y=20)
#canvas.place(x=0,y=0)
#create label date
date = tkinter.Label(window,text="Date: ",bg="red")
date.place(x=0,y=127)
#combobox
combo = Combobox(window)
#read file name from exel
data = pd.ExcelFile("hello.xlsx")
data.sheet_names
sheet_to_df_map = {}
for sheet_name in data.sheet_names:
    sheet_to_df_map[sheet_name] = data.parse(sheet_name)

combo["value"] = list(sheet_to_df_map.keys())
combo.place(x=30,y=125,width=100,height=30)
#================================================================


#create a show button
show_combo = tkinter.Button(window,text="Show record",command=query)
show_combo.place(x=30,y=180,width=100,height=30)

#exit button
exit_button = tkinter .Button (text = " Exit ", command= window.destroy)
exit_button.place (x = 5 ,y = 300 ,width =100 ,height = 50)

window.mainloop()