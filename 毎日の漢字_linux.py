from tkinter import *
import pandas as pd
import numpy as np
from openpyxl import load_workbook
import os


def kanjis():
    n = int(num_kanjis.get())
    df = pd.read_excel("data/漢字.xlsx")
    readings = ['次','くにょみ','おにょみ']
    aux = ['次','いみ','漢字',]
    new_kanji = df.loc[df['新しい']==1]
    df = df.loc[df['新しい']==0]
    df['random'] =  np.random.random(size=len(df))
    df.sort_values(by="random", inplace = True, ascending=False)


    if(new_kj.get()==1):

        if(excel_display.get()==0): 
            daily_kanji_readings = pd.concat([df[readings].head(n).reset_index(drop = True),new_kanji[readings]])
            daily_kanji = pd.concat([df[aux].head(n).reset_index(drop = True),new_kanji[aux]])
            daily_kanji_readings.to_excel('毎日の漢字.xlsx',"lecturas",index = False)

        if(excel_display.get()==1): 
            daily_kanji_readings = pd.concat([df[readings].head(n).reset_index(drop = True),new_kanji[readings]])
            daily_kanji = pd.concat([df[aux].head(n).reset_index(drop = True),new_kanji[aux]])
            daily_kanji_readings.to_excel('毎日の漢字.xlsx',"lecturas",index = False)
            os.system("start EXCEL.EXE 毎日の漢字.xlsx")
    

    if(new_kj.get()==0):
        
        if(excel_display.get()==0): 
            daily_kanji_readings = df[readings].head(n).reset_index(drop = True)
            daily_kanji = df[aux].head(n).reset_index(drop = True)
            daily_kanji_readings.to_excel('毎日の漢字.xlsx',"lecturas",index = False)

        if(excel_display.get()==1): 
            daily_kanji_readings = df[readings].head(n).reset_index(drop = True)
            daily_kanji = df[aux].head(n).reset_index(drop = True)
            daily_kanji_readings.to_excel('毎日の漢字.xlsx',"lecturas",index = False)
            os.system("start EXCEL.EXE 毎日の漢字.xlsx")

              

ventana = Tk()
ventana.title('毎日の漢字プロガム')
icon = PhotoImage(file="日本語.png")
ventana.resizable(False,False)
ventana.iconphoto(False, icon, icon)    


frame=Frame()
frame.pack()
frame.config(bg='white smoke')
frame.config(width="740", height ="320")


image_1 = PhotoImage(file="f1.png")
Label(frame, image= image_1).place(x=-150,y=-50)
Label(frame,text = '毎日の漢字 ',background='white smoke',fg='black',font=('Arial',20)).place(x=50,y=40)
Label(frame,text = '毎日の漢字数',background='white smoke',fg='black',font=('Arial',12)).place(x=50,y=110)


num_kanjis = Entry(frame,width=4, font=('Arial 10'))
num_kanjis.place(x=180,y=110)


new_kj = IntVar()
excel_display = IntVar()
Checkbutton(frame, text='新漢字',background='white',highlightthickness=0,font=('Arial',11),variable = new_kj, onvalue = 1,offvalue = 0,command = kanjis).place(x=50,y=160)
Checkbutton(frame, text='エクセルを開く',background='white',highlightthickness=0,font=('Arial',11),variable = excel_display, onvalue = 1,offvalue = 0,command = kanjis).place(x=50,y=190)


Button(frame,text='引き金',command=kanjis).place(x=100,y=250)



ventana.mainloop()