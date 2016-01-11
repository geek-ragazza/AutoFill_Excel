# -*- coding: utf-8 -*-

import pandas
import os.path, time
import os, sys
import datetime
from datetime import date, timedelta
import shutil
from Tkinter import *
from tkFileDialog import askopenfilename,askdirectory     
import PCTOU

FA_Path = '\\\\ecs01\\ddd00\\Dept\\成本\\5-益鼎專案成本\\會計月結資料\\Acc303\\fa.xlsx'
FA_Path=unicode(FA_Path,'utf8')
TempPath = r'C:\Python2.7.10\Scripts\notebook\seach'
removeFAPath=  os.path.join(TempPath, 'fa.xlsx')
try:
    os.remove(removeFAPath)
except:
    pass
try:
    shutil.copy2(FA_Path,TempPath)
except:
    pass
ACC_Path='\\\\Ecs01\\ddd00\\Dept\\成本\\5-益鼎專案成本\\結帳用資料\\工時\\2015年.xls'
ACC_Path=unicode(ACC_Path,'utf8')

d = datetime.date.today()- timedelta(days=20)
SheetMonth = '{:02d}'.format(d.month)
SheetYear = format(d.year)

ACC_OriginalName = str(format(d.year)) + '年.xls'
ACC_FullPath = os.path.join(TempPath, ACC_OriginalName)
ACC_FullPath=unicode(ACC_FullPath,'utf8')
ACC_Rename = str(format(d.year)) + '.xls'

ACC_RenamePath= os.path.join(TempPath,ACC_Rename)
try:
    os.remove(ACC_RenamePath)
except:
    pass
try:
    shutil.copy2(ACC_Path,TempPath)
    os.rename(ACC_FullPath,ACC_RenamePath)
except:
    pass




def EntryReadOnly(row, col, text, frameName,widthlenth):

        ReadE = StringVar()
        EntryFun = Entry(frameName, width=widthlenth, font = "Helvetica 12",  textvariable=ReadE, state='readonly', borderwidth=1,)   
        EntryFun.grid(row=row, column=col, sticky=W, padx=5)
        #EOil.configure(background ='#FFFFFF')
        ReadE.set(text)
        return ReadE

def OpenExcel(butnum):
        
        if butnum==1:
                global BalanceSheet_path
                BalanceSheet_path = askopenfilename()
                BalanceSheet_path = BalanceSheet_path.replace('/','\\')
                #print BalanceSheet_path
                BSPath=EntryReadOnly(1, 0, BalanceSheet_path, F1,20)
        if butnum==2:
                global T0U_path
                T0U_path = askopenfilename()
                T0U_path = T0U_path.replace('/','\\')
                print T0U_path
                T0UPath=EntryReadOnly(4, 0, T0U_path, F1,20)


def get():
    ProYear=ProYearE.get()
    ProCode=ProCodeE.get()
    
    Alphabet = ['A', 'C', 'P', 'E']
    Pro_Name = map(lambda x:ProYear + x + ProCode, Alphabet)
    
    PCTOU.execute(BalanceSheet_path, T0U_path, ACC_RenamePath, removeFAPath, Pro_Name)


if __name__ == "__main__":
        
        root = Tk()
        root.geometry('660x350+500+300')
        root.configure(background ='#E2E2E2')
        F1 = Frame(root)
        F1.grid(row=0, column=0)
        F1.configure(background ='#E2E2E2')
        F2 = Frame(root)
        F2.grid(row=0, column=1)
        F2.configure(background ='#E2E2E2')

        TimeName = Label(F1, text="工時成本資料為")
        TimeName.grid(row = 0, column = 3)
        YMCombine = str(format(d.year)) + "年" + str(format(d.month))+ "月"
        Timenow = Label(F1, text=YMCombine)
        Timenow.grid(row = 1, column = 3)
        #Set Search directory
        BalanceSheetBut=Button(F1, text='選取Balance Sheet', command=lambda: OpenExcel(1))
        BalanceSheetBut.grid(row = 0, column = 0)
        T0UBut=Button(F1, text='選取301表', command=lambda: OpenExcel(2))
        T0UBut.grid(row = 3, column = 0)
        

        ProYearL = Label(F1, text="輸入專案年份")
        ProYearL.grid(row = 0, column = 1)
        

        ProYearE = Entry(F1, width = 10)
        ProYearE.grid(row = 0, column = 2)

        ProCodeL = Label(F1, text="輸入專案編號")
        ProCodeL.grid(row = 1, column = 1)
        

        ProCodeE = Entry(F1, width = 10)
        ProCodeE.grid(row = 1, column = 2)

        But1 = Button(F1, command = get, text="填入報表!")
        But1.grid(row = 6, column = 2)
        
root.mainloop()

