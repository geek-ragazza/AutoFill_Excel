# -*- coding: utf-8 -*-


import pandas
import os.path, time
import os, sys
import datetime
import shutil
from Tkinter import *
from tkFileDialog import askopenfilename,askdirectory     
import PanelCheckModule




def EntryReadOnly(row, col, text, frameName,widthlenth):

        ReadE = StringVar()
        EntryFun = Entry(frameName, width=widthlenth, font = "Helvetica 12",  textvariable=ReadE, state='readonly', borderwidth=1,)   
        EntryFun.grid(row=row, column=col, sticky=W, padx=5)
        #EOil.configure(background ='#FFFFFF')
        ReadE.set(text)
        return ReadE




def OpenExcel(butnum):
        
        if butnum==1:
                global directory_pathL
                directory_pathL = askopenfilename()
                directory_pathL = directory_pathL.replace('/','\\')
                print directory_pathL
                RemoteL=EntryReadOnly(2, 0, directory_pathL, F1,20)
        if butnum==2:
                global directory_pathR
                directory_pathR = askopenfilename()
                directory_pathR = directory_pathR.replace('/','\\')
                print directory_pathR
                RemoteE=EntryReadOnly(2, 2, directory_pathR, F1,20)
        #start check
        if butnum==3:
                
                LocalPath = directory_pathL.replace('/','\\')
                RemotePath = directory_pathR.replace('/','\\')
                
                ButtonTreAction(LocalPath,RemotePath)




def get():
    TitleCheck1=TitleCheck1E.get()
    TitleCheck2 = TitleCheck2E.get()
    UnoTitleCol = UnoTitleColE.get()
    DueTitleCol = DueTitleColE.get()
    UnoPriceCol = UnoPriceColE.get()
    DuePriceCol = DuePriceColE.get()
    UnoStart = UnoStartE.get()
    UnoEnd = UnoEndE.get()
    DueStart = DueStartE.get()
    DueEnd = DueEndE.get()
    
    directory_pathUno=directory_pathL
    directory_pathDue=directory_pathR
    
    Result = PanelCheckModule.MergenFillPrice(directory_pathUno, directory_pathDue, TitleCheck1, TitleCheck2,
                                              UnoTitleCol, DueTitleCol, UnoPriceCol, DuePriceCol,
                                              UnoStart, UnoEnd, DueStart, DueEnd)
    EntryReadOnly(12,0,Result,F1,10)
    




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

        
        
        
        #Read File Name
        UnoExcelBut=Button(F1, text='Original Excel', command=lambda: OpenExcel(1))
        UnoExcelBut.grid(row = 1, column = 0)
        DueExcelBut=Button(F1, text='Excel Include Price', command=lambda: OpenExcel(2))
        DueExcelBut.grid(row = 1, column = 2)
        




        TitleCheck1L = Label(F1, text="Excel1 輸入任一個配電盤名稱的位置(A1)")
        TitleCheck1L.grid(row = 3, column = 0)
        TitleCheck1E = Entry(F1, width = 10)
        TitleCheck1E.grid(row = 4, column = 0)

        TitleCheck2L = Label(F1, text="Excel2 輸入任一個配電盤名稱的位置(A1)")
        TitleCheck2L.grid(row = 3, column = 2)
        TitleCheck2E = Entry(F1, width = 10)
        TitleCheck2E.grid(row = 4, column = 2)        


        

        UnoTitleColL = Label(F1, text="Excel1 比對範圍欄欄位(字母)")
        UnoTitleColL.grid(row = 5, column = 0)
        UnoTitleColE = Entry(F1, width = 10)
        UnoTitleColE.grid(row = 6, column = 0)

        DueTitleColL = Label(F1, text="Excel2 比對範圍欄欄位(字母)")
        DueTitleColL.grid(row = 5, column = 2)
        DueTitleColE = Entry(F1, width = 10)
        DueTitleColE.grid(row = 6, column = 2)        


        

        UnoPriceColL = Label(F1, text="Excel1 價格欄位(字母)")
        UnoPriceColL.grid(row = 7, column = 0)
        UnoPriceColE = Entry(F1, width = 10)
        UnoPriceColE.grid(row = 8, column = 0)

        DuePriceColL = Label(F1, text="Excel2 價格欄位(字母)")
        DuePriceColL.grid(row = 7, column = 2)
        DuePriceColE = Entry(F1, width = 10)
        DuePriceColE.grid(row = 8, column = 2)




        UnoStartL = Label(F1, text="Excel1 開始範圍(數字)")
        UnoStartL.grid(row = 9, column = 0)
        UnoStartE = Entry(F1, width = 10)
        UnoStartE.grid(row = 10, column = 0)

        UnoEndL = Label(F1, text="Excel1 結束範圍(數字)")
        UnoEndL.grid(row = 9, column = 1)
        UnoEndE = Entry(F1, width = 10)
        UnoEndE.grid(row = 10, column = 1)


        DueStartL = Label(F1, text="Excel2 開始範圍(數字)")
        DueStartL.grid(row = 9, column = 2)
        DueStartE = Entry(F1, width = 10)
        DueStartE.grid(row = 10, column = 2)

        DueEndL = Label(F1, text="Excel2 結束範圍(數字)")
        DueEndL.grid(row = 9, column = 3)
        DueEndE = Entry(F1, width = 10)
        DueEndE.grid(row = 10, column = 3)



        But1 = Button(F1, command = get, text="Start")
        But1.grid(row = 11, column = 0)
        
root.mainloop()






