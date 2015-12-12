
# coding: utf-8



import pandas
import os.path, time
import os, sys
import datetime
import shutil
from Tkinter import *
from tkFileDialog import askopenfilename,askdirectory     
import Search_History




def EntryReadOnly(row, col, text, frameName,widthlenth):

        ReadE = StringVar()
        EntryFun = Entry(frameName, width=widthlenth, font = "Helvetica 12",  textvariable=ReadE, state='readonly', borderwidth=1,)   
        EntryFun.grid(row=row, column=col, sticky=W, padx=5)
        #EOil.configure(background ='#FFFFFF')
        ReadE.set(text)
        return ReadE




def SearchDirectory(butnum):
        
        if butnum==1:
                global search_directory_path
                search_directory_path = askdirectory()
                search_directory_path = search_directory_path.replace('/','\\')
                
                RemoteL=EntryReadOnly(2, 0, search_directory_path, F1,20)


# In[ ]:

def get():
    searhc_type=search_option_var.get()
    search_word=search_wordE.get()
    x=Search_History.SearchFile(search_directory_path,search_word,searhc_type)
    Result=EntryReadOnly(1, 1, x, F1,20)

# In[ ]:

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




        #Set Search directory
        UnoExcelBut=Button(F1, text='選取搜尋路徑', command=lambda: SearchDirectory(1))
        UnoExcelBut.grid(row = 1, column = 0)
        search_option_var = StringVar()
        search_option_var.set("None")
        search_option = OptionMenu(F1, search_option_var, "None", "OR", "AND")
        search_option.grid(row = 3, column = 0)




        search_wordL = Label(F1, text="輸入搜尋條件")
        search_wordL.grid(row = 4, column = 0)


        search_wordE = Entry(F1, width = 10)
        search_wordE.grid(row = 5, column = 0)




        But1 = Button(F1, command = get, text="搜尋")
        But1.grid(row = 6, column = 0)

root.mainloop()






