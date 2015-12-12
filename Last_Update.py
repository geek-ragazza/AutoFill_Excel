#coding=utf-8
import os.path, time
import os, sys
import datetime
import shutil
import betterwalk
from Tkinter import *
Last_Time=0
CLast_Time=0
LocalPath ='//Ecs01//ddb00/final_price'
for pathinfile, subdirs, files in betterwalk.walk(LocalPath):
    for TempFilePath in map(lambda x:os.path.join(pathinfile,x),files):
        Create_Time = os.path.getctime(TempFilePath)
        Modified_Time = os.path.getmtime(TempFilePath)
        if Last_Time<Modified_Time and os.path.splitext(TempFilePath)[1]!='.db':  #except thumb.db 
            Last_Time=Modified_Time
            FileName = TempFilePath
        
        if CLast_Time<Create_Time and os.path.splitext(TempFilePath)[1]!='.db': #except thumb.db 
            CLast_Time=Create_Time
            CFileName = TempFilePath
            
Result = time.ctime(Last_Time)
CResult = time.ctime(CLast_Time)


if __name__ == "__main__":
        
        root = Tk()
        root.geometry('900x250+200+300')
        root.configure(background ='#E2E2E2')
        F1 = Frame(root)
        F1.grid(row=0, column=0)
        F1.configure(background ='#E2E2E2')
      




        MTimeL = Label(F1, text="採購部檔案最後變更時間")
        MTimeL.grid(row = 0, column = 0)


        MTime = Label(F1, text=Result)
        MTime.grid(row = 1, column = 0)
        
        MName = Label(F1, text=FileName)
        MName.grid(row = 2, column = 0)
        
        Empty= Label(F1, text="    ")
        Empty.grid(row=3, column = 0)
        
        CMTimeL = Label(F1, text="採購部檔案最後建立時間")
        CMTimeL.grid(row = 4, column = 0)


        CMTime = Label(F1, text=CResult)
        CMTime.grid(row = 5, column = 0)
        
        CMName = Label(F1, text=CFileName)
        CMName.grid(row = 6, column = 0)






root.mainloop()
