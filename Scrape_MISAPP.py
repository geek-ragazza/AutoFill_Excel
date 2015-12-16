import pandas as pd
import sqlite3 as lite
import numpy as np
from pandas import ExcelWriter
from pandas import ExcelFile
import datetime
import gc
import os
from xlwings import Workbook, Sheet, Range, Chart
import sys,os
from pywinauto import application
import pywinauto
from win32com.client import Dispatch
import subprocess
app=application.Application()
import time
import win32com.client
shell = win32com.client.Dispatch("WScript.Shell")

#Login to MIS_APP

MyPath ='C:\misapp'
os.chdir(MyPath) 
app = app.start_('misapp.exe')

keys ='rapunzel'
pywinauto.SendKeysCtypes.SendKeys(keys, pause=0.050000000000000003, with_spaces=False, with_tabs=False, with_newlines=True, turn_off_numlock=True)
app.misapp.TreeView20WndClass.Click()

# Enter into mls.exe
app.misapp.CtlFrameWork_ReflectWindow.Click()


print app.misapp.CtlFrameWork_ReflectWindow.PrintControlIdentifiers() 
print app.misapp.CtlFrameWork_ReflectWindow.FriendlyClassName()
print app.misapp.CtlFrameWork_ReflectWindow.Class()
app.windows_()
app.misapp.CtlFrameWork_ReflectWindow.ClickInput(coords=(17, 35))
time.sleep(1)
app.misapp.CtlFrameWork_ReflectWindow.DoubleClickInput(coords=(40, 50))

#Rename the excel file

def RenameFile():
    PONumber = '14P1701A'
    tempName='14P1701A PO_Item.xls'
    os.chdir('C:\\misapp\\report')
    xl = pd.ExcelFile(tempName)
    df = xl.parse()
    df.dropna(how='any')
    xlsPO_Number =  df.ix[3,0]
    print xlsPO_Number
    NewName=('%s_%s.xls' %(PONumber,xlsPO_Number))
    os.rename(tempName,NewName)
    
    
#Select Receive Process -> Receive Process 

NowWin = app.windows_()


j=0
for i in NowWin:
    j+=1
    
    try: 
        
        
        if i.IsVisible() == True:
            print i.FriendlyClassName()
            #i.DragMouse(button='left', press_coords=(280,35), release_coords=(280,45))
            
        if i.RunTests()!=None and i.FriendlyClassName()=='CtlFrameWork_Parking':
            
            i.SetFocus()
            i.ClickInput(button='left', coords=(200,30))
            time.sleep(3)
            
            shell.SendKeys('{DOWN}', 0)
            shell.SendKeys('{ENTER}', 0)
            
    except:
        pass

shell.SendKeys('{ENTER}', 0)
time.sleep(1)
#Enter Ponumber
shell.SendKeys('14P1701A{ENTER}', 0)
time.sleep(3)


#Start the loop
NowWinEnter = app.windows_()
for EnterTime in range(1,20):
    time.sleep(1)
    for i in NowWinEnter:

        try:
            if i.RunTests()!=None and i.FriendlyClassName()=='CtlFrameWork_Parking':
                #Select PO Number
                i.ClickInput(button='left', coords=(490,153))
                shell.SendKeys('{DOWN}', 0)
                
                for et in range(1,EnterTime+1):
                    shell.SendKeys('{ENTER}', 0)
                    time.sleep(0.5)
                    
                time.sleep(3)
                #Click Export Excel
                time.sleep(0.5)
                i.ClickInput(button='left', coords=(680,580))
                print "clicked"
        except:
            pass

    time.sleep(5)    

    NowWinConfirm = app.windows_()
    for i in NowWinConfirm:
        try:
            if i.RunTests()!=None and i.FriendlyClassName()=='CtlFrameWork_Parking':
                shell.SendKeys('{ENTER}', 0)

        except:
            pass
    
    time.sleep(5)
    try:
        RenameFile()
    except:
        pass
    time.sleep(2)
