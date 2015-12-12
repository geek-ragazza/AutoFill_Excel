# coding: utf-8

import sqlite3 as lite
import os.path, time
import os, sys
import datetime
import shutil
import pandas as pd
import numpy as np
from pandas import ExcelWriter
from pandas import ExcelFile
from xlwings import Workbook, Sheet, Range, Chart
import betterwalk
import re

def isListEmpty(inList):
    if isinstance(inList, list): # Is a list
        return all( map(isListEmpty, inList) )
    return False # Not a list'

def ANDCheckEmpty(inList):
    for checker in inList:
        if not checker:
            return False
    return True
        
def SearchFile(Path,Search_Condition,Search_Type):
    #Create New DataFrame

    columns1=['File Name', 'Path', 'Folder']
    index1 = np.arange(30000)
    df = pd.DataFrame(columns=columns1, index = index1)



    #Search => None
    if Search_Type == 'None':
        
        i=(-1)
        #Path=unicode(Path,'utf8')
        
        for pathinfile, subdirs, files in betterwalk.walk(Path):
        
            for name in files:
                if Search_Condition in name: 
                    i+=1
                    fullPath = os.path.join(pathinfile,name)
                    df.loc[i, 'Path']=fullPath
                    df.loc[i, 'File Name']=name

        #drop N/A          
        df = df[(pd.notnull(df['File Name']))]

    #Search => OR
    if Search_Type == 'OR':
        
        
        SearchORArr=Search_Condition.split(',')

        i=(-1)
        for pathinfile, subdirs, files in betterwalk.walk(Path):
        
            for name in files:
                ORresult = map(lambda x:re.findall(x,name),SearchORArr)
                if not isListEmpty(ORresult): 
                    i+=1
                    fullPath = os.path.join(pathinfile,name)
                    df.loc[i, 'Path']=fullPath
                    df.loc[i, 'File Name']=name

        #drop N/A          
        df = df[(pd.notnull(df['File Name']))]

    
    #Search => AND
    if Search_Type == 'AND':
        
        
        SearchANDArr=Search_Condition.split(',')

        i=(-1)
        for pathinfile, subdirs, files in betterwalk.walk(Path):
        
            for name in files:
                ANDresult = map(lambda x:re.findall(x,name),SearchANDArr)
                if ANDCheckEmpty(ANDresult)== True: 
                    i+=1
                    fullPath = os.path.join(pathinfile,name)
                    df.loc[i, 'Path']=fullPath
                    df.loc[i, 'File Name']=name

        #drop N/A          
        df = df[(pd.notnull(df['File Name']))]
        

    if df.empty:
        return ('No Results')
    os.chdir('//ecsbks01/swap/DDD00/virtualenv/WinPython-32bit-2.7.10.2/python-2.7.10/Project_Evaluate_Excel/Search_History')
    #Search for files
    #word1=Search_Condition.decode('utf-8')
    #df['Search Result']=df['File Name'].str.contains(Search_Condition)
    #result = df[(df['Search Result']==True)]
    #search for files write into excel
    write_df=df.loc[:,['File Name','Path']]
    writer = ExcelWriter('Result-Output.xls')
    write_df.to_excel(writer,'Result',index=False)
    
    writer.save()







    #turn search to files into hyperlink
    CWPath = '\\\\ecsbks01\\swap\\DDD00\\virtualenv\\WinPython-32bit-2.7.10.2\\python-2.7.10\\Project_Evaluate_Excel\\Search_History'
    Excel_Path = os.path.join(CWPath, 'Result-Output.xls')
    wb = Workbook(Excel_Path)
    wb = Workbook.caller()
    checkArr = Range('B2').vertical.value
    i = 2
    for check in checkArr:
    
        RangeName=('B%d' % (i))
        displayRange=('A%d' % (i))
        address=Range(RangeName).value
        display_name = Range(displayRange).value
        i+=1
        try:
            Range(RangeName).add_hyperlink(address, text_to_display=address)
        except:
            pass
    return "FINISH"
   



