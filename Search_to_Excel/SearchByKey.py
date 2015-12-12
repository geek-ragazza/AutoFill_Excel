from xlwings import Workbook, Sheet, Range, Chart
import win32com.client
import re
import pandas as pd
import sqlite3 as lite
import numpy as np
import gc
import os
import datetime
import Search_History


con = lite.connect('Evaluate.sqlite3')
NoFoundArr=[]
with con:
    cur=con.cursor()
    sql_action='Select * From Search_Key'
    for row in con.execute(sql_action):
        print row[2]
        #Create folder
        MainFolder = 'C:\\Search_Result'
        FolderPath = os.path.join(MainFolder, row[1])
        if not os.path.exists(FolderPath):
            os.makedirs(FolderPath)
        
        #Search and generate excel
        path='//Ecs01//ddb00/final_price'
        GetResult = Search_History.SearchFile(path,row[2],'None')
        if GetResult != 'No Results':
            CWPath = '\\\\ecsbks01\\swap\\DDD00\\virtualenv\\WinPython-32bit-2.7.10.2\\python-2.7.10\\Project_Evaluate_Excel\\Search_History'
            Excel_Path = os.path.join(CWPath, 'Result-Output.xlsx')
            wb = Workbook(Excel_Path)
            #wb = Workbook.caller()
            ExcelName = ('%s.xlsx' % row[2])
            ExcelPath = os.path.join(FolderPath,ExcelName)
            wb.save(ExcelPath)
            wb.close()
        else:
            print "No result"
            NoFoundArr.append([row[1],row[2]])




#Write No result tag in to NotFound.xlsx
print NoFoundArr 
NotFoundPath = 'C:\\Search_Result\\Not_Found.xlsx'
wb = Workbook(NotFoundPath)
wb = Workbook.caller()
i=0
for line in NoFoundArr:
    i+=1
    CateR = ('A%d' % i)
    NameR = ('B%d' % i)
    Range(CateR).value =line[0]
    Range(NameR).value =line[1]
wb.save(NotFoundPath)
wb.close()
