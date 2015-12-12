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
with con:
    cur=con.cursor()
    select_fix = 'SELECT * FROM Search_Key_Fix'
    for row in con.execute(select_fix):
        path='//Ecs01//ddb00/final_price'
        GetResult = Search_History.SearchFile(path,row[3],row[2])
        #if have data
        if GetResult != 'No Results':
            #Select Folder
            select_ID = "SELECT * FROM Search_Key WHERE ID=?"
            ID_para = [row[1]]
            for MainKey in con.execute(select_ID,ID_para):
                FolderName = MainKey[1]
                XlsName = ('%s.xlsx' % MainKey[2])
                print FolderName, XlsName
                
                #Select Folder or Create Folder
                MainFolder = 'C:\\Search_Result'
                FolderPath = os.path.join(MainFolder, FolderName)
                if not os.path.exists(FolderPath):
                    os.makedirs(FolderPath)
                TargetPath = os.path.join(MainFolder,FolderName,XlsName)
                CWPath = '\\\\ecsbks01\\swap\\DDD00\\virtualenv\\WinPython-32bit-2.7.10.2\\python-2.7.10\\Project_Evaluate_Excel\\Search_History'
                NewSheetName=row[3]+row[2]
                
                try:
                    wbTarget = Workbook(TargetPath)
                    Excel_Path = os.path.join(CWPath, 'Result-Output.xlsx')
                    wb = Workbook(Excel_Path)
                    
                    #rename worksheet
                    Sheet('Result',wkb=wb).name=NewSheetName
                    wb.set_current()
                    #Copy All Range
                    AllRRow = len(Range('A1').vertical.value)
                    AllRCol = len(Range('A1').horizontal.value)
                    print AllRCol,AllRRow
                    RangeLimit = ('A1:B%d' % AllRRow)
                    TempData = Range(RangeLimit).value
                    wbTarget.set_current()
                    lastSheetVal = Sheet.count()
                    OriginSArr=[]
                    for SheetName in xrange(1,lastSheetVal+1):
                        OriginSArr.append(Sheet(SheetName).name)
                    if NewSheetName not in OriginSArr:
                        print "I should Add new array"
                        Sheet.add('abc',after=lastSheetVal)
                        Sheet('abc').name=NewSheetName
                        #Paste Data
                        Range(RangeLimit).value=TempData
                        #hypelink
                        for hyperRow in xrange(2,AllRRow):
    
                            RangeName=('B%d' % (hyperRow))
                            address=Range(RangeName).value
                            
        
                            try:
                                Range(RangeName).add_hyperlink(address, text_to_display=address)
                            except:
                                pass
                    wbTarget.save()
                    wb.close()
                    wbTarget.close()
                except:
                    #No Main Excel Exist, Create New One
                    
                    Excel_Path = os.path.join(CWPath, 'Result-Output.xlsx')
                    wb = Workbook(Excel_Path)
                    #rename worksheet
                    Sheet('Result',wkb=wb).name=NewSheetName
                    
                    wb.save(TargetPath)
                    wb.close()
                    pass
                
