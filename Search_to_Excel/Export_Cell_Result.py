import pandas as pd
import sqlite3 as lite
import numpy as np
from pandas import ExcelWriter
from pandas import ExcelFile
import datetime
import gc
import os
from xlwings import Workbook, Sheet, Range, Chart
def DFtoExcel(df,FolderName,FileName):
    write_df=df.loc[:,['FileName','hyperlink','Sheet Name']]
    
    #Path Cell_Search_By_Key 
    MainFolder = 'C:\\Cell_Search_By_Key'
    FolderPath = os.path.join(MainFolder, FolderName)
    if not os.path.exists(FolderPath):
        os.makedirs(FolderPath)
    os.chdir(FolderPath)
    ExcelName=('%s.xlsx' % FileName)
    writer = ExcelWriter(ExcelName)
    write_df.to_excel(writer,'Result',index=False)
    writer.save()
    #turn path into hyperlink
    Excel_Path = os.path.join(FolderPath, ExcelName)
    wb = Workbook(Excel_Path)
    #wb = Workbook.caller()
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
    wb.save()
    wb.close()
    return "FINISH"

MyPath ='C:\\Python2.7.10\\Scripts\\notebook\\Eating Data'
os.chdir(MyPath)    
con = lite.connect('../cable/Evaluate.sqlite3')

def seperateUno(row):
    if row['href']!=None:
        Arr = row['href']
        return Arr[0]
    else:
        return None
def seperateDue(row):
    if row['href']!=None:
        Arr = row['href']
        return Arr[1]
    else:
        return None
def seperateTre(row):
    if row['href']!=None:
        Arr = row['href']
        return Arr[2]
    else:
        return None
def gethref(row):
    SheetID=int(row['Sheet_ID'])
    os.chdir(MyPath)    
    cony = lite.connect("database.db")
    with cony:
        select_hrefsheet ='SELECT * FROM excel_file WHERE ID=?'
        select_para = [SheetID]
        
        try:
            for sel in cony.execute(select_hrefsheet,select_para):
                sel[2],sel[3]
            if sel[1][-4:]== '.msg':
                
                msgPath = sel[1].split('\\')[:-1]
                for i in range(0,len(msgPath)):
                    if i==0:

                        MAIN = msgPath[i]+'\\'
                        PATH = os.path.join(MAIN)
                    else:
                        PATH = os.path.join(PATH,msgPath[i])
            else:
                PATH = os.path.join(sel[1],sel[2])

            return [PATH,sel[2],sel[3]]
        except:
            return None
            
        
with con:
    cur=con.cursor()
    
    sql_action='SELECT * FROM Search_Key'
    for row in con.execute(sql_action):
        FolderName = row[1]
        FileName = row[2]
        os.chdir(MyPath)    
        conx = lite.connect("database.db")
        sql_exe=("SELECT * from Train_Set WHERE Search_KeyID = %d and Table_ID = 1" % row[0])
        df = pd.read_sql_query(sql_exe, conx)
        
        if not df.empty:
            df.dropna(how='any')
            df['href']=df.apply(gethref,axis=1)
            
            df['hyperlink']=df.apply(seperateUno,axis=1)
            df['FileName']=df.apply(seperateDue,axis=1)
            df['Sheet Name']=df.apply(seperateTre,axis=1)
            try:
                print DFtoExcel(df,FolderName,FileName)
            except:
                print "pass"
                pass
