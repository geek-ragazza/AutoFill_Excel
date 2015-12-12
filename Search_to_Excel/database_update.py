import shutil
import pandas as pd
from pandas import *
import numpy as np
from pandas import ExcelWriter
from pandas import ExcelFile
from xlwings import Workbook, Sheet, Range, Chart
import sqlite3 as lite
from pandas.io import sql
from dateutil.parser import parse
import betterwalk
import os
import subprocess
from nt import chdir
import gc

def EatExcel(OriginalPath, OriginalName, tempName):
    dstroot = 'C:\\temp_read_Excel'
    os.chdir(dstroot)
    try:  #avoid utf-16 error
        xl = pd.ExcelFile(tempName)
        SheetArr = xl.sheet_names
        for SheetName in SheetArr:
            df = xl.parse(SheetName, header=None)
            df.dropna(how='any')
            ColumnNum = len(df.columns)

            df.columns= range(1,(ColumnNum+1))
            df['ix']=df.index+1
            #df['Sheet_Name']=SheetArr[0]
            #print df.head()


            #excel file
            df_file = pd.DataFrame({'File_SheetName':[SheetName]})
            df_file['PATH'] = OriginalPath
            df_file['FileName'] = OriginalName
            #print df_file
            os.chdir('C:\\Python2.7.10\\Scripts\\notebook\\Eating Data')
            cnx = lite.connect('database.db')
            sql_df=df_file

            #Write into database 'excel_file'
            try:
                sql.write_frame(sql_df, name='excel_file', con=cnx, if_exists='append')

                del df_file 
                gc.collect()
                
            except:

                
                pass

            #Get Sheet_Name ID 
            con = lite.connect('database.db')
            with con:
                cur=con.cursor()
                sql_action='SELECT * FROM excel_file WHERE PATH=? and FileName=? and File_SheetName=? '
                sql_para =[OriginalPath, OriginalName, SheetName]
                for row in con.execute(sql_action,sql_para):
                    df['Sheet_Name'] = row[0]


            for col in range(1,(ColumnNum+1)):
                #print df[col].head()
                #print df.loc[:,[col,'ix','Sheet_Name']]
                Write_df =df.loc[:,[col,'ix','Sheet_Name']]
                Write_df.columns=['Cell_data','Row_Number','Sheet_Name']
                Write_df['Col_Number'] = col
                Write_df = Write_df[Write_df['Cell_data'].isnull()!=True]


                #Write into database 'excel_content'
                try:
                    sql.write_frame(Write_df, name='excel_content', con=cnx, if_exists='append')
                    del write_frame 
                    gc.collect()
                    
                except:

                    
                    pass

            del df 
            gc.collect()
    except:
        pass
            

def UnMsg(pathinfile,name):
    dstroot ='C:\\temp_read_Excel'
    srcfile=os.path.join(pathinfile,name)
    #copy msg to dstroot
    shutil.copy(srcfile, dstroot)
    #locat to dstroot
    os.chdir(dstroot)
    
    #rename msg file
    OriginalMsgName=os.path.join('C:\\', 'temp_read_Excel', name)
    MailmsgPath=os.path.join('C:\\', 'temp_read_Excel', '123143.msg')
    os.rename(OriginalMsgName,MailmsgPath)
    
    #start extract attchment
    DIR = os.path.join('C:\\', 'Python2.7.10', 'msg-extractor-master', 'ExtractMsg.py')
    subprocess.call(['python', DIR, '123143.msg'])
    
    TempNewNamePathXLS=os.path.join('C:\\', 'temp_read_Excel', '123.xls')
    TempNewNamePathXLSX=os.path.join('C:\\', 'temp_read_Excel', '123.xlsx')
    os.chdir(dstroot)
    try:
        os.remove(TempNewNamePathXLSX)
    except:
        pass
    
    try:
        os.remove(TempNewNamePathXLS)
    except:
        pass
    
    for folderName in os.listdir(dstroot):
        if folderName!='123143.msg' and folderName!='chinese_dictionary.db' and folderName!='database.db':
            MsgFolder= os.path.join('C:\\', 'temp_read_Excel', folderName)
            xlsinMsgCount=0
            for MsgfolderFile in os.listdir(MsgFolder):
                fileEXT = os.path.splitext(MsgfolderFile)[1]
                
                if fileEXT=='.xls' or fileEXT=='.xlsx':
                    
                    MEsrcfile=os.path.join(MsgFolder,MsgfolderFile)
                    TempFile=os.path.join('C:\\','temp_read_Excel', MsgfolderFile)
                    
                    shutil.copy(MEsrcfile, dstroot)
                    if fileEXT=='.xls':
                        os.rename(TempFile,TempNewNamePathXLS)
                        # save in database
                        try:
                            EatExcel(srcfile,MsgfolderFile,'123.xls')
                            os.remove(TempNewNamePathXLS)
                        except:
                            os.remove(TempNewNamePathXLS)
                            pass
                    if fileEXT=='.xlsx':
                        # save in database
                        os.rename(TempFile,TempNewNamePathXLSX)
                        try:
                            EatExcel(srcfile,MsgfolderFile,'123.xlsx')
                            os.remove(TempNewNamePathXLSX)
                            
                        except:
                            os.remove(TempNewNamePathXLSX)
                            pass
                        # save in database
                        
                        
    try:
        shutil.rmtree(MsgFolder)
    except:
        pass
    os.remove(MailmsgPath)



newpath='C:\\temp_read_Excel\\'
directory_path='P:\\final_price'
if not os.path.exists(newpath): 
    os.makedirs(newpath)
fileNum=0
for pathinfile, subdirs, files in betterwalk.walk(directory_path):
    for name in files:
        fileEXT = os.path.splitext(name)[1]  #filename extension
        fileNum+=1
        print fileNum , "exception"
        #if the file is msg
        if (fileEXT=='.msg') and fileNum>90:
            MsgPath=os.path.join(pathinfile,name)
            print pathinfile
            try:
                UnMsg(pathinfile,name)
            except:
                pass
            
        #if the file is Excel
        if (fileEXT == '.xls' or fileEXT== '.xlsx') and fileNum>90: #and fileNum<=355 and fileNum>=0:
                if fileEXT=='.xls':
                    tempName='123.xls'
                if fileEXT=='.xlsx':
                    tempName='123.xlsx'
                
                #move the file to the 'temp_read_Excel' folder
                oldNamePath = os.path.join(pathinfile,name)
                
                srcfile = oldNamePath
                dstroot ='C:\\temp_read_Excel'
                
                #copy file
                shutil.copy(srcfile, dstroot)
                TempOldNamePath = os.path.join(dstroot,name)
                TempNewNamePath = os.path.join(dstroot,tempName)
                #rename
                os.rename(TempOldNamePath,TempNewNamePath)
                #Eat Eat Eat
                EatExcel(pathinfile, name, tempName)
                
                #remove
                os.remove(TempNewNamePath)
                
                
