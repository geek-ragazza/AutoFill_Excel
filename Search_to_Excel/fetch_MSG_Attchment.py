import pyexcel as pe
import pyexcel.ext.xls # import it to handle xls file
import pyexcel.ext.xlsx # import it to handle xlsx file
import json
import pyexcel
import sqlite3 as lite

from openpyxl import Workbook
from openpyxl import load_workbook
import datetime, time
import os
from nt import chdir
import shutil

import subprocess

#tempName is 123.xls, or 123.xlsx
def WriteInDatabase(pathinfile,name,tempName):
    dstroot ='C:\\temp_read_Excel'
    pathstring='C:\temp_read_Excel'
    filepath_data=unicode(pathinfile)
    
    FileNamestring=name.decode("BIG5")
   
    TempNewNamePath = os.path.join(dstroot,tempName)
    os.chdir(dstroot)
    print os.getcwd()
    while True:
        try:
            book_dict = pyexcel.get_book_dict(file_name=tempName, path=dstroot)
            break
        except ValueError:
            os.remove(TempNewNamePath)
            print "Oops!  That was no valid number.  Try again..."
    #book_dict = pyexcel.get_book_dict(file_name=tempName, path=dstroot)
    #isinstance(book_dict, OrderedDict)
    con = lite.connect('database.db')
    with con:
        cur=con.cursor()
        sheetnum=0
        for key, val in book_dict.items():
            #sheetName_ID
            sheetnum+=1
            
            sql_datacheck="SELECT * FROM excel_file WHERE PATH=? AND FileName=? AND File_SheetName=?"
            sql_pathnamesheet="INSERT INTO excel_file(PATH,FileName,File_SheetName,File_Datetime) VALUES(?,?,?,?)"
            now = datetime.datetime.now()
            PathNameSheet_para=[filepath_data,FileNamestring,key,now]
            fileCheck_Para=[filepath_data,FileNamestring,key]
            for checker in con.execute(sql_datacheck, fileCheck_Para):
                rownum=0
                for valrow in val:
                    rownum+=1
                    colnum=0
                    for valcol in valrow:
                        colnum+=1
                        if valcol!='':
                            sql_check="SELECT * FROM excel_content WHERE Sheet_Name=? and Row_Number=? and Col_Number =?"
                            sql_update="UPDATE excel_content SET Cell_Data=? WHERE Sheet_Name=? and Row_Number=? and Col_Number =?"
                            sql_insert="INSERT INTO excel_content(Sheet_Name,Row_Number,Col_Number,Cell_Data) VALUES(?,?,?,?)"
                            content_check_para=[checker[0],rownum,colnum]
                            insert_para=[checker[0],rownum,colnum,valcol]
                            update_para=[valcol,checker[0],rownum,colnum]
                            #cell exist
                            for celldataval in con.execute(sql_check, content_check_para):
                                con.execute(sql_update, update_para)
                                        
                                break
                            else:
                                    
                                con.execute(sql_insert, insert_para)
                             
                #os.remove(TempNewNamePath)
                break
            else:    
                con.execute(sql_pathnamesheet, PathNameSheet_para)
                for thisID in con.execute(sql_datacheck, fileCheck_Para):
                            
                    rownum=0
                        
                    for valrow in val:
                        rownum+=1
                        colnum=0
                        for valcol in valrow:
                            colnum+=1
                            if valcol!='':
                                sql_check="SELECT * from excel_content"
                                sql_insert="INSERT INTO excel_content(Sheet_Name,Row_Number,Col_Number,Cell_Data) VALUES(?,?,?,?)"
                                insert_para=[thisID[0],rownum,colnum,valcol]
                                con.execute(sql_insert, insert_para)

def MsgToExcelDatabase(srcfilePath,fileName):
    
    dstroot ='C:\\temp_read_Excel'
    srcfile=os.path.join(srcfilePath,fileName)
    #copy msg to dstroot
    shutil.copy(srcfile, dstroot)
    #locat to dstroot
    os.chdir(dstroot)
    #rename msg file
    OriginalMsgName=os.path.join('C:\\', 'temp_read_Excel', fileName)
    MailmsgPath=os.path.join('C:\\', 'temp_read_Excel', '123143.msg')
    os.rename(OriginalMsgName,MailmsgPath)
    #start extract attchment
    DIR = os.path.join('C:\\', 'Python2.7.10', 'msg-extractor-master', 'ExtractMsg.py')
    subprocess.call(['python', DIR, '123143.msg'])
    
    TempNewNamePathXLS=os.path.join('C:\\', 'temp_read_Excel', '123.xls')
    TempNewNamePathXLSX=os.path.join('C:\\', 'temp_read_Excel', '123.xlsx')
    os.chdir(dstroot)
    
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
                            WriteInDatabase(srcfile,MsgfolderFile,'123.xls')
                            os.remove(TempNewNamePathXLS)
                        except:
                            os.remove(TempNewNamePathXLS)
                            pass
                    if fileEXT=='.xlsx':
                        
                        os.rename(TempFile,TempNewNamePathXLSX)
                        try:
                            WriteInDatabase(srcfile,MsgfolderFile,'123.xlsx')
                            os.remove(TempNewNamePathXLSX)
                            
                        except:
                            os.remove(TempNewNamePathXLSX)
                            pass
                        # save in database
                        
                        
    shutil.rmtree(MsgFolder)
    os.remove(MailmsgPath)
