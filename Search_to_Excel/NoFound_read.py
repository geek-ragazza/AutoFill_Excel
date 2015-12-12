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

#read in Excel
UpdateArr = []
NotFoundPath = 'C:\\Search_Result\\Not_Found.xlsx'
wb = Workbook(NotFoundPath)
wb = Workbook.caller()
print Range('A1').horizontal.value
AllRow = Range('A1').vertical.value
Max_Row = len(AllRow)
print Max_Row
for i in xrange(1,Max_Row+1):
    RName = ('A%d' % i)
    Col =  Range(RName).horizontal.value
    if len(Col)>=3:
        
        for addcol in xrange(3,len(Col)+1):
            #or (255, 192, 0)
            #and  (0, 176, 240)
            ConColor = Range((i,addcol)).color
            if ConColor == None:
                UpdateArr.append([Col[0], Col[1] , "None", Range((i,addcol)).value])
            if ConColor == (255, 192, 0):
                UpdateArr.append([Col[0], Col[1] , "OR", Range((i,addcol)).value])
            if ConColor == (0, 176, 240):
                UpdateArr.append([Col[0], Col[1] , "AND", Range((i,addcol)).value])
wb.close()

#Save into the SQLite
con = lite.connect('Evaluate.sqlite3')
with con:
    cur=con.cursor()
    for data in UpdateArr:
        #Selete Index ID
        sql_action='Select ID From Search_Key Where Cate=? and Name=?'
        select_para = [data[0],data[1]]
        for select_id in con.execute(sql_action, select_para):
            Search_ID = select_id[0]
            print Search_ID
        sql_action_fix='INSERT INTO Search_Key_Fix (Search_ID, Search_Type, Search_Condition) VALUES (?,?,?)'
        fix_para = [Search_ID, data[2], data[3]]
        try:
            con.execute(sql_action_fix, fix_para)
            print "insert"
        except:
            pass
