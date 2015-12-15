import pandas as pd
import sqlite3 as lite
import numpy as np
from pandas import ExcelWriter
from pandas import ExcelFile
import datetime
import gc
import os
from xlwings import Workbook, Sheet, Range, Chart


def SearchTarget(SearchString):
    SearchArr =[]
    SearchArr = SearchString.split(',')
    for searchrangeI in xrange(1,2434414,10000):
        #start time
        
        #>= searchrangeI <searchrangeI+1000
        #Read data
        con = lite.connect("database.db")
        sql_des = ("SELECT * from excel_content where ID>=%d and ID<%d" % (searchrangeI,(searchrangeI+10000)))
        sql_exe=(sql_des)
        df = pd.read_sql_query(sql_exe, con)
        
        #Search Condition 
        for SString in SearchArr:
            #SearchTag = SString.decode('utf-8')
            SearchTag = SString
            SStringID = SearchArr.index(SString)
            
            if SStringID == 0 :
                result_df =  df[df['Cell_Data'].str.contains(SearchTag)]
                
            else:
                result_df = result_df[result_df['Cell_Data'].str.contains(SearchTag)]
                
        
        del df
        gc.collect()
        temp_df = result_df
        if searchrangeI == 1:
            Allresult_df = temp_df
        else:
            conframe = [Allresult_df, temp_df]
            Allresult_df =pd.concat(conframe)
        #print  "END:", datetime.datetime.now()
    return Allresult_df


con = lite.connect('Evaluate.sqlite3')
cony = lite.connect('database.db')
with con and cony:
    cur=con.cursor()
    cury = cony.cursor()
    sql_action='SELECT * FROM Search_Key WHERE ID>250'
    #parameters
    for row in con.execute(sql_action):
        SearchKey=row[2]
        print  "Start:", datetime.datetime.now()
        TargetResult_df = SearchTarget(SearchKey)
        print  "END:", datetime.datetime.now()
        if not TargetResult_df.empty:
            SearchKeyID =  row[0]
            #print TargetResult_df
            #Count_inSheet =  TargetResult_df['Sheet_Name'].value_counts(3)
            
            UniqueSheetName = pd.unique(TargetResult_df.Sheet_Name.ravel())
            for SheetID in UniqueSheetName:
                Occurrences = (TargetResult_df['Sheet_Name']==SheetID).sum()
                #print SheetID, Occurrences, SearchKeyID, 1
                
                Insert_Train = 'INSERT INTO Train_Set(Search_KeyID, Table_ID, Sheet_ID, Occurrences) VALUES (?,?,?,?)'
                Insert_Para = [int(SearchKeyID),1, int(SheetID), int(Occurrences)]
                cony.execute(Insert_Train,Insert_Para)
