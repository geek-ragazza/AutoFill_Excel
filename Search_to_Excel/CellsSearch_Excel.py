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
    for searchrangeI in xrange(1,4673319,100000):
        #start time
        print  "Start:", datetime.datetime.now()
        #>= searchrangeI <searchrangeI+1000
        #Read data
        con = lite.connect("database.db")
        sql_des = ("SELECT * from excel_content where ID>=%d and ID<%d" % (searchrangeI,(searchrangeI+100000)))
        sql_exe=(sql_des)
        df = pd.read_sql_query(sql_exe, con)
        
        #Search Condition 
        for SString in SearchArr:
            SearchTag = SString.decode('utf-8')
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
        print  "END:", datetime.datetime.now()
    return Allresult_df
TargetResult_df = SearchTarget("XLPE,電纜,600V")




def TargetFile(Allresult_df):
    def FillTarget(row):
        Target_df = Allresult_df[Allresult_df['Sheet_Name'] == row['ID']]
        #Target_df['Cell_Data']
        result = Target_df.get_value(Target_df.index[0],'Cell_Data')
        return result
    def FillHyperlink(row):
        #print row['PATH'][-4:]
        if row['PATH'][-4:] == '.msg':
            return row['PATH']
        else:
            return row['PATH']+ "\\" + row['FileName']
    con = lite.connect("database.db")
    sql_name = ("SELECT * from excel_file")
    name_df = pd.read_sql_query(sql_name, con)
    name_df.dropna(how='any')
    file_result_df = name_df[name_df['ID'].isin(Allresult_df['Sheet_Name'])]
    file_result_df['Target'] =file_result_df.apply(FillTarget, axis=1)
    #file_result_df['hyperlink']=file_result_df['PATH']+ "\\" + file_result_df['FileName']
    file_result_df['hyperlink'] = file_result_df.apply(FillHyperlink, axis=1)
    return file_result_df
    
    
Finale_df = TargetFile(TargetResult_df) 


write_df=Finale_df.loc[:,['Target','hyperlink','File_SheetName']]
writer = ExcelWriter('Result-Output.xls')
write_df.to_excel(writer,'Result',index=False)
    
writer.save()


#create hyperlink

CWPath = 'C:\\Python2.7.10\\Scripts\\notebook\\Search_Inside_Excel'
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
