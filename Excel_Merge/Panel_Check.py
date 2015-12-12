# -*- coding: utf-8 -*-
from xlwings import Workbook, Sheet, Range, Chart
import win32com.client
import re
import pandas as pd
import sqlite3
import numpy as np
import gc
gc.collect()
wb = Workbook(r'C:\Python2.7.10\project_database\xlwings\NET.xls')
wb = Workbook.caller()
#get title color
colorcode=Range('C2539').color
print colorcode
columns=['Title Row Num1', 'Title Text1']
index = np.arange(5000)
df1 = pd.DataFrame(columns=columns, index = index)
j=-1
#12028
#5435,9137
for i in range(8724,9126):
    RangeName=('C%d' % i)
    colorread=Range(RangeName).color
    if colorread == colorcode:
        j+=1
        TitleText=Range(RangeName).value
        df1.loc[j, 'Title Row Num1']=i
        df1.loc[j, 'Title Text1']=TitleText




df1 = df1[(pd.notnull(df1['Title Row Num1']))]

print "Book one DONE"

wb2 = Workbook(r'C:\Python2.7.10\project_database\xlwings\FINAL_PRICE.xlsx')
wb2 = Workbook.caller()
colorcode=Range('L19').color

print colorcode
#9771
#create empty dataframe
columns=['Title Row Num2', 'Title Text2']
index = np.arange(12000)
df2 = pd.DataFrame(columns=columns, index = index)

#data = pd.DataFrame({"Title Column": range(2000)})
#df.append(data)
j=-1
for i in range(18,5604):
    RangeName=('L%d' % i)
    colorread=Range(RangeName).color
    if colorread == colorcode:
        j+=1
        TitleText=Range(RangeName).value
        df2.loc[j, 'Title Row Num2']=i
        df2.loc[j, 'Title Text2']=TitleText

df2 = df2[(pd.notnull(df2['Title Row Num2']))]
print "Book two DONE"



def my_test(row):

    #create item data frame
    columns1=['Item Row Num1', 'Item Text1']
    index1 = np.arange(250)
    df_item1 = pd.DataFrame(columns=columns1, index = index1)
    columns2=['Item Row Num2', 'Item Text2', 'Item Price2']
    index2 = np.arange(250)
    df_item2 = pd.DataFrame(columns=columns2, index = index2)
    
    #read work book one
    wb = Workbook(r'C:\Python2.7.10\project_database\xlwings\NET.xls')
    wb = Workbook.caller()
    RowANum=row['Title Row Num1']
    RangeText=('C%d' % RowANum)
    colorcode=Range(RangeText).color
    count = RowANum
    con=0
    while (con!=colorcode):
        count = count + 1
        RangeText=('C%d' % count)
        con=Range(RangeText).color
        ItemText = Range(RangeText).value
        
        df_item1.loc[count-RowANum-1, 'Item Row Num1']=count
        df_item1.loc[count-RowANum-1, 'Item Text1']=ItemText
    #print "Good bye!"
    df_item1 = df_item1[(pd.notnull(df_item1['Item Row Num1']))]
    
    
    #read work book two
    wb2 = Workbook(r'C:\Python2.7.10\project_database\xlwings\FINAL_PRICE.xlsx')
    wb2 = Workbook.caller()
    RowBNum=row['Title Row Num2']
    RangeText=('L%d' % RowBNum)
    colorcode=Range(RangeText).color
    
    count = RowBNum
    con=0
    while (con!=colorcode):
        count = count + 1
        RangeText=('L%d' % count)
        RangePrice=('O%d' % count)
        con=Range(RangeText).color
        ItemText = Range(RangeText).value
        ItemPrice = Range(RangePrice).value
        df_item2.loc[count-RowBNum-1, 'Item Row Num2']=count
        df_item2.loc[count-RowBNum-1, 'Item Text2']=ItemText
        df_item2.loc[count-RowBNum-1, 'Item Price2']=ItemPrice
    
    
    df_item2 = df_item2[(pd.notnull(df_item2['Item Row Num2']))]
    

    df_item1['key']=df_item1['Item Text1']
    df_item2['key']=df_item2['Item Text2']
    #print df_item1
    #print df_item2
    MergeItem = pd.merge(df_item1, df_item2, on='key', how='left')
    del df_item1
    del df_item2
    gc.collect()
    #print MergeItem.loc[:,['Item Row Num1','Item Text1','Item Row Num2','Item Price2']]

    #fill in  price
    
    
    def FillPrice(row):
        wb = Workbook(r'C:\Python2.7.10\project_database\xlwings\NET.xls')
        wb = Workbook.caller()
        
        RangeText=('L%d' % row['Item Row Num1'])
        Range(RangeText).value=row['Item Price2']
    MergeItem.apply(FillPrice, axis=1)
    del MergeItem
    gc.collect()
    print row['Title Row Num1'],row['Title Text1'] 
    return row['Title Row Num1'] + row['Title Row Num2']



df1['key']=df1['Title Text1']
df2['key']=df2['Title Text2']
print "Start Merge"
MergeTitle = pd.merge(df1, df2, on='key', how='inner')
del df1
del df2
gc.collect()
print MergeTitle
MergeTitle['Value'] = MergeTitle.apply(my_test, axis=1)



