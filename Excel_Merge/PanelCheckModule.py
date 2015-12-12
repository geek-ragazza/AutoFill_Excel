
# -*- coding: utf-8 -*-



from xlwings import Workbook, Sheet, Range, Chart
import win32com.client
import re
import pandas as pd
import sqlite3
import numpy as np
import gc

#Excel1, Excel2, TitleCheck1, TitleCheck2, 
#UnoTitleCol, DueTitleCol, UnoPriceCol, DuePriceCol
#UnoStart, UnoEnd, DueStart, DueEnd
def MergenFillPrice(Excel1, Excel2, TitleCheck1, TitleCheck2,
                    UnoTitleCol, DueTitleCol, UnoPriceCol, DuePriceCol,
                    UnoStart, UnoEnd, DueStart, DueEnd):
    
    UnoStart=int(UnoStart)
    UnoEnd=int(UnoEnd)
    DueStart=int(DueStart)
    DueEnd=int(DueEnd)
    
    #Open Original Excel File
    wb = Workbook(Excel1)
    wb = Workbook.caller()


    

    #Get title color
    #[Input] Range
    colorcode = Range(TitleCheck1).color

    
    

    #Create Empty Dataframe for original excel
    columns = ['Title Row Num1', 'Title Text1']
    ArangeLen=(UnoEnd-UnoStart+200)
    index = np.arange(int(ArangeLen))
    df1 = pd.DataFrame(columns=columns, index = index)

    
    
    #Put Title in Dataframe
    j=-1
    
    for i in xrange(int(UnoStart),int(UnoEnd)):
        RangeName=('%s%d' % (UnoTitleCol,i))
        colorread=Range(RangeName).color
        if colorread == colorcode:
            j+=1
            TitleText=Range(RangeName).value
            df1.loc[j, 'Title Row Num1']=i
            df1.loc[j, 'Title Text1']=TitleText
            


    

    #remove the N/A row
    df1 = df1[(pd.notnull(df1['Title Row Num1']))]
    


    

    #Read the Excel including price
    wb2 = Workbook(Excel2)
    wb2 = Workbook.caller()


    

    #Get title color
    #[Input] Range
    colorcode=Range(TitleCheck2).color


    

    #Create Empty Dataframe
    columns=['Title Row Num2', 'Title Text2']
    ArangeLen2=DueEnd-DueStart+200
    index = np.arange(int(ArangeLen2))
    df2 = pd.DataFrame(columns=columns, index = index)


    

    #Put Title in Dataframe
    j=-1
    for i in range(DueStart,DueEnd):
        RangeName=('%s%d' % (DueTitleCol,i))
        colorread=Range(RangeName).color
        if colorread == colorcode:
            j+=1
            TitleText=Range(RangeName).value
            df2.loc[j, 'Title Row Num2']=i
            df2.loc[j, 'Title Text2']=TitleText


    

    #remove the N/A row
    df2 = df2[(pd.notnull(df2['Title Row Num2']))]
    


    

    def ReadItem(row):
        #create item data frame for two excel
        columns1=['Item Row Num1', 'Item Text1']
        index1 = np.arange(600)
        df_item1 = pd.DataFrame(columns=columns1, index = index1)
        columns2=['Item Row Num2', 'Item Text2', 'Item Price2']
        index2 = np.arange(600)
        df_item2 = pd.DataFrame(columns=columns2, index = index2)

        #Read Original Excel File
        wb = Workbook(Excel1)
        wb = Workbook.caller()
        #Get the title row number
        RowANum=row['Title Row Num1']
        RangeText=('%s%d' % (UnoTitleCol,RowANum))
        colorcode=Range(RangeText).color
        count = RowANum
        con=0
        #read the row as item until next title color
        while (con!=colorcode or count==UnoEnd):
            count = count + 1
            RangeText=('%s%d' % (UnoTitleCol,count))
            con=Range(RangeText).color
            ItemText = Range(RangeText).value
            #put item row number and text into dataframe
            df_item1.loc[count-RowANum-1, 'Item Row Num1']=count
            df_item1.loc[count-RowANum-1, 'Item Text1']=ItemText
        #remove the N/A row
        df_item1 = df_item1[(pd.notnull(df_item1['Item Row Num1']))]
        

        #Read the excel which include price
        wb2 = Workbook(Excel2)
        wb2 = Workbook.caller()
        #Get the title row number
        RowBNum=row['Title Row Num2']
        RangeText=('%s%d' % (DueTitleCol,RowBNum))
        colorcode=Range(RangeText).color
        count = RowBNum
        con=0
        #read the row as item until next title color
        while (con!=colorcode or count==DueEnd):
            count = count + 1
            RangeText=('%s%d' % (DueTitleCol,count))
            RangePrice=('%s%d' % (DuePriceCol,count))
            con=Range(RangeText).color
            ItemText = Range(RangeText).value
            ItemPrice = Range(RangePrice).value
            #put item row number, text and price into dataframe
            df_item2.loc[count-RowBNum-1, 'Item Row Num2']=count
            df_item2.loc[count-RowBNum-1, 'Item Text2']=ItemText
            df_item2.loc[count-RowBNum-1, 'Item Price2']=ItemPrice
        
        #remove the N/A row
        df_item2 = df_item2[(pd.notnull(df_item2['Item Row Num2']))]
        
        #Generate the key to merge
        df_item1['key']=df_item1['Item Text1']
        df_item2['key']=df_item2['Item Text2']
        
        #item merge left
        MergeItem = pd.merge(df_item1, df_item2, on='key', how='left')
        
        #garbage collector
        del df_item1
        del df_item2
        gc.collect()

        #Fill the price into the original Excel
        def FillPrice(row):
            #Open original excel
            wb = Workbook(Excel1)
            wb = Workbook.caller()
            
            #fill in  the price
            RangeText=('%s%d' % (UnoPriceCol,row['Item Row Num1']))
            Range(RangeText).value=row['Item Price2']
        

        #Apply fill in price
        MergeItem.apply(FillPrice, axis=1)
        
        del MergeItem
        gc.collect()
        
        


    

    #Generate the key to merge
    df1['key']=df1['Title Text1']
    df2['key']=df2['Title Text2']
    print "Start Merge"
    #title merge inner
    MergeTitle = pd.merge(df1, df2, on='key', how='inner')
    print MergeTitle
    del df1
    del df2
    gc.collect()
    if MergeTitle.empty:
        return "Empty"
    else:
        MergeTitle['Value'] = MergeTitle.apply(ReadItem, axis=1)

        return "FINISH"
    

