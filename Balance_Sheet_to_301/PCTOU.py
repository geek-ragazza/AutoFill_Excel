# -*- coding: utf-8 -*-
from xlwings import Workbook, Sheet, Range, Chart
import win32com.client
import re
import pandas as pd
import sqlite3
import numpy as np
import gc
import os
import datetime
from datetime import date, timedelta

#variable

#Tre0Uno_Sheet_Path
#Balance_Sheet_Path=r'C:\Python2.7.10\Scripts\notebook\xlwings\Balance_Sheet.xls'
#Tre0Uno_Sheet_Path=r'C:\Python2.7.10\Scripts\notebook\xlwings\301.xls'
#ManHour_Path=r'C:\Python2.7.10\Scripts\notebook\xlwings\MH.xls'
#ACHour_Path=r'C:\Python2.7.10\Scripts\notebook\xlwings\fa.xlsx'
#Pro_Name = ['14A1701A', '14C1701A', '14P1701A', '14E1701A']

def execute(Balance_Sheet_Path, Tre0Uno_Sheet_Path, ManHour_Path, ACHour_Path, Pro_Name):
    # # Balance Sheet

    wb = Workbook(Balance_Sheet_Path)
    wb = Workbook.caller()
    SheetA = u'預算平衡表-Phase Summary A'
    SheetE = u'預算平衡表-Phase Summary E'
    SheetP = u'預算平衡表-Phase Summary P'
    SheetC = u'預算平衡表-Phase Summary C'

    SheetArr=[SheetA, SheetE, SheetP, SheetC]
    SheetAlphabetArr = ['A', 'E', 'P','C']


    ColumnAlphabet=['D','E','F','G','H','I', 'J', 'K', 'L', 'M',
                    'N', 'O', 'P']

    AllRow = []
    for SName in SheetArr: 
        
        AllRange = Range(SName, 'D10').vertical.get_address()
        LastRow = AllRange.split(':')[1].split('$')[2]
        LastRow = int(LastRow)
        AllRow.append(LastRow)
    print AllRow
    for SName in SheetArr:
        print SheetArr.index(SName)
    #Create Empty Dataframe for Balance Sheet
    columns = ['Req No', 'Req Description', 'A', 'W', 'Z', 
               'Current Budget', 'Exe Budget', 'C', 
               'PO Num', 'Excepted Budget', 'EAC Now', 'G', 'H']
    ArangeLen=(sum(AllRow)+100)
    index = np.arange(int(ArangeLen))
    df_bs = pd.DataFrame(columns=columns, index = index) 




    def BSreadRange(row, SheetIX):
        SName = SheetArr[SheetIX]
        SAlphabet = SheetAlphabetArr[SheetIX]
        Val = Range(SName, row).value
        if (Val != 0):
            for j in xrange(0,13):
                if(row[0] == ColumnAlphabet[j]):
                    rowNum = '%s%d' % (SAlphabet, int(row[1:]))
                    df_bs.loc[rowNum, columns[j]]=Val
                        
    for SName in SheetArr:   
        SheetIX = SheetArr.index(SName)
        SheetLastRow = AllRow[SheetIX]
        for i in xrange(11,(SheetLastRow+1)):
            CellNameArr = map(lambda x:('%s%d' % (x,i)), ColumnAlphabet)
            [BSreadRange(x, SheetIX) for x in CellNameArr]
        
        

    df_bs = df_bs[(pd.notnull(df_bs['Req No']))]    


    # # 讀取工時



    d = datetime.date.today()- timedelta(days=20)
    SheetMonth = '{:02d}'.format(d.month)
    SheetYear = format(d.year)[2:]
    MH_SheetName=('%s%s' % (SheetYear, SheetMonth))
    print MH_SheetName
    MH_Arr=[]

    #Pro_Name = ['14A1701A', '14C1701A', '14P1701A', '14E1701A']

    #ManHour_Path=r'C:\Python2.7.10\Scripts\notebook\xlwings\MH.xls'
    wb = Workbook(ManHour_Path)
    wb = Workbook.caller()
    AllRange_MH = Range(MH_SheetName, 'A6').vertical.get_address()
    LastRow_MH = AllRange_MH.split(':')[1].split('$')[2]
    LastRow_MH = int(LastRow_MH)
    print LastRow_MH
    for i in xrange(6,LastRow_MH+1):
        SearchLocate = ('A%d'% (i))
        
        s = Range(MH_SheetName, SearchLocate).value
        if s in Pro_Name:
            TargetLocate = ('B%d'% (i))
            MH_Arr.append([s[2],  Range(MH_SheetName, TargetLocate).value])
    print MH_Arr
    Workbook(ManHour_Path).close()
    
    # # 其他費用



    #ACHour_Path=r'C:\Python2.7.10\Scripts\notebook\xlwings\fa.xlsx'
    ACSheetName='acc303'
    AC_Arr=[]

    wb = Workbook(ACHour_Path)
    wb = Workbook.caller()
    AllRange_ACC = Range(ACSheetName, 'A2').vertical.get_address()
    LastRow_ACC = AllRange_ACC.split(':')[1].split('$')[2]
    LastRow_ACC = int(LastRow_ACC)
    print LastRow_ACC


    for i in xrange(1,LastRow_ACC+1):
        SearchLocate = ('A%d'% (i))
        
        s = Range(ACSheetName, SearchLocate).value
        if s in Pro_Name:
            CatLocate = ('D%d'% (i))
            TargetLocate = ('H%d'% (i))
            AC_Arr.append([s[2], Range(ACSheetName, CatLocate).value,
                           Range(ACSheetName, TargetLocate).value])

    print len(AC_Arr)
    columns_AC = ['Type', 'Amount']
    ACrangeLen=(len(AC_Arr))
    index_AC = np.arange(int(ACrangeLen))
    df_ac = pd.DataFrame(columns=columns_AC, index = index_AC)
    j=0
    for i in AC_Arr:
        
        df_ac.loc[j, 'Type']=('%s%d' % (i[0],i[1]))
        
        df_ac.loc[j, 'Amount']=i[2]
        j+=1
    #print df_ac
    AC_Result_df = df_ac.groupby(by=['Type'])['Amount'].sum()

    for i in range(0,len(AC_Result_df)-1): 
        print AC_Result_df.index[i], AC_Result_df[i]
    

    Workbook(ACHour_Path).close()
    # # 301表 




    wb = Workbook(Tre0Uno_Sheet_Path)
    wb = Workbook.caller()

    ColumnAlphabet301 = ['B', 'C', 'D', 'E', 'F', 'G', 'H']
    AllRange301 = Range('D10').vertical.get_address()
    LastRow301 = AllRange301.split(':')[1].split('$')[2]
    LastRow301 = int(LastRow301)
    LastRow301 = LastRow *3
    Tre0Uno_columns = ['PO Num', 'Req No', 'Req Description', 'A', 
                       'Current Budget', 'Exe Budget', 'C']
    Tre0UnorangeLen=(LastRow301+100)
    Tre0Uno_index = np.arange(int(Tre0UnorangeLen))
    df_T0U = pd.DataFrame(columns=Tre0Uno_columns, index = Tre0Uno_index) 


    # #### 讀取原始301資料



    #Read 301 data into the Dataframe function
    def T0UreadRange(row):
        Val = Range(row).value
        if (Val != 0):
            for j in xrange(0,7):
                if(row[0] == ColumnAlphabet301[j]):
                    rowNum = int(row[1:])
                    df_T0U.loc[rowNum, Tre0Uno_columns[j]]=Val
                    df_T0U.loc[rowNum, 'RowID']=rowNum
                    
                    
    for i in xrange(7,(LastRow301+1)):
        #Change This month EAC to last month EAC
        BeforeEACName = ('J%d' % (i))
        AfterEACName = ('K%d' % (i))
        Range(AfterEACName).value = Range(BeforeEACName).value
        
        #Generate all columns name
        CellNameArr301 = map(lambda x:('%s%d' % (x,i)), ColumnAlphabet301)
        
        #Read 301 data into the Dataframe action
        map(T0UreadRange, CellNameArr301)  
        
    #drop N/A    
    df_T0U = df_T0U[(pd.notnull(df_T0U['Req Description']))]    


    # #### 寫入Balance Sheet



    def BSintoT0U(row):
        XColumnName = ['PO Num_x', 'A_x', 'Current Budget_x', 'Exe Budget_x', 'C_x']
        YColumnName = ['PO Num_y', 'A_y', 'Current Budget_y', 'Exe Budget_y', 'C_y']
        for checkCol in YColumnName:
            Col_IX = YColumnName.index(checkCol)
            bs_Val = row[checkCol]
            
            T0U_Val = row[XColumnName[Col_IX]]
            
            if (bs_Val != None) and (pd.isnull(bs_Val) != True) and (bs_Val != T0U_Val):
                
                #301 Column Name to Column Alphabet
                ColAnchor = Tre0Uno_columns.index(checkCol[:-2])
                Input_Location = ('%s%d' % (ColumnAlphabet301[ColAnchor], row['RowID']))
                
            
                Range(Input_Location).value = bs_Val
                Range(Input_Location).color = (136, 153, 238)



    #Generate the key to merge
    df_T0U['key']=df_T0U['Req No']
    df_bs['key']=df_bs['Req No']

    #merge left
    MergeItem = pd.merge(df_T0U, df_bs, on='key', how='left')
    #if Balance sheet has diff between 301
    diff =  MergeItem[(MergeItem['Current Budget_x'] != MergeItem['Current Budget_y']) | 
                      (MergeItem['PO Num_x'] != MergeItem['PO Num_y']) | 
                      (MergeItem['Exe Budget_x'] != MergeItem['Exe Budget_y']) |
                      (MergeItem['C_x'] != MergeItem['C_y'])]

    #Merge and Fill in
    diff['RELSULT'] = diff.apply(BSintoT0U, axis=1)


    # #### 填入其他費用及自辦工時



    FeeArr=[u'其他費用',u'自辦工時', u'自辦工時(MH)', u'間接分攤']
    FeeSymbol=['3','4','mh', '5']
    TypeAnchor = 0

    FillInQuery=[]

    mhTitle =  map(lambda x:('%smh' % (x[0])), MH_Arr)

    for i in xrange(7,(LastRow301+1)):
        CheckRName = ('D%d' % (i))
        CheckTName = ('A%d' % (i))
        InputName = ('H%d' % (i))
        CheckRValue = Range(CheckRName).value
        if  CheckRValue in FeeArr:
            CheckTValue = Range(CheckTName).value
            
            if CheckTValue != None:
                TypeAnchor = CheckTValue[2]
                SymbolIX = FeeArr.index(CheckRValue)
                
                SymbolName = ('%s%s' % (TypeAnchor, FeeSymbol[SymbolIX] ))
                
                mhComment = Range(InputName).comment
                FillInQuery.append([SymbolName,CheckRName,InputName, mhComment])
            else:
                SymbolIX = FeeArr.index(CheckRValue)
                SymbolName = ('%s%s' % (TypeAnchor, FeeSymbol[SymbolIX] ))
                mhComment = Range(InputName).comment
                FillInQuery.append([SymbolName,CheckRName,InputName, mhComment])
    print FillInQuery

    #Fill in MH
    QueryTitleAnchor = map(lambda x:x[0], FillInQuery)
    for mhcheck in QueryTitleAnchor:
        if mhcheck in mhTitle:
            mhTrueIX = mhTitle.index(mhcheck)
            mhQueryIX = QueryTitleAnchor.index(mhcheck)
            FillinTargetRange = FillInQuery[mhQueryIX][2]
            FillinTargetComment =FillInQuery[mhQueryIX][3]
            MH_Data = int(MH_Arr[mhTrueIX][1])
            print mhcheck, MH_Arr[mhTrueIX][1] ,FillInQuery[mhQueryIX][2]
            #Check the Comment Month
            if str(format(d.month)) not in FillinTargetComment:
                # Update and Change the Comment
                print "Update"
                OValue = Range(FillinTargetRange).value
                Range(FillinTargetRange).value = OValue + MH_Data
                Range(FillinTargetRange).comment = str(format(d.month))
            

    #fill in other fee
    for i in range(0,len(AC_Result_df)-1): 
        
        if AC_Result_df.index[i] in QueryTitleAnchor:
            QueryFeeIX = QueryTitleAnchor.index(AC_Result_df.index[i])
            QueryTargetRange = FillInQuery[QueryFeeIX][2]
            FeeFillInValue = AC_Result_df[i]
            Range(QueryTargetRange).value = FeeFillInValue
            Range(QueryTargetRange).color = (136, 153, 238)
            print FillInQuery[QueryFeeIX][0], QueryTargetRange, FeeFillInValue

