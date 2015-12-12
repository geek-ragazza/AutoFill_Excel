# coding: utf-8



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




#Read data
con = lite.connect("Evaluate.sqlite3")
sql_exe=("SELECT * from Cable_Data")
df_cable = pd.read_sql_query(sql_exe, con)
df_cable.dropna(how='any')

sql_cableinfo=("SELECT * from Cable")
df_cableINFO = pd.read_sql_query(sql_cableinfo, con)
df_cableINFO.dropna(how='any')

sql_wage = ("SELECT * from Wages")
df_wage = pd.read_sql_query(sql_wage, con)
df_wage.dropna(how='any')

sql_pro = ("SELECT * from Project")
df_pro = pd.read_sql_query(sql_pro, con)
df_pro.dropna(how='any')

#Read Daily Price
con2 = lite.connect("daily_price.db")
sql_price = ("SELECT * from daily")
df_dprice = pd.read_sql_query(sql_price, con2)
df_dprice.dropna(how='any')
df_dprice = df_dprice.sort_values(['Record_Date'],ascending=1)
df_dprice = df_dprice[(pd.notnull(df_dprice['LME_Copper']) & pd.notnull(df_dprice['USD']) & pd.notnull(df_dprice['Oil']))]





#發包銅指數
def FBCopper(row):
    temp_df = df_pro[(df_pro['ID']==row['Pro_ID'])]
    #print temp_df
    try:
        result = temp_df.get_value(temp_df.index[0],'Copper')
        return result
    except:
        pass




#發包匯率
def FBCurrency(row):
    temp_df = df_pro[(df_pro['ID']==row['Pro_ID'])]
    #print temp_df
    try:
        result = temp_df.get_value(temp_df.index[0],'Currency')
        return result
    except:
        pass




#芯數
def CPQty(row):
    temp_df = df_cableINFO[(df_cableINFO['ID']==row['Cable_ID'])]
    try:
        result = temp_df.get_value(temp_df.index[0],'CP')
        
        if result[-1]=='C':
            #print result[0:-1]
            #print int(result[0:-1])
            return int(result[0:-1])
        if result[-1]=='P':
            return int(result[0:-1])*2
    except:
        pass




#工資面積   
def WageArea(row):
    if row['Size']<3.5:
        return row['Size'] * row['CP_Calculate']
    else:
        return np.sqrt((row['Size'] * row['CP_Calculate']))




#工資類型 價格    
def WageCate(row):
    temp_df = df_wage[(df_wage['ID']==row['Wage_Type'])]
    try:
        result = temp_df.get_value(temp_df.index[0],'Price')
        
        return result
    except:
        pass


def ApplySubCate(row):
    temp_df = df_cableINFO[(df_cableINFO['ID'] == row['Cable_ID'])]
    try:
        result = temp_df.get_value(temp_df.index[0],'Sub_ID')
        
        return result
    except:
        pass


#等比例發包價
def DBLFBJ(row):
    
    temp_df= grouped[(grouped.index == row['Material_Name'])]
    
    #temp_df = grouped[(grouped.index == row['Sub_ID'])]
    try:
        result = temp_df.get_value(temp_df.index[0],'Base_FB')
        
        return row['List_Price']*result
    except:
        return row['List_Price']
    




#修正銅指數及匯率單價
def CorrectPrice(row):
    
    try:
        Current_Copper = df_dprice.get_value(df_dprice.index[-1],'LME_Copper')
        Current_Currency = df_dprice.get_value(df_dprice.index[-1],'USD')
        c1 = row['DBLFBJ']*(1-row['Copper_Percentage']) 
        c2 = row['DBLFBJ']*row['Copper_Percentage']
        copper = Current_Copper/row['FB_copper']
        currency = Current_Currency /row['FB_currency']
        return c1 +(c2*copper*currency)
    
    except:
        pass
    
#發包銅指數    
df_cable['FB_copper'] = df_cable.apply(FBCopper, axis=1)
#發包匯率
df_cable['FB_currency'] = df_cable.apply(FBCurrency, axis=1)
#實際芯數
df_cable['CP_Calculate']  = df_cable.apply(CPQty, axis=1)
#工資面積
df_cable['Wage_Area'] =  df_cable.apply(WageArea, axis=1)
#工資類型 價格
df_cable['Wage_Val'] =  df_cable.apply(WageCate, axis=1)
#工資
df_cable['Wage_Result'] = np.round((df_cable['Wage_Area']*df_cable['Wage_Val']), decimals=1)

















df_cable['Material_Name'] =  df_cable.apply(ApplySubCate, axis=1)

df_RLPrice = df_cable[(pd.notnull(df_cable['List_Price']) & pd.notnull(df_cable['Real_Price']))]
grouped =  df_RLPrice.groupby(['Material_Name']).sum()
#基價發包
grouped['Base_FB'] =  grouped['Real_Price']/grouped['List_Price']



df_cable['DBLFBJ'] = df_cable.apply(DBLFBJ, axis=1)
df_cable['Correct_Price'] = df_cable.apply(CorrectPrice, axis=1)
print df_cable.tail()
def CableFormula(ID):
    
    select = df_cable[df_cable['ID']==ID]
    Wage = select.get_value(select.index[0],'Wage_Val')
    SPrice = select.get_value(select.index[0],'Correct_Price')
    SCopper = select.get_value(select.index[0],'Copper_Percentage')
    #print SPrice*(1-SCopper), "+", SPrice * SCopper
    AF = SPrice*(1-SCopper)
    BF = SPrice * SCopper
    FormulaString = ("=ROUND(%f+%f*單價基準!B4 / 單價基準!B3 * 單價基準!H4 / 單價基準!H3,1)"%(AF,BF))
    ReturnArr = [FormulaString,Wage,Pro_Name]
    Pro_ID =select.get_value(select.index[0],'Pro_ID')
    Select_Pro = df_pro[(df_pro['ID']==Pro_ID)]
    Pro_Name = Select_Pro.get_value(Select_Pro.index[0],'Name')
    print Pro_Name
    return ReturnArr


print CableFormula(23)



