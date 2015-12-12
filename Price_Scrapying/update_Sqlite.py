# -*- coding: utf-8 -*-
import datetime
import sqlite3 as lite
import array
import scraping
import re
#Scrape Data

lme = scraping.ScrapingLMELogin()
lme.LoginGetValue()
lmeValues =  lme.LoginGetValue()

SHtime=datetime.datetime.now()
SHmin = SHtime.minute
SHhour = SHtime.hour
if SHhour >= 16 and SHmin>=1:
    sh = scraping.ScrapingSHCopper()
    sh.SHGetValue()
    shValues = sh.SHGetValue()

oil = scraping.ScrapingOil()
oil.WTexasOil()

currency = scraping.ScrapingCurrency()
currency.CurrencySelect()
OilDateArr = []
OilPriceArr = []
for exoildate in oil.DateArr:
    
    dateSlice = re.split('[.]', exoildate[1])
    oilYM = dateSlice[0]
    oilday = int(dateSlice[1])
    oildd = '{:02d}'.format(oilday)
    oilyy = oilYM[0:4]
    oilmonth = int(oilYM[4:])
    oilmm = '{:02d}'.format(oilmonth)
    oil_relDate =  ("%s-%s-%s"%(oilyy, oilmm, oildd))
    OilDateArr.append(oil_relDate)
for exoilPrice in oil.PriceArr:
    OilPriceArr.append(exoilPrice[1])

NewOPrice =  OilPriceArr[-1]
NewODate = OilDateArr[-1]

#TIME
DailyNow= datetime.datetime.today()
print DailyNow 
print DailyNow.year
mm = '{:02d}'.format(DailyNow.month)
dd = '{:02d}'.format(DailyNow.day)
date_stamp = ("%s-%s-%s"%(DailyNow.year,mm,dd))
print date_stamp
print DailyNow
#fetch database n write to the html
Hcon = lite.connect('daily_price.db')

"""
#get last two days (no weekend)
monday -3,-4
2 -1 -3
(3,4,5) -1 -2

with Hcon:
    cur=Hcon.cursor()
    sql_action=''
    parameters = [date_stamp]
    Hcon.execute(sql_action,parameters):
"""
       
#write data in html file
f = open('price.html','w')
date = "date"
copper = "copper"
al = "al"
nick = "nick"
zinc = "zinc"
message = """html"""




f.write(message)
f.close()



#SQLite
con = lite.connect('daily_price.db')
with con:
    cur=con.cursor()
    sql_action='SELECT Record_Date FROM daily WHERE Record_Date=?'
    parameters = [date_stamp]
    for row in con.execute(sql_action,parameters):
        print "find",row[0]
        break
    else:
        sql_InsertDate = 'INSERT INTO daily(Record_Date) VALUES (?)'
        con.execute(sql_InsertDate, parameters)
        
    #update LME LOGIN DATA
    sql_LMEaction='SELECT Record_Date FROM daily WHERE Record_Date=?'
    sql_LMEupdate='UPDATE daily SET LME_Copper=?, LME_AL=?, LME_Nick=?, LME_Zinc=? WHERE Record_Date=?'
    LME_parameters = [lmeValues[0]]
    LME_Upparameters = [lmeValues[1],lmeValues[2],lmeValues[3],lmeValues[4],lmeValues[0]]
    for row in con.execute(sql_LMEaction,LME_parameters):
        print "find",row[0]
        con.execute(sql_LMEupdate,LME_Upparameters)
        break
    else:
        sql_LME_InsertDate = 'INSERT INTO daily(Record_Date) VALUES (?)'
        con.execute(sql_LME_InsertDate, LME_parameters)
        con.execute(sql_LMEupdate,LME_Upparameters)
    
    if SHhour >= 16 and SHmin>=1:
        #update ShangHai Copper Price
        sql_SHaction='SELECT Record_Date FROM daily WHERE Record_Date=?'
        sql_SHupdate='UPDATE daily SET Shang_Hai_Copper=? WHERE Record_Date=?'
        SH_parameters = [shValues[0]]
        SH_Upparameters = [shValues[1], shValues[0]]
        for row in con.execute(sql_SHaction, SH_parameters):
            print "find",row[0]
            con.execute(sql_SHupdate,SH_Upparameters)
            break
        else:
            sql_SH_InsertDate = 'INSERT INTO daily(Record_Date) VALUES (?)'
            con.execute(sql_SH_InsertDate, SH_parameters)
            con.execute(sql_SHupdate,SH_Upparameters)
    
    #update CURRENCY
    CUR = currency.CResultArr
    sql_CURaction='SELECT Record_Date FROM daily WHERE Record_Date=?'
    sql_CURupdate='UPDATE daily SET USD=?, JPY=?, EUR=?, CNY=? WHERE Record_Date=?'
    CUR_parameters = [currency.Date]
    CUR_Upparameters = [CUR[0],CUR[1],CUR[2],CUR[3],currency.Date]
    for row in con.execute(sql_CURaction,CUR_parameters):
        print "find",row[0]
        con.execute(sql_CURupdate,CUR_Upparameters)
        break
    else:
        sql_CUR_InsertDate = 'INSERT INTO daily(Record_Date) VALUES (?)'
        con.execute(sql_CUR_InsertDate, CUR_parameters)
        con.execute(sql_CURupdate,CUR_Upparameters)
    #update Oil Price
    
    sql_OILaction='SELECT Record_Date FROM daily WHERE Record_Date=?'
    sql_OILupdate='UPDATE daily SET Oil=?, Update_Datetime=? WHERE Record_Date=?'
    OIL_parameters = [NewODate]
    OIL_Upparameters = [NewOPrice, DailyNow, NewODate]
    for row in con.execute(sql_OILaction,OIL_parameters):
        print "find",row[0]
        con.execute(sql_OILupdate,OIL_Upparameters)
        break
    else:
        sql_OIL_InsertDate = 'INSERT INTO daily(Record_Date) VALUES (?)'
        con.execute(sql_OIL_InsertDate, OIL_parameters)
        con.execute(sql_OILupdate,OIL_Upparameters)

