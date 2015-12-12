# -*- coding: utf-8 -*-
import urllib2
from HTMLParser import HTMLParser
from array import *
from Tkinter import *
import Tkinter as tk
import requests
import requests.auth
from bs4 import BeautifulSoup
from ntlm import HTTPNtlmAuthHandler
from datetime import date, timedelta
import time
import datetime
from selenium import webdriver
from selenium import *
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
import sqlite3 as lite

#Oil Price
#driver = webdriver.PhantomJS()
driver = webdriver.Chrome()
driver.get("http://web3.moeaboe.gov.tw/oil102/oil1022010/A00/Oil_Price2.asp")
OilDate = driver.find_element_by_xpath('/html/body/table/tbody/tr[3]/td/table[1]/tbody/tr[1]/td[9]').text
OilPrice = driver.find_element_by_xpath('/html/body/table/tbody/tr[3]/td/table[1]/tbody/tr[2]/td[8]').text
#Format Date
OilDatedot = OilDate.replace("\n", ".")
OilDateHyphen = OilDatedot.replace(".", "-")

#Shang Hai Copper
driver.get("http://market.cnal.com/changjiang/")
for i in xrange(2,10):
    TitleTimeXPath = ('/html/body/div[7]/div[3]/div[1]/div/ul/li[%d]/span' % i)
    SHTitleTime = driver.find_element_by_xpath(TitleTimeXPath).text
    SHDT = SHTitleTime.split(' ')
    SHTime=SHDT[1]
    TimeSeparate =SHTime.split(':')
    if int(TimeSeparate[0])>=15:
        SHDate = SHDT[0]
        TitleXPath =('/html/body/div[7]/div[3]/div[1]/div/ul/li[%d]/a' % i)
        SHNextHref =  driver.find_element_by_xpath(TitleXPath).get_attribute('href')
        break
    #get link and date after 15:00
driver.get(SHNextHref) 
SH_CopperVal = driver.find_element_by_xpath('/html/body/div[7]/div[3]/table/tbody/tr[2]/td[4]').text


#LME Price

driver.get("https://secure.lme.com/Data/Community/Login.aspx")
driver.find_element_by_id('_logIn__userID').send_keys("USERNAME")
driver.find_element_by_id('_logIn__password').send_keys("PSWORD")
driver.find_element_by_id('_logIn__logIn').click()
#enter the page
driver.find_element_by_id('_subMenu__dailyStocksPricesMetals').click()
date = driver.find_element_by_xpath("//*[@id='Table3']/tbody/tr[5]/td/table/tbody/tr[6]/td[1]").text
        
Copper = driver.find_element_by_xpath("//*[@id='Table3']/tbody/tr[5]/td/table/tbody/tr[7]/td[8]").text
Aluminium = driver.find_element_by_xpath("//*[@id='Table3']/tbody/tr[5]/td/table/tbody/tr[7]/td[6]").text
Nickel = driver.find_element_by_xpath("//*[@id='Table3']/tbody/tr[5]/td/table/tbody/tr[7]/td[12]").text
Zinc = driver.find_element_by_xpath("//*[@id='Table3']/tbody/tr[5]/td/table/tbody/tr[7]/td[16]").text

date1 = date.encode("utf-8")
        
dateConvert = ("%s-%s-%s"%(date1[11:], date1[8:10], date1[5:7]))
        
driver.quit()


LMEArr=[Copper.encode('utf-8'), Aluminium.encode('utf-8') ,Nickel.encode('utf-8'), Zinc.encode('utf-8')]


#Currency

CResultArr=[]
user = 'DOMAIN\\USERNAME'
password = "PSWORD"
url = "http://www.ctci.com.tw/Acc_Rep/rate/rate.asp"
passman = urllib2.HTTPPasswordMgrWithDefaultRealm()
passman.add_password(None, url, user, password)
# create the NTLM authentication handler
auth_NTLM = HTTPNtlmAuthHandler.HTTPNtlmAuthHandler(passman)
# create and install the opener
opener = urllib2.build_opener(auth_NTLM)
urllib2.install_opener(opener)
data="""
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" >
<soapenv:Header/>
<soapenv:Body>
<ns:FromTimestamp>2012-05-10</ns:FromTimestamp>
</soapenv:Body>
</soapenv:Envelope>
"""
headers={"SOAPAction":"SomeSoapFunc","Content-Type":"text/xml;charset=UTF-8"}
req = urllib2.Request(url, data, headers)
response=urllib2.urlopen(req)
# retrieve the result
parser = response.read()
soup = BeautifulSoup(parser, "html.parser")
#fetch last day (not today)
today = datetime.date.today( )
ytd = today - datetime.timedelta(days=1)
ymm =  '{:02d}'.format(ytd.month)
ydd =  '{:02d}'.format(ytd.day)
divYdate = ("%s%s%s"%(ytd.year,ymm,ydd))

            

#get today weekday
#yWeekday = date.today().weekday()
cc = (1, 17, 4, 13)
count = 0  #can fit yesterday
t = 0
        
divArr=[]
for divcheck in soup.findAll('div'):
    div_Class = (divcheck.get('id')).encode('utf8')
    if div_Class[0:3] == "div":
        divArr.append(div_Class)
        
LastDiv = divArr[-1]
#convert to date
CUDate = ("%s-%s-%s"%(LastDiv[3:7],LastDiv[7:9],LastDiv[9:]))
                   

for div in soup.findAll('div'):
    if LastDiv == (div.get('id')).encode('utf8'):
        for table in div.findAll('table', align="left", border="0", cellpadding="3", cellspacing="0"):
            for title in table.findAll('td', {'class':'DET2'}):
                t += 1
                #choose the currency from cc USD JPY EUR CNY
                for checkcc in cc:
                    if ((checkcc*5)-3) == t:
                        Cvalue = (title.text).encode('utf-8')
                        CResultArr.append(Cvalue)
                        
                        
                        
def WriteInSQL(Date,Type,Price):
    DailyNow = datetime.datetime.now()
    
    con = lite.connect('daily_price.db')
    with con:
        cur=con.cursor()
        sql_action='SELECT * FROM daily WHERE Record_Date=?'
        parameters = [Date]

        for row in con.execute(sql_action,parameters):
            print row
            break

        else: 
            #Create
            
            Insert_date = 'INSERT INTO daily (Record_Date) VALUES (?)'
            Insert_date_para =[Date]
            con.execute(Insert_date,Insert_date_para)
            
        if Type == 1: #oil
            print "oil"
            sql_Oil='UPDATE daily SET Oil=?, Update_Datetime=? WHERE Record_Date=?'
            OIL_para = [Price, DailyNow, Date]
            con.execute(sql_Oil,OIL_para)
        if Type == 2: #ShangHai
            sql_SH='UPDATE daily SET Shang_Hai_Copper=?, Update_Datetime=? WHERE Record_Date=?'
            
            SHL_para = [Price, DailyNow, Date]
            con.execute(sql_SH,SHL_para)
        if Type == 3: #LME
            sql_LME='UPDATE daily SET LME_Copper=?, LME_AL=?, LME_Nick=?, LME_Zinc=?, Update_Datetime=? WHERE Record_Date=?'
            LME_para = Price
            LME_para.append(DailyNow)
            LME_para.append(Date)
            con.execute(sql_LME,LME_para)
        if Type == 4: #Currency
            sql_CU='UPDATE daily SET USD=?, JPY=?, EUR=?, CNY=?, Update_Datetime=? WHERE Record_Date=?'
            CU_para = Price
            CU_para.append(DailyNow)
            CU_para.append(Date)
            con.execute(sql_CU,CU_para)
                
#Oil
WriteInSQL(OilDateHyphen.encode('utf-8'),1,OilPrice.encode('utf-8'))

#ShangHai Copper
WriteInSQL(SHDate.encode('utf-8'),2,SH_CopperVal.encode('utf-8'))

#LME
WriteInSQL(dateConvert,3,LMEArr)

#Currency
WriteInSQL(CUDate,4,CResultArr)
