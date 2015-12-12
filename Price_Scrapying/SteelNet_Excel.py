# -*- coding: utf-8 -*-

from array import *
from datetime import date, timedelta
import time
import datetime
from selenium import webdriver
from selenium import *
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains

#generate random time
from __future__ import division
import random
randomTimeArr = []
for x in range(6):
    MsSecond = random.randint(22291,50000)
    randomTimeArr.append( MsSecond/10000 )


#login
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
#driver = webdriver.PhantomJS()
driver = webdriver.Chrome()
driver.get("https://www.steelnet.com.tw/index.jsp")


driver.maximize_window()
#driver.manage().window().maximize();
time.sleep(randomTimeArr[0])
driver.find_element_by_xpath('//*[@id="close"]').click()
time.sleep(randomTimeArr[1])
driver.find_element_by_name('account').send_keys("USERNAME")
time.sleep(randomTimeArr[2])
driver.find_element_by_name('password').send_keys("PASSWORD")
time.sleep(randomTimeArr[3])
variable = driver.find_element_by_xpath('/html/body/center/table[2]/tbody/tr/td[1]/table/tbody/tr[1]/td/table[1]/tbody/tr[1]/td/table[2]/tbody/tr/td[2]')

#Click too login
actions = ActionChains(driver)
actions.move_to_element(variable)
actions.double_click(variable)
actions.perform()


time.sleep(randomTimeArr[4])


#click Taiwan Daily Price
driver.find_element_by_xpath('/html/body/center/table[2]/tbody/tr/td[3]/table[3]/tbody/tr/td[1]/table[1]/tbody/tr/td/table[2]/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td[1]/a').click()
time.sleep(randomTimeArr[5])  



XlsAlphabetArr=['F', 'G', 'H', 'I']
#Open Excel 
from xlwings import Workbook, Sheet, Range, Chart
import win32com.client

wb = Workbook(r'C:\Python2.7.10\Scripts\notebook\cable\steelnet2015.xls')
wb = Workbook.caller()


#Start Get Table information
#tr range 1 to 48
tdArr = [7,9,7,7,6,8,7,5,6,7,
        7,8,5,5,5, 6,7,6,6,7,
        6,6,6,8,7, 7,8,7,7,7,
        7,6,8,7,7,7,7,8,6,6,
        7,8,7,7,7,7,7,7]
for trVal in xrange(2,49):
    tdValRange =  tdArr[(trVal-1)]
    for tdVal in xrange((tdValRange-3),(tdValRange+1)):
        XPathString = ('/html/body/center/table[3]/tbody/tr/td[3]/table[6]/tbody/tr[%d]/td[%d]' %(trVal, tdVal))
        TEXTVAL = driver.find_element_by_xpath(XPathString).text
        RangeString = ('%s%d' % (XlsAlphabetArr[(tdVal-tdValRange+3)],trVal))
        Range(RangeString).value = TEXTVAL
        Range(RangeString).fontcolor =(0,0,0)
        #check + or - and change font color
        if (tdVal-tdValRange+3)==1 or (tdVal-tdValRange+3)==3:
            RangeString_Shift = ('%s%d' % (XlsAlphabetArr[(tdVal-tdValRange+2)],trVal))
            if '-' in TEXTVAL:
                Range(RangeString).fontcolor =(0,176,80)
                Range(RangeString_Shift).fontcolor =(0,176,80)
            else:
                if TEXTVAL != '0':
                    Range(RangeString).fontcolor =(255,0,0)
                    Range(RangeString_Shift).fontcolor =(255,0,0)
        
        
        time.sleep(0.1)
        
TimeString = driver.find_element_by_xpath('/html/body/center/table[3]/tbody/tr/td[3]/table[2]/tbody/tr/td[2]/table/tbody/tr/td[1]').text
CurrencyString = driver.find_element_by_xpath('/html/body/center/table[3]/tbody/tr/td[3]/table[4]/tbody/tr[1]/td[3]/table/tbody/tr[1]/td[1]').text
print TimeString, CurrencyString
