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
import betterwalk
import os
import subprocess
from nt import chdir
import gc
import random
import math
columns=['Title', 'Cate', 'Keyword','Times']
index = np.arange(30)
#Create train set
Title =['配電盤.xls','變頻器報價檔.xlsx','123配電盤.xlsx','發電機.xls','冰水機報價.xls','冰機.xls']
Cate = ['電器','電器','電器','電器','空調','空調']
KeyWord = ['Cate','配電盤','變頻器','變壓器','冰機','離心式冰機']
df = pd.DataFrame(columns=KeyWord, index = Title)

for i in range(0,len(Title)):
    
    df['Cate'].ix[i] = Cate[i]
    for j in KeyWord[1:]:
        df[j].ix[i] = (random.choice([x for x in range(1,60)]))

#Guess the category
GUESS=[30,99,23,11,2]   

def vector_distance(row):
    #Cosine Similarity
    v1 = map(lambda x:x, row)[1:]
    
    v2 = GUESS

    product = sum([a*b for a,b in zip(v1,v2)])
    len1 = math.sqrt(sum([a*b for a,b in zip(v1,v1)]))
    len2 = math.sqrt(sum([a*b for a,b in zip(v2,v2)]))

    return product / (len1 * len2)
#get vector

#Guess 'GUESS.xls' Cate

df['distance']= df.apply(vector_distance, axis=1)

df = df.sort_values("distance",ascending=True)

#Find max Cate times in k closer

ResultCate = df.ix[0:3]

print ResultCate['Cate'].value_counts().idxmax()
