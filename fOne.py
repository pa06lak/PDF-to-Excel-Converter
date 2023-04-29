import tabula
import pandas as pd
from pathlib import Path
from pprint import pprint
import os
import camelot
import openpyxl
from openpyxl import Workbook
#import pyexcel-xlsx

def getRes():
    dir_path = './newFolder3'
    res = []

    for path in os.listdir(dir_path):
        if os.path.isfile(os.path.join(dir_path, path)):
            res.append(path)
    for file in res: 
        if file[0] != "L": 
            res.remove(file)
    return (res)

def getLANnum(f,LANlist):
    LANnum = f.replace("LAN", "")
    LANnum = LANnum.replace(".pdf", "")
    return LANnum

def getName(dataframe): 
    name1 = dataframe.iat[2,4]
    name2 = dataframe.iat[2,5]
    if len(name1) != 0: 
        name = name1
    else:
        name = name2
    sep = '\n'
    stripped = name.split(sep, 1)[0]
    stripped = stripped.title()
    return stripped

def getDOB(dataframe,store): 
    DOB = dataframe.iat[2,store]
    return DOB

def GetDateOfReferral(dataframe): 
    dateOfReferral = dataframe.iat[2,0]
    return dateOfReferral

def getpracticeAddress(dataframe): 
    practiceAddress = dataframe.iat[13,0] + " " + dataframe.iat[13,4] + " " + dataframe.iat[13,2]
    return practiceAddress

def getTelephone(dataframe): 
    telephone = ""
    t1 = dataframe.iat[6,5]
    t2 = dataframe.iat[6,4]
    if t2 != "": 
        telephone = t2
        sep = "\n"
        stripped = telephone.split(sep,1)[0]
        telephone = telephone[0:12]
    elif t1 != "": 
        telephone = t1
        sep = "\n"
        stripped = telephone.split(sep,1)[0]
        stripped = stripped.replace(" ", "")
        telephone = telephone[0:12]
    else: 
        telephone = "n/a"
    if telephone.lower() == "work": 
        pass
    return telephone

def getInfo(res):
    csvList = []
    for f in res:
        newlist = []

        csv_fileName = f[0:9] + ".csv"
        fileName = "newFolder3/" + str(f)   
        tables = camelot.read_pdf(fileName,flavour = "lattice")
        store = str(tables[0])
        store = int(store[18]) - 1
        
        first_table = tables[0]
        df = first_table.df
        
        LANNumber = getLANnum(f,newlist)
        newlist.append(LANNumber)
        newlist.append(getName(df))
        newlist.append(getDOB(df,store))
        newlist.append(GetDateOfReferral(df))
        newlist.append(getpracticeAddress(df))
        newlist.append(getTelephone(df))
        for y in range(len(newlist)): 
            print(newlist[y])

        
        csvList.append(newlist)
    return csvList

def makeTable(csvList): 
    workbook = openpyxl.load_workbook('newExcel.xlsx')
    sheet = workbook["pyexcel_sheet1"]
    max_row = sheet.max_row
    for row in csvList: 
        sheet.append(row)

    workbook.save("newExcel.xlsx")

#res = getRes()
#csvList = getInfo(res)
#makeTable(csvList)

if __name__ == "__main__":
    res = getRes()
    csvList = getInfo(res)
    makeTable(csvList)
    
