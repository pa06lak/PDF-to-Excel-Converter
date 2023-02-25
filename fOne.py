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
    #print(res)
    for file in res: 
        if file[0] != "L": 
            res.remove(file)
    return (res)


def getInfo(res):
    csvList = []
    for f in res:
        newlist = []
        
        LAN = f.replace("LAN", "")
        LAN = LAN.replace(".pdf", "")
        print(LAN, "this is the LAN")
        newlist.append(LAN)
        
        csv_fileName = f[0:9] + ".csv"
        fileName = "newFolder3/" + str(f)
        print(fileName)

        tables = camelot.read_pdf(fileName,flavour = "lattice")
        store = str(tables[0])
        store = int(store[18]) - 1
        
        first_table = tables[0]
        df = first_table.df
        
        name1 = df.iat[2,4]
        name2 = df.iat[2,5]
        if len(name1) != 0: 
            name = name1
        else:
            name = name2
        sep = '\n'
        stripped = name.split(sep, 1)[0]
        stripped = stripped.title()
        newlist.append(stripped)
        print(stripped)
        
        DOB = df.iat[2,store]
        newlist.append(DOB)
        print(DOB)
        
        dateOfReferral = df.iat[2,0]
        newlist.append(dateOfReferral)
        print(dateOfReferral)
        
        practiceAddress = df.iat[13,0] + " " + df.iat[13,4] + " " + df.iat[13,2]
        print(practiceAddress)
        newlist.append(practiceAddress)
        
        telephone = ""
        t1 = df.iat[6,5]
        t2 = df.iat[6,4]
        if t2 != "": 
            telephone = t2
            sep = "\n"
            stripped = telephone.split(sep,1)[0]
            telephone = telephone[0:11]
            print(telephone)
        elif t1 != "": 
            telephone = t1
            sep = "\n"
            stripped = telephone.split(sep,1)[0]
            stripped = stripped.replace(" ", "")
            telephone = telephone[0:11]
            print(telephone)
        else: 
            telephone = "n/a"
            print(telephone)
        
        newlist.append(telephone)
            
        csvList.append(newlist)
    return csvList

def makeTable(csvList): 
    #type(tables)
    #tables
    #first_table = tables[0]
    workbook = openpyxl.load_workbook('newExcel.xlsx')
    sheet = workbook["pyexcel_sheet1"]
    max_row = sheet.max_row
    for row in csvList: 
        sheet.append(row)

    workbook.save("newExcel.xlsx")

res = getRes()
csvList = getInfo(res)
makeTable(csvList)

if __name__ == "__main__":
    print('Hello World!')
    
