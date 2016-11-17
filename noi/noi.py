# -*- coding: utf-8 -*-
import xlrd
import os
import json

def read_excel():
    #open file
    curdict=os.getcwd()
    print(curdict)
    file_path=curdict+r'\NOI.xlsx'
    workbook=xlrd.open_workbook(file_path)
    sheet_name=workbook.sheet_names()[3]
    print("sheetName")
    print(sheet_name)
    sheet=workbook.sheet_by_index(3)
    print(sheet)
    excArray=[]
    curRID=0
    print(sheet.name,sheet.nrows,sheet.ncols)

    rowsNum=sheet.nrows
    colsNum=sheet.ncols

    formName = 'noi'
    ignoreRowNum=2
    for index in range(rowsNum-ignoreRowNum):
        i=index+ignoreRowNum
        if(sheet.cell(i,0).ctype==0):
            continue
        else:
            excArray.append({})
            rowValues=[]
            for j in range(colsNum):
                rowValues.append({})
                    #the sub name of the cell is i-1,j
                excArray[curRID]['name']= "noi"
                #rowValues[j]['rid']= formName+"_col"+str(j)+'_row'+str(curRID)
                rowValues[j]['name']= sheet.cell((i-1),j).value
                rowValues[j]['value']= sheet.cell(i,j).value
            excArray[curRID]['values']=rowValues
            curRID+=1
    return excArray


def storeJson(obj):
    #with codecs.open('arrearage.json','w','utf-8') as f:
    with open('noi.json', 'w') as f:
        f.write(json.dumps(obj, indent=4))

if __name__ == "__main__":
    income = read_excel()
    print(income)
    storeJson(income)




