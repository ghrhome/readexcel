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
    sheet_name=workbook.sheet_names()[4]
    print("sheetName")
    print(sheet_name)
    sheet=workbook.sheet_by_index(4)
    print(sheet)
    excArray=[]
    curRID=0
    print(sheet.name,sheet.nrows,sheet.ncols)

    rowsNum=sheet.nrows
    colsNum=sheet.ncols

    formName = 'chart'
    ignoreRowNum=0
    for index in range(rowsNum-ignoreRowNum):
        i=index+ignoreRowNum
        if(sheet.cell(i,0).ctype==0):
            continue
        else:
            excArray.append({})
            rowValues=[]
            if(curRID==0):
                excArray[curRID]['name']='year'
                #excArray[curRID]['lables']=[]
                for j in range(colsNum-1):
                    #rowValues.append({})
                   rowValues.append(sheet.cell(i,(j+1)).value)
            else:
                for j in range(colsNum):
                   # rowValues.append({})
                    if j==0:
                        excArray[curRID]['name']= sheet.cell(i,0).value
                    else:
                        if(sheet.cell(i,j).ctype==0):
                            rowValues.append('')
                        else:
                            rowValues.append(sheet.cell(i,j).value)
            excArray[curRID]['values']=rowValues
            curRID+=1
    return excArray


def storeJson(obj):
    #with codecs.open('arrearage.json','w','utf-8') as f:
    with open('chart.json', 'w') as f:
        f.write(json.dumps(obj, indent=4))

if __name__ == "__main__":
    income = read_excel()
    print(income)
    storeJson(income)




