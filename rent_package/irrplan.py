# -*- coding: utf-8 -*-
import xlrd
import os
import json

def read_excel():
    #open file
    curdict=os.getcwd()
    print(curdict)
    file_path=curdict+r'\rent_package.xlsx'
    workbook=xlrd.open_workbook(file_path)
    sheet_name=workbook.sheet_names()[8]
    print("回报计划")
    print(sheet_name)
    sheet=workbook.sheet_by_index(8)
    print(sheet)
    excArray=[]
    curRID=0
    print(sheet.name,sheet.nrows,sheet.ncols)

    rowsNum=sheet.nrows
    colsNum=sheet.ncols

    formName = 'irr'
    ignoreRowNum=0
    for index in range(rowsNum-ignoreRowNum):
        i=index+ignoreRowNum
        if(sheet.cell(i,0).ctype==0):
            continue
        else:
            excArray.append({})
            rowValues=[]

            if(curRID==0):
                #这里预留做头部逻辑
                for j in range(colsNum):
                    rowValues.append({})
                    if(sheet.cell(i,j).ctype==0):
                        rowValues[j]['value']=''
                    else:
                        rowValues[j]['rid']= formName+"_col"+str(j)+'_row'+str(curRID)
                        rowValues[j]['value']= sheet.cell(i,j).value
            else:
                for j in range(colsNum):
                    rowValues.append({})
                    if(sheet.cell(i,j).ctype==0):
                        rowValues[j]['value']=''
                    else:
                        rowValues[j]['rid']= formName+"_col"+str(j)+'_row'+str(curRID)
                        rowValues[j]['value']= sheet.cell(i,j).value

            excArray[curRID]['values']=rowValues
            curRID+=1
    return excArray


def storeJson(obj):
    #with codecs.open('arrearage.json','w','utf-8') as f:
    with open('irrplan.json', 'w') as f:
        f.write(json.dumps(obj, indent=4))

if __name__ == "__main__":
    income = read_excel()
    print(income)
    storeJson(income)




