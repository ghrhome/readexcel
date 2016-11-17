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
    sheet_name=workbook.sheet_names()[0]
    sheet=workbook.sheet_by_index(0)
    print(sheet)
    incomeArray=[]
    curRID=0
    print(sheet.name,sheet.nrows,sheet.ncols)

    rowsNum=sheet.nrows
    colsNum=sheet.ncols

    formName='income'
    ignoreRowNum=3
    for i in range(rowsNum-ignoreRowNum):
        incomeArray.append({})
        rowValues=[]
        for j in range(colsNum):
          # print(sheet.cell(i,j).value.encode('utf-8'));
            #0- empty
            rowValues.append({})
            if j==0:
                incomeArray[i]['name']= sheet.cell(i,j).value
                rowValues[j]['rid']= formName+"_col"+str(j)+'_row'+str(curRID)
                rowValues[j]['value']= sheet.cell(i,j).value
            else:
                rowValues[j]['rid']= formName+"_col"+str(j)+'_row'+str(curRID)
                rowValues[j]['value']= sheet.cell(i,j).value

        incomeArray[i]['values']=rowValues

        curRID+=1
    return incomeArray


def storeJson(obj):
    #with codecs.open('arrearage.json','w','utf-8') as f:
    with open('income.json','w') as f:
        f.write(json.dumps(obj, indent=4))

if __name__ == "__main__":
    income=read_excel()
    print(income)
    storeJson(income)




