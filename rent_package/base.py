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
    sheet_name=workbook.sheet_names()[0]
    print("店铺信息")
    print(sheet_name)
    sheet=workbook.sheet_by_index(0)
    print(sheet)
    excArray=[]
    curRID=0
    print(sheet.name,sheet.nrows,sheet.ncols)

    rowsNum=sheet.nrows
    colsNum=sheet.ncols

    formName = 'shop'
    ignoreRowNum=0
    for index in range(rowsNum-ignoreRowNum):
        i=index+ignoreRowNum
        if(sheet.cell(i,0).ctype==0):
            continue
        else:
            excArray.append({})
            rowValues=[]

            if(curRID==0):
                for j in range(colsNum):
                    rowValues.append({})
                    if j==0:
                        excArray[curRID]['name']= "shop_info"
                        rowValues[j]['rlabel']="shopIndex"
                    elif j==1:
                        rowValues[j]['rlabel']="floorIndex"
                    elif j==2:
                        rowValues[j]['rlabel']="floorArea"
                    elif j==3:
                        rowValues[j]['rlabel']="indoorArea"
                    elif j==4:
                        rowValues[j]['rlabel']="retailForm"
                    elif j==5:
                        rowValues[j]['rlabel']="property"
                    else :
                        rowValues[j]['rlabel']="rentStandard"
                    rowValues[j]['rid']= formName+"_col"+str(j)+'_row'+str(curRID)
                    rowValues[j]['value']= sheet.cell(i,j).value
            else:
                for j in range(colsNum):
                    rowValues.append({})
                    rowValues[j]['rid']= formName+"_col"+str(j)+'_row'+str(curRID)
                    rowValues[j]['value']= sheet.cell(i,j).value

            excArray[curRID]['values']=rowValues
            curRID+=1
    return excArray


def storeJson(obj):
    #with codecs.open('arrearage.json','w','utf-8') as f:
    with open('shopInfo.json', 'w') as f:
        f.write(json.dumps(obj, indent=4))

if __name__ == "__main__":
    income = read_excel()
    print(income)
    storeJson(income)




