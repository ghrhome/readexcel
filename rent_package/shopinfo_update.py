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
            #rowValues=[]

            if(curRID==0):
                for j in range(colsNum):
                    #rowValues.append({})
                    if j==0:
                        excArray[curRID]['shopIndex']= sheet.cell(i,j).value
                    elif j==1:
                        excArray[curRID]["floor"]= sheet.cell(i,j).value
                    elif j==2:
                        excArray[curRID]["floorArea"]= sheet.cell(i,j).value
                    elif j==3:
                        excArray[curRID]["indoorArea"]= sheet.cell(i,j).value
                    elif j==4:
                        excArray[curRID]["form"]= sheet.cell(i,j).value
                    elif j==5:
                        excArray[curRID]["property"]= sheet.cell(i,j).value
                    else :
                        excArray[curRID]["rentStandard"]= sheet.cell(i,j).value

            else:
                for j in range(colsNum):
                    if j==0:
                        excArray[curRID]['shopIndex']= sheet.cell(i,j).value
                    elif j==1:
                        excArray[curRID]["floor"]= sheet.cell(i,j).value
                    elif j==2:
                        excArray[curRID]["floorArea"]= sheet.cell(i,j).value
                    elif j==3:
                        excArray[curRID]["indoorArea"]= sheet.cell(i,j).value
                    elif j==4:
                        excArray[curRID]["form"]= sheet.cell(i,j).value
                    elif j==5:
                        excArray[curRID]["property"]= sheet.cell(i,j).value
                    else :
                        excArray[curRID]["rentStandard"]= sheet.cell(i,j).value
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




