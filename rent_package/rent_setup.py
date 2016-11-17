# -*- coding: utf-8 -*-
import xlrd
import os
import json
import codecs
import datetime,calendar

def strtodatetime(datestr,format):
    return datetime.datetime.strptime(datestr,format)
def datediff(beginDate,endDate):
    format="%Y-%m-%d";
    bd=strtodatetime(beginDate,format)
    ed=strtodatetime(endDate,format)
    oneday=datetime.timedelta(days=1)
    count=0
    while bd!=ed:
        ed=ed-oneday
        count+=1
    return count
def datetostr(date):
    return  str(date)[0:10]

def read_excel():
    #open file
    curdict=os.getcwd()
    print(curdict)
    file_path=curdict+r'\rent_package.xlsx'
    workbook=xlrd.open_workbook(file_path)
    sheet_name=workbook.sheet_names()[1]
    print("租金包设定")
    print(sheet_name)
    sheet=workbook.sheet_by_index(1)
    print(sheet)
    excArray=[]
    curRID=0
    print(sheet.name,sheet.nrows,sheet.ncols)

    rowsNum=sheet.nrows
    colsNum=sheet.ncols

    formName = 'rent_setup'
    ignoreRowNum=1
    for index in range(rowsNum-ignoreRowNum):
        i=index+ignoreRowNum
        if(sheet.cell(i,0).ctype==0):
            continue
        else:
            excArray.append({})
            rowValues=[]
            excArray[curRID]['name']= "rent_setup"
            for j in range(colsNum):
                rowValues.append({})
                    #the sub name of the cell is i-1,j

                if sheet.cell(i,j).ctype==3:
                    date=xlrd.xldate_as_tuple(sheet.cell_value(i,j),workbook.datemode)
                    dateStr=str(date[0])+"-"+str(date[1])+"-"+str(date[2])
                    rowValues[j]["value"]= dateStr
                else:
                    rowValues[j]['value']= sheet.cell(i,j).value

                rowValues[j]['name']= sheet.cell((i-1),j).value

            excArray[curRID]['values']=rowValues
            curRID+=1
    return excArray


def storeJson(obj):
    #with codecs.open('arrearage.json','w','utf-8') as f:
    with open('rent_setup.json', 'w') as f:
        f.write(json.dumps(obj, indent=4))

if __name__ == "__main__":
    income = read_excel()
    print(income)
    storeJson(income)




