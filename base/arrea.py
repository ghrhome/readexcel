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
    workbook=xlrd.open_workbook(r'e:\readexcel\arrea.xlsx')
    sheet_name=workbook.sheet_names()[1]
    sheet=workbook.sheet_by_index(1)
    print(sheet)
    shops=[]
    curShop=''
    curSID=0
    print(sheet.name,sheet.nrows,sheet.ncols)

    rowsNum=sheet.nrows
    colsNum=sheet.ncols
    for i in range(rowsNum):

        shops.append({})
        #shops[i]={}
        for j in range(colsNum):
          # print(sheet.cell(i,j).value.encode('utf-8'));
            #0- empty
            if j==0:
                if(sheet.cell(i,j).value==curShop):
                    shops[i]["sid"]="s"+str(curSID)
                    #shops[i]["name"]=sheet.cell(i,j).value.encode("utf-8")
                    shops[i]["name"]=sheet.cell(i,j).value
                    #shops[i]["id"]="id"+str(i)
                else:
                    curSID+=1
                    #curShop=shops[i]["name"]=sheet.cell(i,j).value.encode("utf-8")
                    curShop=shops[i]["name"]=sheet.cell(i,j).value
                    shops[i]["sid"]="s"+str(curSID)
                    shops[i]["name"]=curShop
                    #shops[i]["id"]="id"+str(i)
            elif j==1:
                  if(sheet.cell(i,j).ctype==0):
                        shops[i]["rentType"]==""
                  else:
                        shops[i]['rentType']=sheet.cell(i,j).value

            elif j==2:
                 if(sheet.cell(i,j).ctype==0):
                    arrearage=0
                    shops[i]["arrearage"]=arrearage
                 else:
                    arrearage=sheet.cell(i,j).value
                    shops[i]["arrearage"]=arrearage
    return shops


def storeJson(obj):
    #with codecs.open('arrearage.json','w','utf-8') as f:
    with open('all.json','w') as f:
        f.write(json.dumps(obj,indent=4))



if __name__ == "__main__":
    arrearage=read_excel()
    print(arrearage)
    storeJson(arrearage)




