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
    workbook=xlrd.open_workbook(r'e:\readexcel\tall.xlsx')
    sheet_name=workbook.sheet_names()[0]
    sheet=workbook.sheet_by_index(0)
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
                    shops[i]["id"]="id"+str(i)
                else:
                    curSID+=1
                    #curShop=shops[i]["name"]=sheet.cell(i,j).value.encode("utf-8")
                    curShop=shops[i]["name"]=sheet.cell(i,j).value
                    shops[i]["sid"]="s"+str(curSID)
                    shops[i]["name"]=curShop
                    shops[i]["id"]="id"+str(i)
            elif j==1:
                 if(sheet.cell(i,j).ctype==0):
                    loc=""
                    shops[i]["loc"]=loc
                 else:
                    loc=sheet.cell(i,j).value
                    #loc=sheet.cell(i,j).value.encode("utf-8")
                    #print(loc.decode("utf-8"))
                    shops[i]["loc"]=loc
            elif j==2:
                 if(sheet.cell(i,j).ctype==0):
                    arrearageAll=0
                    shops[i]["arrearageAll"]=arrearageAll
                 else:
                    arrearageAll=sheet.cell(i,j).value
                    shops[i]["arrearageAll"]=arrearageAll


            elif j==3:
                 if(sheet.cell(i,j).ctype==0):
                    arrearage=0
                    shops[i]["r-0-30"]=arrearage
                 else:
                    arrearage=sheet.cell(i,j).value
                    shops[i]["r-0-30"]=arrearage

            elif j==4:
                 if(sheet.cell(i,j).ctype==0):
                    arrearage=0
                    shops[i]["r-31-60"]=arrearage
                 else:
                    arrearage=sheet.cell(i,j).value
                    shops[i]["r-31-60"]=arrearage

            elif j==5:
                 if(sheet.cell(i,j).ctype==0):
                    arrearage=0
                    shops[i]["r-61-90"]=arrearage
                 else:
                    arrearage=sheet.cell(i,j).value
                    shops[i]["r-61-90"]=arrearage
            elif j==6:
                 if(sheet.cell(i,j).ctype==0):
                    arrearage=0
                    shops[i]["r-90more"]=arrearage
                 else:
                    arrearage=sheet.cell(i,j).value
                    shops[i]["r-90-more"]=arrearage
    return shops


def storeJson(obj):
    #with codecs.open('arrearage.json','w','utf-8') as f:
    with open('all.json','w') as f:
        f.write(json.dumps(obj,indent=4))



if __name__ == "__main__":
    arrearage=read_excel()
    print(arrearage)
    storeJson(arrearage)




