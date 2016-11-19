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
    workbook=xlrd.open_workbook(r'e:\readexcel\test.xlsx')
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

        shops.append({});
        #shops[i]={}
        for j in range(colsNum):
          # print(sheet.cell(i,j).value.encode('utf-8'));
            #0- empty
            if j==0:
                if(sheet.cell(i,j).value==curShop):
                    shops[i]["sid"]="s"+str(curSID)
                    #shops[i]["name"]=sheet.cell(i,j).value.encode("utf-8")
                    #shops[i]["name"]=sheet.cell(i,j).value
                else:
                    curSID+=1
                    #curShop=shops[i]["name"]=sheet.cell(i,j).value.encode("utf-8")
                    curShop=shops[i]["name"]=sheet.cell(i,j).value
                    shops[i]["sid"]="s"+str(curSID)
                    shops[i]["name"]=curShop
            elif j==1:
                if(sheet.cell(i,j).ctype==0):
                    shops[i]["rentType"]==""
                else:
                    #shops[i]['rentType']=sheet.cell(i,j).value.encode("utf-8")
                    shops[i]['rentType']=sheet.cell(i,j).value

            elif j==2:
                 if(sheet.cell(i,j).ctype==0):
                    monthStr=""
                    shops[i]["dutyMonth"]= monthStr
                 elif sheet.cell(i,j).ctype==3:
                    date=xlrd.xldate_as_tuple(sheet.cell_value(i,j),workbook.datemode)
                    monthStr=str(date[0])+"-"+str(date[1])
                    shops[i]["dutyMonth"]= monthStr


            elif j==3:
                 if(sheet.cell(i,j).ctype==0):
                    dateStr=""
                    shops[i]["payDate"]=dateStr
                    shops[i]['outDate']=""
                    shops[i]['rangeType']="0-30"
                 elif sheet.cell(i,j).ctype==3:
                    date=xlrd.xldate_as_tuple(sheet.cell_value(i,j),workbook.datemode)
                    dateStr=str(date[0])+"-"+str(date[1])+"-"+str(date[2])

                    shops[i]["payDate"]=dateStr

                    today=datetostr(datetime.date.today())
                    outDate=datediff(dateStr,today)
                    shops[i]['outDate']=outDate
                    if outDate<=30:
                        shops[i]['rangeType']="0-30"
                    elif outDate>30 and outDate<=60:
                        shops[i]['rangeType']="31-60"
                    elif outDate>60 and outDate<=90:
                        shops[i]['rangeType']="61-90"
                    elif outDate>90:
                        shops[i]['rangeType']="90more"
            elif j==4:
                 if(sheet.cell(i,j).ctype==0):
                    dateRange=""
                    shops[i]["dateRange"]=dateRange

                 else:
                    #dateRange=sheet.cell(i,j).value.encode("utf-8")
                    dateRange=sheet.cell(i,j).value
                    shops[i]["dateRange"]=dateRange


            elif j==5:
                 if(sheet.cell(i,j).ctype==0):
                    bugetIncome=0
                    shops[i]["bugetIncome"]=bugetIncome
                 else:
                    bugetIncome=sheet.cell(i,j).value
                    shops[i]["bugetIncome"]=bugetIncome

            elif j==6:
                 if(sheet.cell(i,j).ctype==0):
                    income=0
                    shops[i]["income"]=income
                 else:
                    income=sheet.cell(i,j).value
                    shops[i]["income"]=income

            elif j==7:
                 if(sheet.cell(i,j).ctype==0):
                    arrearage=0
                    shops[i]["arrearage"]=arrearage
                 else:
                    arrearage=sheet.cell(i,j).value
                    shops[i]["arrearage"]=arrearage

            elif j==8:
                 if(sheet.cell(i,j).ctype==0):
                    arrearageAll=0
                    shops[i]["arrearageAll"]=arrearageAll
                 else:
                    arrearageAll=sheet.cell(i,j).value
                    shops[i]["arrearageAll"]=arrearageAll

            elif j==9:
                 if(sheet.cell(i,j).ctype==0):
                    loc=""
                    shops[i]["loc"]=loc
                 else:
                    loc=sheet.cell(i,j).value
                    #loc=sheet.cell(i,j).value.encode("utf-8")
                    #print(loc.decode("utf-8"))
                    shops[i]["loc"]=loc
    return shops


def storeJson(obj):
    #with codecs.open('arrearage.json','w','utf-8') as f:
    with open('arrearage.json','w') as f:
        f.write(json.dumps(obj,indent=4))



if __name__ == "__main__":
    arrearage=read_excel()
    print(arrearage)
    storeJson(arrearage)




