#%%EXCEL
import csv

# 開啟輸出的 CSV 檔案
with open('read.csv', newline='') as csvfile:
  # 建立 CSV 檔寫入器
  #writer = csv.writer(csvfile)
  reader=csv.reader(csvfile)
  li=list(reader)
  MGRNOread=li[0]
  MGRNAMEread=li[1]
  
#%% 名字處理

MGRNOsearch=[]
MGRNAMEsearch=[]



for NAME,NO in zip(MGRNAMEread,MGRNOread):
    
    N=NAME.split(' ')
    NAME=''
    count=0
    for n in N:
        if len(n)==1:   
            NAME+=n
        else:
            NAME+=' '+n
            if count==1:
                break
            count=1
    
    if len(MGRNAMEsearch)>0:
        if NAME==MGRNAMEsearch[-1]:
            continue
    MGRNAMEsearch.append(NAME)
    MGRNOsearch.append(NO)
    print(NAME)







#%% GET CIK


headers = {
'Host': 'www.sec.gov',
'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64; rv:66.0) Gecko/20100101 Firefox/66.0',
'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
'Accept-Language': 'zh-TW,zh;q=0.8,en-US;q=0.5,en;q=0.3',
'Accept-Encoding': 'gzip, deflate, br',
'Referer': 'https://www.sec.gov/edgar/searchedgar/cik.htm?fbclid=IwAR3PklbVKs_q431XjrBtfahYvVAOY2LeQPv3aajiosJZToV1Xv8VjMipHr8',
'Content-Type': 'application/x-www-form-urlencoded',
'Content-Length': '10',
'Connection': 'keep-alive',
'Cookie': '_ga=GA1.2.478561897.1555077471; _4c_=jVLbjtMwEP0V5OeSxrdLKiFUhIDlIgqCssvLyrHHm7Tppjhpsqjqv%2B%2B4qfrCC5GieM6cORPPmSMZK3gkCyql0iaX%2BCoxI1v425HFkcTap89AFoSrIFyQZWAhF057nxeeu2CNsormwMmMPCUdxYwUTEjG89OMeNSe6j0Ee2j6K03kTOeCayqRVl9Z%2F%2BQFx7zbXwhH4loPSKRFJjOG7ENsMKz6ft8t5vNxHLMOXPbQDnPwDzbOO7DRVdPZ1dus6nevQ%2Bma2r%2B6GZff%2BWrblOtP3f0fwentJr7pg63uhvXy6x37DN9WA7d2U7fdx98%2F2jW9Hcx686Xef4gGW5exHTuI2P5dHSG0Ty%2BUQjjgz5KcMsapo6y0GpwsrOOWFqUC8FDAmdfiiMmv%2BtGjCoaoADGe5f7vNrtDf7DNhKRrNSjS1X2azqXoAqCbE%2FZywnqIu9Qbj%2Fs09DTGpnW2SaW4DTPyfnn%2F8%2BYtRkIbqagpdIYbInOthaaYT7ZeHUV%2FzpZJZZQqNFVGY48ebTHoYnpOp9Mz; _gid=GA1.2.601269144.1556779739; _gat_UA-30394047-1=1; _gat=1; _gat_GSA_ENOR0=1; _gat_GSA_ENOR1=1',
'Upgrade-Insecure-Requests': '1'}
URL='https://www.sec.gov/cgi-bin/cik_lookup'

import requests
from bs4 import BeautifulSoup
import time

session_requests = requests.session()

CIKGet=[]
nameGet=[]
MGRNO=[]
MGRNAME=[]



#%%
for NAME,NO,x in zip(MGRNAMEsearch,MGRNOsearch,range(5212,len(MGRNOsearch))):
    
    tStart = time.time()#計時開始
    print('Searching CIK ...'+str(x))
    
    NAME=NAME.strip()
    payload = {
        'company':NAME
    }

    
    r = requests.post(URL, data = payload, headers = headers) # search
    while(r.status_code != 200):
        time.sleep(2)
        print('No Response !')
        r = requests.post(URL, data = payload, headers = headers)

    
    soup = BeautifulSoup(r.text,"html.parser")

    N= soup.select("pre a")
    print('Get ...'+str(len(N)))
    for n in N:
        CIKGet.append(n.string)
        name= n.next_sibling #soup.select("pre")[1].contents[1]
        nameGet.append(name.replace("\n","").strip())
        MGRNO.append(NO)
        MGRNAME.append(NAME)
        
    tEnd = time.time()#計時結束
    print ("It cost %f sec" % (tEnd - tStart))#列印結果

    
#%% Get DocumentURL
import requests
from bs4 import BeautifulSoup

test = open("test.txt","w",encoding='UTF-8')

aCIK=[]
aNAME=[]
aMGRNO=[]
aMGRNAME=[]

Date=[]
DocumentURL=[]
OverPage=[]

for cik,xn in zip(CIKGet, range(len(CIKGet))):
    
    if cik not in aCIK:
        
        tStart = time.time()#計時開始
        print('Get DocumentURL ...'+str(xn))
        
        r = requests.get("https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK="+cik+"&type=N-PX&dateb=&count=100&scd=filings") #將網頁資料GET下來
        while(r.status_code != 200):
            time.sleep(2)
            print('No Response !')
            r = requests.get("https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK="+cik+"&type=N-PX&dateb=&count=100&scd=filings") #將網頁資料GET下來
    
        soup = BeautifulSoup(r.text,"html.parser") #將網頁資料以html.parser 
        sel= soup.select("table.tableFile2 tr") 
        
        #Name
        if len(soup.select('div.companyInfo span.companyName')) != 0:
            name=soup.select('div.companyInfo span.companyName')[0]
            n=name.text
            n=n.split('CIK#')[0]
                  
        #document
        for s,x in zip(sel, range(len(sel))):
            print("第"+str(x+1)+"個檔案:")
            if x >= 100:
                OverPage.append("https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK="+cik+"&type=N-PX&dateb=&count=100&scd=filings")
            y= s.select("td:nth-of-type(4)")
            y1=s.select("#documentsbutton")
            if y != []:
                z= y[0].string
                aCIK.append(cik)
                aNAME.append(n)
                aMGRNO.append(MGRNO[xn])
                aMGRNAME.append(MGRNAME[xn])
                Date.append(z)
                DocumentURL.append(y1[0]["href"])
                print(z)
                print(y1[0]["href"])
                
        tEnd = time.time()#計時結束
        print ("It cost %f sec" % (tEnd - tStart))#列印結果


    else:
        print('Skip DocumentURL ...'+str(xn))

#%%
aS1=[]       
aS2=[]
aCITY=[]     
aSTATE=[]
aZIP=[]
aSeriesName=[]
aClassContract=[]
MAILADDRESS=['']*len(DocumentURL)

#%% Get  ADDRESS
for s,x in zip(DocumentURL, range(len(DocumentURL))):
    
    tStart = time.time()#計時開始
    print('Get Address ...'+str(x))
    
    
    r = requests.get("https://www.sec.gov"+s) #將網頁資料GET下來
    while(r.status_code != 200):
        time.sleep(2)
        print('No Response !')
        r = requests.get("https://www.sec.gov"+s) #將網頁資料GET下來

    soup = BeautifulSoup(r.text,"html.parser") #將網頁資料以html.parser 
    

    seriesName= soup.select("table.tableSeries tr td.seriesName a")
    classContract=soup.select("table.tableSeries tr.contractRow td.classContract a")
    
    SeriesName=''
    for sn in seriesName:
        SeriesName=SeriesName+sn.string+'\r\n'
    ClassContract=''
    for cc in classContract:
        ClassContract=ClassContract+cc.string+'\r\n'
    
    sel= soup.select("table.tableFile td a")  
    r=requests.get("https://www.sec.gov"+sel[-1]["href"])
    tex = r.text.splitlines()
    
    if '\tBUSINESS ADDRESS:\t' in tex:
        start=tex.index('\tBUSINESS ADDRESS:\t')
    else:
        start=tex.index('\tMAIL ADDRESS:\t')
        MAILADDRESS[x]='1'
    
    S1=tex[start+1].split('\t')
    S2=tex[start+2].split('\t')
    CITY=tex[start+3].split('\t')
    STATE=tex[start+4].split('\t')
    ZIP=tex[start+5].split('\t')
    
    if len(CITY[-1]) == 2:
        aS1.append(S1[-1])
        aS2.append('')
        aCITY.append(S2[-1])
        aSTATE.append(CITY[-1])
        aZIP.append(STATE[-1])
    else:
        aS1.append(S1[-1]) 
        aS2.append(S2[-1]) 
        aCITY.append(CITY[-1])  
        aSTATE.append(STATE[-1]) 
        aZIP.append(ZIP[-1])
    
    aSeriesName.append(SeriesName)
    aClassContract.append(ClassContract)
    
    tEnd = time.time()#計時結束
    print ("It cost %f sec" % (tEnd - tStart))#列印結果
    

#

with open('inst address.csv','w', newline='') as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(['MGRNO','Inst. Name','CIK','Date','Street 1','Street 2','CITY','State','ZIP','DocumentURL','SeriesName','ClassContract','Mail Address'])
    
    for n in range(len(aNAME)):
        writer.writerow([aMGRNO[n],aNAME[n],aCIK[n],Date[n],aS1[n],aS2[n],aCITY[n],aSTATE[n],aZIP[n],"https://www.sec.gov"+DocumentURL[n],aSeriesName[n],aClassContract[n],MAILADDRESS[n]])

    csvfile.close()


