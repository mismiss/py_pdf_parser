# -*- coding: utf-8 -*-

import requests
from urllib import parse
from requests.cookies import RequestsCookieJar
from lxml import etree
import json
import os
import time

import xlrd

# import sys
# reload(sys) 
# sys.setdefaultencoding('utf-8')

baseUrl = "http://www.cninfo.com.cn"

keyWordsList = ["深康佳A"]
#http://www.cninfo.com.cn/new/fulltextSearch/full?searchkey=平安银行&sdate=&edate=&isfulltext=false&sortName=nothing&sortType=desc&pageNum=1
#http://www.cninfo.com.cn/new/fulltextSearch/full?searchkey=%E5%B9%B3%E5%AE%89%E9%93%B6%E8%A1%8C&sdate=&edate=&isfulltext=false&sortName=nothing&sortType=desc&pageNum=1
#http://www.cninfo.com.cn/new/fulltextSearch/full?searchkey=%C6%BD%B0%B2%D2%F8%D0%D0&sdate=&edate=&isfulltext=false&sortName=nothing&sortType=desc&pageNum=1
headers = {
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3",
    "Accept-Encoding": "gzip, deflate",
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Cache-Control": "max-age=0",
    "Connection": "keep-alive",
    "Cookie": "noticeTabClicks=%7B%22szse%22%3A1%2C%22sse%22%3A0%2C%22hot%22%3A0%2C%22myNotice%22%3A0%7D; tradeTabClicks=%7B%22financing%20%22%3A0%2C%22restricted%20%22%3A0%2C%22blocktrade%22%3A0%2C%22myMarket%22%3A0%2C%22financing%22%3Anull%7D; JSESSIONID=FFF38C0A86FAB304DAE427ED882AA20F; _sp_ses.2141=*; UC-JSESSIONID=B837D3E7FD971AD71EE7BEFF5144D4A5; insert_cookie=37836164; cninfo_search_record_cookie=%E5%B9%B3%E5%AE%89%E9%93%B6%E8%A1%8C; cninfo_user_browse=000929,gssz0000929,%E5%85%B0%E5%B7%9E%E9%BB%84%E6%B2%B3|000002,gssz0000002,%E4%B8%87%20%20%E7%A7%91%EF%BC%A1; _sp_id.2141=bb90964c-5a5d-4a48-b5c9-a2f56648b50e.1573199366.1.1573199638.1573199366.2aa81b98-4c5f-45e9-972c-ab3bfd9af710",
    "Host": "www.cninfo.com.cn",
    "Referer": "www.cninfo.com.cn",
    "Upgrade-Insecure-Requests": "1",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.25 Safari/537.36 Core/1.70.3626.400 QQBrowser/10.4.3211.400",
}

cookie_jar = RequestsCookieJar()
# cookie_jar.set("4bd54_ol_offset","166646")
# cookie_jar.set("4bd54_ipstate","1550729151")
# cookie_jar.set("4bd54_readlog","%2C1944825%2C")
# cookie_jar.set("4bd54_c_stamp","1550729821")
# cookie_jar.set("4bd54_lastpos","F6")
# cookie_jar.set("4bd54_lastvisit","674%091550729821%09%2Fbbs%2Fthread.php%3Ffid6")
# cookie_jar.set("sc_is_visitor_unique","rx4629288.1550729958.D340E62FE57B4FF097095612EED92D8C.1.1.1.1.1.1.1.1.1")

def requestUrl(pageurl):
    r = requests.get(pageurl,headers = headers,cookies = cookie_jar)
    if(len(r.text) < 21):
        time.sleep(1)
        # setCookies(r)
        r = requests.get(pageurl,headers = headers,cookies = cookie_jar)
        # print(r.text.encode('unicode_escape').decode('string_escape'))
    # return etree.HTML(r.text)
    return r

def requestPost(url,formData,headerForm):
    r = requests.post(url,data= formData, headers = headerForm,cookies = cookie_jar)
    return r

def requestOrgId(pageUrl,companinyName,companyType = 'SZ'):
    print(pageUrl)
    jsonResult = requestUrl(pageUrl).text
    jsonData = json.loads(jsonResult)

    if jsonData["announcements"] and len(jsonData["announcements"]) > 0:
        
        if companyType.find('SZ') > -1 :
            columnCode = "szse"
            plateCode = 'sz'
        else:
            columnCode = "sse"
            plateCode = 'sh' 

        j = jsonData["announcements"][0]
        orgetId = j["orgId"]
        secCode = j["secCode"]
        postData = {
            "pageNum": "1",
            "pageSize": "30",
            "tabName": "fulltext",
            "column": columnCode,
            "stock": secCode + "," + orgetId,
            "plate": plateCode,
            "category": "category_ndbg_szsh;",
            "seDate": "2007-01-01 ~ 2019-05-19",
        }
        headerForm = {
            "Host":"www.cninfo.com.cn",
            "Content-Length":"166",
            "Accept":"application/json, text/javascript, */*; q=0.01",
            "Origin":"http://www.cninfo.com.cn",
            "X-Requested-With":"XMLHttpRequest",
            "User-Agent":"Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36",
            "Content-Type":"application/x-www-form-urlencoded; charset=UTF-8",
            "Referer":"http://www.cninfo.com.cn/new/disclosure/stock?orgId="+orgetId+"&stockCode=" + secCode,
            "Accept-Encoding":"gzip, deflate",
            "Accept-Language":"zh-CN,zh;q=0.9",
            "Cookie":"noticeTabClicks=%7B%22szse%22%3A1%2C%22sse%22%3A0%2C%22hot%22%3A0%2C%22myNotice%22%3A0%7D; tradeTabClicks=%7B%22financing%20%22%3A0%2C%22restricted%20%22%3A0%2C%22blocktrade%22%3A0%2C%22myMarket%22%3A0%2C%22financing%22%3Anull%7D; JSESSIONID=2E754292617CB2D6599D867C76FC4072; insert_cookie=37836164; cninfo_search_record_cookie=%E5%B9%B3%E5%AE%89%E9%93%B6%E8%A1%8C; _sp_ses.2141=*; UC-JSESSIONID=B06237E0A7911FC8C717F4FA1789EE22; cninfo_user_browse=000001,gssz0000001,%E5%B9%B3%E5%AE%89%E9%93%B6%E8%A1%8C|000989,gssz0000989,%E4%B9%9D%20%E8%8A%9D%20%E5%A0%82|000929,gssz0000929,%E5%85%B0%E5%B7%9E%E9%BB%84%E6%B2%B3|000002,gssz0000002,%E4%B8%87%20%20%E7%A7%91%EF%BC%A1; _sp_id.2141=bb90964c-5a5d-4a48-b5c9-a2f56648b50e.1573199366.2.1573201813.1573199638.24385f78-d923-48fa-bbc9-9f403d0cf662",
            "Connection":"keep-alive",
        }
        #请求财报列表
        quaryPostUrl = "http://www.cninfo.com.cn/new/hisAnnouncement/query"
        jsonResult = requestPost(quaryPostUrl,postData,headerForm).text
        jsonData = json.loads(jsonResult)
        if jsonData["announcements"] and len(jsonData["announcements"]) > 0:
            headerGet = {
                "Host":"www.cninfo.com.cn",
                "Upgrade-Insecure-Requests":"1",
                "User-Agent":"Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.103 Safari/537.36",
                "Accept":"text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3",
                "Accept-Encoding":"gzip, deflate",
                "Accept-Language":"zh-CN,zh;q=0.9",
                "Cookie":"noticeTabClicks=%7B%22szse%22%3A1%2C%22sse%22%3A0%2C%22hot%22%3A0%2C%22myNotice%22%3A0%7D; tradeTabClicks=%7B%22financing%20%22%3A0%2C%22restricted%20%22%3A0%2C%22blocktrade%22%3A0%2C%22myMarket%22%3A0%2C%22financing%22%3Anull%7D; JSESSIONID=FFDAD48B8E19962E0FCA978FC6D68F46; insert_cookie=37836164; cninfo_search_record_cookie=%E5%B9%B3%E5%AE%89%E9%93%B6%E8%A1%8C; _sp_ses.2141=*; UC-JSESSIONID=B06237E0A7911FC8C717F4FA1789EE22; cninfo_user_browse=000001,gssz0000001,%E5%B9%B3%E5%AE%89%E9%93%B6%E8%A1%8C|000989,gssz0000989,%E4%B9%9D%20%E8%8A%9D%20%E5%A0%82|000929,gssz0000929,%E5%85%B0%E5%B7%9E%E9%BB%84%E6%B2%B3|000002,gssz0000002,%E4%B8%87%20%20%E7%A7%91%EF%BC%A1; _sp_id.2141=bb90964c-5a5d-4a48-b5c9-a2f56648b50e.1573199366.2.1573202408.1573199638.24385f78-d923-48fa-bbc9-9f403d0cf662",
                "Connection":"keep-alive",
            }
            for jsonItem in jsonData["announcements"]:
                # try:
                #     folderName = "pdfdownload/" + companinyName
                #     # fileNameGBK = jsonItem["secName"].encode('gbk')
                #     if jsonItem["announcementTitle"].find(u"\u6458\u8981") > -1 :
                #         folderName = folderName + "/摘要"
                #     if not os.path.exists(folderName):
                #         os.makedirs(folderName)
                    
                #     title = "/" + jsonItem["secName"] + "_"+ jsonItem["announcementTitle"]+".pdf"
                #     announceId = jsonItem["announcementId"]
                #     downloadPdfUrl = "http://www.cninfo.com.cn/new/announcement/download?bulletinId=" + announceId
                #     print(downloadPdfUrl,jsonItem["announcementTitle"].encode('utf-8'))
                #     r = requests.get(downloadPdfUrl,headers = headerGet)
                    
                #     fileName = folderName + title

                #     with open(fileName,'wb') as img:
                #         img.write(r.content)
                # except Exception as e:
                #     print ("download 1 item error=>>>>>>>>>",title)
                try:
                    downloadPDF(jsonItem,companinyName,headerGet)
                except Exception as e1:
                    try:
                        print ("download 1 item error=>>>>>>>>>,retring,remian 0 times",companinyName,e1)
                        time.sleep(2)
                        downloadPDF(jsonItem,companinyName,headerGet)
                    except Exception as e2:
                        print ("download 1 item error=>>>>>>>>>,skip the file!!!",companinyName,e2)

                time.sleep(0.5)
            time.sleep(1)

        # http://www.cninfo.com.cn/new/announcement/download?bulletinId=1205937912&announcementTime=2019-03-26

def downloadPDF(jsonItem,companinyName,headerGet):
    folderName = "pdfdownload/" + companinyName
    folderName = folderName.replace('*','')
    # fileNameGBK = jsonItem["secName"].encode('gbk')
    if jsonItem["announcementTitle"].find(u"\u6458\u8981") > -1 :
        folderName = folderName + "/摘要"
    if not os.path.exists(folderName):
        os.makedirs(folderName)
    
    title = "/" + jsonItem["secName"] + "_"+ jsonItem["announcementTitle"]+".pdf"
    announceId = jsonItem["announcementId"]
    downloadPdfUrl = "http://www.cninfo.com.cn/new/announcement/download?bulletinId=" + announceId
    print(downloadPdfUrl,companinyName,jsonItem["announcementTitle"])
    r = requests.get(downloadPdfUrl,headers = headerGet)
    
    fileName = folderName + title
    fileName =  fileName.replace('*','')
    with open(fileName,'wb') as img:
        img.write(r.content)

def readXlsx():
    workbook = xlrd.open_workbook(r'11.xlsx')
    print(workbook.sheet_names()) # [u'sheet1', u'sheet2']
    sheet1 = workbook.sheet_by_index(0) # sheet索引从0开始
    cols0 = sheet1.col_values(0) # 获取第er列内容
    cols1 = sheet1.col_values(1) # 获取第三列内容
    return cols0 , cols1

def main():
    codeList ,keyWordsList = readXlsx()
    for index,keyWords in enumerate(keyWordsList):
        if len(keyWords) < 1:
            continue
        
        # keyWordsUtf8 = keyWords.encode('utf-8')
        # keyWordsGBK = keyWordsUtf8.encode('gbk')
        print(keyWords)

        keyWordsSafe = keyWords.replace('*','')

        if not os.path.exists("pdfdownload/" + keyWordsSafe):
            os.makedirs("pdfdownload/" +keyWordsSafe)

        # print repr(keyWords)
        # print type(keyWords)
        # print(type(keyWords))
        encodeurlparam = parse.quote("pdfdownload/" + keyWords)
        if encodeurlparam == "nothing":
            print("jump 1 item=>>>>>>>>>" + keyWords) 
            continue

        searchOrgidUrl = "http://www.cninfo.com.cn/new/fulltextSearch/full?searchkey=" +  parse.quote(keyWords) +"&sdate=&edate=&isfulltext=false&sortName=nothing&sortType=desc&pageNum=1"
        # searchOrgidUrl = "http://www.cninfo.com.cn/new/fulltextSearch?keyWord=" + quote(keyWords)
        try:
            requestOrgId(searchOrgidUrl,keyWords,codeList[index])
        except Exception as e:
            print("download 1 company error=>>>>>>>>>" + keyWords)
            print(e)

if __name__ == "__main__":
    main()



