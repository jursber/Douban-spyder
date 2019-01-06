#-*- coding: UTF-8 -*-

import time
import urllib
import random
import requests
from bs4 import BeautifulSoup
import xlwt
from openpyxl import Workbook

#Some User Agents
hds=[{'User-Agent':'Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6'},\
{'User-Agent':'Mozilla/5.0 (Windows NT 6.2) AppleWebKit/535.11 (KHTML, like Gecko) Chrome/17.0.963.12 Safari/535.11'},\
{'User-Agent': 'Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.2; Trident/6.0)'}]


def book_spider(book_tag):
    page_num = 0
    data_list = []
    try_times = 0
    timeout=0
    while (1):
        try:
            url = 'http://www.douban.com/tag/' + urllib.parse.quote(book_tag) + '/book?start=' + str(page_num)
            req = requests.get(url, headers=hds[random.randint(0,2)])
        except:
            timeout+=1
            if timeout>5:
                print('连接失败')
                break
            else:
                continue
        #开始爬取
        try_times+=1
        soup = BeautifulSoup(req.text, 'html.parser')
        
        #停止运行条件
        if try_times>200 and len(soup.select('div.book-list > dl > dd > a'))==0:
            break
        
        name_tem = [each.get_text() for each in soup.select('div.book-list > dl > dd > a')]
        info_tem = [each.get_text() for each in soup.select('div.book-list > dl > dd > div.desc')]
        url_tem=[each.get('href') for each in soup.select('div.book-list > dl > dd > a')]
        rating_tem=[each.get_text() for each in soup.select('div.book-list > dl > dd > div.rating > span.rating_nums')]
        
        for name,info,url,rate in zip(name_tem,info_tem,url_tem,rating_tem):
            data={'book_name':name,'book_info':info,'book_url':url,'rating':rate}
            data_list.append(data)
        
        page_num += 15
        time.sleep(random.random())

        if page_num > 15:
            break
    print(data_list) 
    return data_list


def save_to_xlsx(book_info_lists):
    heads=['book_name','book_info','book_url','rating']
    workbook=xlwt.Workbook()
    for record in book_info_lists:
        worksheet=workbook.add_sheet(record)
        i=0
        for head in heads:
            worksheet.write(0,i,head)
            i+=1

        
        j=1
        for data in book_info_lists[record]:
            worksheet.write(j, 0,data['book_name'])
            worksheet.write(j, 1, data['book_info'])
            worksheet.write(j, 2, data['book_url'])
            worksheet.write(j, 3, data['rating']) 
            j+=1
    workbook.save('test.xls')
    return

list1=book_spider('名著')
dic={'名著':list1}

save_to_xlsx(dic)