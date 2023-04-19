# -*- coding:utf-8 -*-
import re
import time
import random
import requests
import pandas as pd
from selenium import webdriver
from bs4 import BeautifulSoup
from openpyxl.styles import PatternFill
from datetime import datetime, timedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.edge.options import Options
from fake_useragent import UserAgent

#導入課程資訊
name=pd.read_excel("../1112航測分組.xlsx")
memo=pd.read_excel('../1112教學日誌.xlsx')
#個人隱私資料：分享前應先銷毀PPI資料文字檔
PPI = open('../PPI.txt', 'r',encoding="utf-8")
PP_Information=[]
for line in PPI.readlines():
    PP_Information.append(line.split('：')[1])
PPI.close
mail = PP_Information[0] #課程平台電子郵件
passw = PP_Information[1] #課程平台密碼

TodayDateAndTime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
#正式用：辨識目前時間
Today = datetime.now().strftime('%Y-%m-%d')
Today = datetime.strptime(Today,'%Y-%m-%d')
Totime = datetime.now().strftime('%H:%M:%S')
Totime = datetime.strptime(Totime,'%H:%M:%S')
#測試用：固定時間
#Today = datetime.strptime('2023-02-23','%Y-%m-%d')
print("開始執行時間", TodayDateAndTime)

# 2023更新：加入週次時間管理(但每年仍須更新)
def today_to_week():
    week=[]
    date=[]
    start_date='2023-02-16'
    dt=datetime.strptime(start_date,'%Y-%m-%d')
    week.append('week0')
    date.append(dt)
    for i in range(1,18):
        week.append('week{}'.format(i))
        date.append((dt+timedelta(days=7*i)).strftime('%Y-%m-%d'))
    week_date=pd.DataFrame({'week':week,'date':date})
    week_date[week_date.date==Today].week
    nextdate=(Today+timedelta(days=1))
    todayweek=week_date.loc[week_date.date<=nextdate, 'week'].iloc[-1]
    return todayweek
today=today_to_week()
dayint=re.findall(r'\d+',today)

delay_choices = [3, 1, 2]  #延遲的秒數
delay = random.choice(delay_choices)  #隨機選取秒數

web_options = Options()
web_options.add_argument("--disable-notifications") #取消網頁中的彈出視窗，避免妨礙網路爬蟲的執行
web_options.add_argument("--headless") # 2023更新：加入"--headless"不顯示視窗

user_agent = UserAgent()
try:
    browser = webdriver.Chrome('./chromedriver.exe',options=web_options)
except:
    browser = webdriver.Edge('./msedgedriver.exe',options=web_options)
browser.implicitly_wait(10) # 等待頁面好 10秒內
url = 'http://gis519.logyuan.idv.tw/dashboard'
browser.get(url) # 前往指定網址
time.sleep(3)

# 2023更新：webdriver指令大改版，詳見https://pythoninoffice.com/fixing-attributeerror-webdriver-object-has-no-attribute-find_element_by_xpath/
#id = browser.find_element_by_class_name("input-block")(過去舊版本寫法，WebDriver4.3.0後不可用)
id = browser.find_element(By.ID, "login-email")
id.send_keys(mail)
#password = browser.find_element_by_name("password")
password = browser.find_element(By.ID, "login-password")
password.send_keys(passw)
button = browser.find_element(By.CLASS_NAME, "action.action-primary.action-update.js-login.login-button").click()
time.sleep(delay)

url = 'http://gis519.logyuan.idv.tw/courses/course-v1:Pccu_Geography+256600+2023_03/course/'
browser.get(url) # 前往指定網址
time.sleep(delay)
url = 'http://gis519.logyuan.idv.tw/courses/course-v1:Pccu_Geography+256600+2023_03/instructor#view-course_info'
browser.get(url) # 前往指定網址
time.sleep(delay)
button = browser.find_element(By.CLASS_NAME,"btn-link.student_admin").click()
time.sleep(delay)
button = browser.find_element(By.CLASS_NAME,"gradebook-link").click()
time.sleep(delay)

# 2023更新：自動判別行數(thead)及名稱
diet = {}
diet['id']=[]
for i in range(0,len(browser.find_elements(By.XPATH,"/html/body/div[2]/div[2]/section/div/section/div/table/thead/tr/th/div"))):
    diet[browser.find_elements(By.XPATH,"/html/body/div[2]/div[2]/section/div/section/div/table/thead/tr/th/div")[i].text]=[]

# 2023更新：自動辨識頁面及翻頁，不受原限制只能三頁(最多60筆)影響，並美化程式碼
html_source = browser.page_source
soup = BeautifulSoup(html_source, 'lxml')
all_page=re.findall(r'\d+',soup.select_one("html body div div section div section span").text)
nxpage=0
for page in range(0,int(all_page[1])):
    html_source = browser.page_source
    soup = BeautifulSoup(html_source, 'lxml')
    pathlen=len (browser.find_elements(By.XPATH,"/html/body/div[2]/div[2]/section/div/section/table/tbody/tr/td/a"))
    nxpage=nxpage+pathlen
    for i in range(0,pathlen):
        diet['id'].append(browser.find_elements(By.XPATH,"/html/body/div[2]/div[2]/section/div/section/table/tbody/tr/td/a")[i].text)
    for tr in soup.select_one('html body div div section div section div table tbody').select('tr'):
        row=0
        for col in diet:
            if col!='id':
                diet[col].append(tr.select('td')[row].text.strip())
                row=row+1
    print("end{}".format(nxpage))
    url = 'http://gis519.logyuan.idv.tw/courses/course-v1:Pccu_Geography+256600+2023_03/instructor/api/gradebook?offset={}'.format(nxpage)
    browser.get(url)
browser.close()

dit=pd.DataFrame(diet)

score={
    'id':[],
    'Week1': [],
    'Week2': [],
    'Week3': [],
    'Week4': [],
    'Week5': [],
    'Week6': [],
    'Week9': [],
    'Week10': [],
    'Week12': [],
    'Week13': [],
    'Pass':[],
}

# 20230223備註：待思考如何直接辨識該週有哪些欄位，以及應該要有幾分算通過(但分數部分大概還是手動比較快)  
# 可以的話再請郁展幫忙更新這邊的部分  
#未完成待更新
score['id']=dit['id']
dayint=int(dayint[0])
for i in range (0,len(dit)):
    if dayint>=1:
        score["Week1"].append("Y" if int(dit["HW 01"][i])>=11 and int(dit["LB 01"][i])==100 else "N")
    if dayint>=2:
        score["Week2"].append("Y" if int(dit["LB 02"][i])==100 and int(dit["LB 03"][i])==100 and int(dit["LB 04"][i])==100 and int(dit["LB 05"][i])==100 else "N")
    else:
        score["Week2"].append('==')
    if dayint>=3:
        score["Week3"].append("Y" if int(dit["LB 06"][i])==100 else "N")
    else:
        score["Week3"].append('==')
    if dayint>=4:
        score["Week4"].append("Y" if int(dit["LB 08"][i])==100 else "N")
    else:
        score["Week4"].append('==')
    if dayint>=5:
        score["Week5"].append("Y" if int(dit["LB 09"][i])==100 and int(dit["LB 10"][i])==100 else "N")
    else:
        score["Week5"].append('==')
    if dayint>=6:
        score["Week6"].append("Y" if int(dit["LB 07"][i])>25 and int(dit["LB 11"][i])==100 and int(dit["LB 12"][i])==100 else "N")
    else:
        score["Week6"].append('==')
    score["Week9"].append('==')
    score["Week10"].append('==')
    score["Week12"].append('==')
    score["Week13"].append('==')
    #PASS這邊每周要人工更新(尚未想到自動化的方法)
    if dayint==1:
        score["Pass"].append("Pass" if str(score["Week1"][i])=='Y' else "N")
    if dayint==2:
        score["Pass"].append("Pass" if str(score["Week1"][i])=='Y' and str(score["Week2"][i])=='Y' else "N")
    if dayint==3:
        score["Pass"].append("Pass" if str(score["Week1"][i])=='Y' and str(score["Week2"][i])=='Y' and str(score["Week3"][i])=='Y' else "N")
    if dayint==4:
        score["Pass"].append("Pass" if str(score["Week1"][i])=='Y' and str(score["Week2"][i])=='Y' and str(score["Week3"][i])=='Y' and str(score["Week4"][i])=='Y' else "N")
    if dayint==5:
        score["Pass"].append("Pass" if str(score["Week1"][i])=='Y' and str(score["Week2"][i])=='Y' and str(score["Week3"][i])=='Y' and str(score["Week4"][i])=='Y' and str(score["Week5"][i])=='Y' else "N")
    if dayint==6:
        score["Pass"].append("Pass" if str(score["Week1"][i])=='Y' and str(score["Week2"][i])=='Y' and str(score["Week3"][i])=='Y' and str(score["Week4"][i])=='Y' and str(score["Week5"][i])=='Y' and str(score["Week6"][i])=='Y' else "N")
    if dayint==7:
        score["Pass"].append("Pass" if str(score["Week1"][i])=='Y' and str(score["Week2"][i])=='Y' and str(score["Week3"][i])=='Y' and str(score["Week4"][i])=='Y' and str(score["Week5"][i])=='Y' and str(score["Week6"][i])=='Y' else "N")
    if dayint==8:
        score["Pass"].append("Pass" if str(score["Week1"][i])=='Y' and str(score["Week2"][i])=='Y' and str(score["Week3"][i])=='Y' and str(score["Week4"][i])=='Y' and str(score["Week5"][i])=='Y' and str(score["Week6"][i])=='Y' else "N")
    if dayint==9:
        score["Pass"].append("Pass" if str(score["Week1"][i])=='Y' and str(score["Week2"][i])=='Y' and str(score["Week3"][i])=='Y' and str(score["Week4"][i])=='Y' and str(score["Week5"][i])=='Y' and str(score["Week6"][i])=='Y' else "N")
    if dayint==10:
        score["Pass"].append("Pass" if str(score["Week1"][i])=='Y' and str(score["Week2"][i])=='Y' and str(score["Week3"][i])=='Y' and str(score["Week4"][i])=='Y' and str(score["Week5"][i])=='Y' and str(score["Week6"][i])=='Y' else "N")
    if dayint==11:
        score["Pass"].append("Pass" if str(score["Week1"][i])=='Y' and str(score["Week2"][i])=='Y' and str(score["Week3"][i])=='Y' and str(score["Week4"][i])=='Y' and str(score["Week5"][i])=='Y' and str(score["Week6"][i])=='Y' else "N")
    if dayint==12:
        score["Pass"].append("Pass" if str(score["Week1"][i])=='Y' and str(score["Week2"][i])=='Y' and str(score["Week3"][i])=='Y' and str(score["Week4"][i])=='Y' and str(score["Week5"][i])=='Y' and str(score["Week6"][i])=='Y' else "N")
    if dayint==13:
        score["Pass"].append("Pass" if str(score["Week1"][i])=='Y' and str(score["Week2"][i])=='Y' and str(score["Week3"][i])=='Y' and str(score["Week4"][i])=='Y' and str(score["Week5"][i])=='Y' and str(score["Week6"][i])=='Y' else "N")
score=pd.DataFrame(score)
score2=pd.merge(name,score,left_on='Username', right_on='id',how='left').drop(['id','信箱'],axis=1)
score3=score2.sort_values(["組別","Pass",'學號'])
final_score=pd.DataFrame(score3,columns=["系級","學號","姓名","組別","Pass"])

all_stu=len(final_score)
all_len_unfstu=len(final_score[final_score.Pass!='Pass'])
all_unstu=final_score[final_score.Pass!='Pass']['姓名']

# 2023年更新：加入Line Notify執行「成績統計完成後自動通知」任務
# Line Notify實作參考見https://github.com/jiangsir/PythonBasic/blob/master/%E5%AF%A6%E4%BD%9C--%E5%82%B3%E9%80%81Line%20%E8%A8%8A%E6%81%AF.ipynb
token = PP_Information[2] # LINE Notify 權杖
weekmemo=memo[memo.週數==today].reset_index(drop=True)
message = '早安!!成績已統計完成\r\n本周是課程的{4}\r\n統計開始時間：{0}\r\n修課學生總數：{1}\r\n未完成作業人數：{2}\r\n未完成作業同學：\r\n{3}'.format(TodayDateAndTime,all_stu,all_len_unfstu,all_unstu,today)
message3 = '提醒您今日課程安排：\r\n{0}\r\n\n課程備註：\r\n{1}'.format(weekmemo.實習[0],weekmemo.備註[0])

# HTTP 標頭參數與資料
headers = {
    "Authorization": "Bearer " + token, 
    "Content-Type" : "application/x-www-form-urlencoded"
}
data = { 'message': message }
data3 = { 'message': message3 }

message2 = '午安~以下為課程預報\r\n{0}--{1}\r\n\n課程影片進度：\r\n{2}\r\n\n預計實作安排：\r\n{3}\r\n\n課程備註：\r\n{4}\r\n\n影片進度未完成人數：{5}人\r\n(完成率：{6:.2%})'.format(Today,today,weekmemo.課程介紹[0],weekmemo.實習[0],weekmemo.備註[0],all_len_unfstu,(all_stu-all_len_unfstu)/all_stu)
data2 = { 'message': message2 }

# 以 requests 發送 POST 請求
#requests.post("https://notify-api.line.me/api/notify",headers = headers, data = data, files = files)
#時間管理
if (Totime>=datetime.strptime('08:00:00','%H:%M:%S')) and (Totime<=datetime.strptime('10:00:00','%H:%M:%S')):
    requests.post("https://notify-api.line.me/api/notify",
    headers = headers, data = data)
    time.sleep(3)
    requests.post("https://notify-api.line.me/api/notify",
    headers = headers, data = data3)
    # 2023更新：改以xlsx輸出(不想再處理中文漏字問題)，並加入顏色警告
    #顏色標記參照https://xlsxwriter.readthedocs.io/working_with_conditional_formats.html or https://techoverflow.net/2021/09/24/pandas-xlsx-export-with-background-color-based-on-cell-value/
    writer = pd.ExcelWriter('../1112航測完成名單.xlsx', mode='a', engine='openpyxl', if_sheet_exists='new')

    final_score.to_excel(writer, sheet_name='{}完成名單'.format(today),index=False)

    worksheet = writer.sheets['{}完成名單'.format(today)]
    
    for cell, in worksheet[f'E2:E{len(final_score) + 1}']: # Skip header row, process as many rows as there are DataFrames
            value = final_score["Pass"].iloc[cell.row - 2] # value is "True" or "False"
            cell.fill = PatternFill("solid", start_color=("5cb800" if value == "Pass" else 'ff2800'))
    writer.save()
    writer.close()

else:
    requests.post("https://notify-api.line.me/api/notify",
        headers = headers, data = data2)



