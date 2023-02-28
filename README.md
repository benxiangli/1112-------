# 1112航測成績辨識君

### 更新內容(2023.02.27)
+ 改善整體程式碼執行
+ 更新webdriver指令
+ 更改輸出檔案格式及內容格式
+ 加入時間辨識與管理
+ 加入Line Notify功能
---
## 流程說明
 ```mermaid
graph TD;
程式開始--run.bat-->載入環境;
載入環境--identify.py-->讀取課程資訊/權限管理者個人資料-->抓取目前時間-->進行學生成績抓取-->學生成績統計整理-->呼叫LineNotify;
呼叫LineNotify-->判斷時間為非上課當日-->傳送課程預報
呼叫LineNotify-->判斷時間為上課當日08:00至10:00-->傳送成績統計成果--1112航測完成名單.xlsx-->加入當週作業統計成果新頁籤
```
## 文件說明
### run.bat
Windows批次檔，可透過Windows工作排程器設定時段執行bat檔達到定時運作程式的效果  
bat腳本內的路徑內容須自行修改
```
cd /d C:\Users\xiang\Desktop\1112航測資料\1112航測成績辨識君  #前往py檔所在資料夾位置
call C:/Users/xiang/anaconda3/Scripts/activate              #開啟conda虛擬環境位置
call conda activate pywebcrawler                            #切換至指定環境
call python identify.py                                     #執行py腳本
pause
```
### PPI.txt **(必要文件)**
個人資料檔，內容依順序分別為
1. 課程平台電子郵件(需使用課程管理者/助教郵件)  
2. 課程平台密碼  
3. LINE NOTIFY權杖(有關LINE NOTIFY如何設定及指定傳送群組位置可參考[jiangsir文章：實作--傳送Line 訊息.ipynb](https://github.com/jiangsir/PythonBasic/blob/master/%E5%AF%A6%E4%BD%9C--%E5%82%B3%E9%80%81Line%20%E8%A8%8A%E6%81%AF.ipynb))  

此文字檔務必以全形`：`分段

### 航測分組名單.xlsx
學生分組清單，可參考以下格式

|| 組別 | 系級 | Username | 學號 | 姓名 | 信箱 |  
|:---:|:---:|:---:|:---:|:---:|:---:|:---:|  
|說明|課程分組組別|學生系級|課程平台註冊帳號|學生學號|學生姓名|課程平堂註冊信箱(可有可無)|
|範例|第一組|地理二|B1234567|B1234567|王小明|SAMPLE@gmail.com|  

### msedgedriver.exe
微軟edge瀏覽器驅動程式，目前版本 110.0.1587.50 (官方組建) (64 位元)  
請根據自身使用瀏覽器版本及Selenium版本下載對應驅動程式版本
+ [Microsoft Edge](https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/)
+ [Google Chrome](https://sites.google.com/a/chromium.org/chromedriver/)
+ [Mozilla Firefox](https://firefox-source-docs.mozilla.org/testing/geckodriver/Support.html)
+ [Opera](https://github.com/operasoftware/operachromiumdriver/releases)
+ [Safari](https://developer.apple.com/documentation/webkit/about_webdriver_for_safari)
