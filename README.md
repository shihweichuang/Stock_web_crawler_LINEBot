# 股票機器人

參照凱衛資訊股份有限公司 [LINE Bot【查股價】](https://line.me/R/ti/p/@799milhe)，製作一股票機器人，幫助快速查詢台股股價、K線、營收、EPS、股利等資訊。

<br>

# 開發過程紀錄

1. 規劃預期效果、使用者介面、整體使用流程

2. 製作 Rich Menu (即為圖文選單) 版型及對應功能

3. 觀察資訊頁面：需要爬取甚麼資料、所需資料是否橫跨多個頁面

4. 進行爬蟲，取出所有資料

5. 整理爬取下的資料

6. 利用爬取下的資料製成圖表

7. 使用 Python Flask 串接 LINE Bot ，嘗試匯出所需內容

8. 確認成功匯出後，打包成函式

9. 問題克服

<br>

# 開發期間遇到的問題

## 1. Rich Menu 非預設格式

該 Rich Menu 並非一般常見的規格，而是自訂尺寸，故無法直接使用 LINE 官方內建功能，而需要另外使用程式碼進行處理，該部分花了一些時間進行調整。

## 2. 點擊文字後自動帶入文字到對話框

該功能在過去接觸過的 LINE Bot 中較不常見，故花了些時間找尋正確的URI。

## 3. 製作股票相關圖表

該功能在過去較少接觸，故該開發階段有搭配 ChatGPT 進行製作。透過下對指令說明，才能夠為調整成功。

## 4. 在 LINE Bot 中插入指定圖片

過去接觸過的 LINE Bot 功能大多是插入固定的指定圖片，而不是插入會變動的圖片。最後是研究出可以透過 IMGUR API 進行圖片的上傳。

## 5. 爬取資料的時間較長

由於爬取資料的時間較長，可能會讓使用者等得不耐煩，甚至會覺得是系統當機。故針對不會每日更新的資料，固定一週期(如每周、每月)只會執行一次爬取資料的過程，並將爬取下來的資料另存為csv檔案，並於檔案名稱中加上當日日期。在往後的每次執行，都會先判斷是否含有符合條件的檔案名稱。若有，則讀取該檔案；若無，則進行資料爬取。

## 6. 製作時間不長

因需在短期間製作出一版本，故僅先針對幾個部分進行開發，往後可再進行補充。

<br>

# 功能介紹

## 圖文選單

1. 使用者開啟 ---> 【圖文選單】<br>
<img src="https://i.imgur.com/KzRYPkt.jpg" alt="【圖文選單】" width="278" height="202">

2. 使用者點選 ---> 【圖文選單】中的【指令教學】<br>
   系統帶入 ---> 文字【指令】<br>
   系統回傳 ---> Bubble【指令教學】
<img src="https://i.imgur.com/KD9f858.jpg" alt="Bubble【指令教學】" width="278" height="385">

3. 使用者點選 ---> 【圖文選單】中的【好友分享】<br>
   系統開啟 ---> 【選擇傳送對象】
<img src="https://i.imgur.com/H4M3gWG.jpg" alt="【選擇傳送對象】" width="278" height="304">

4. 使用者點選 ---> 【圖文選單】中的【APP推薦】<br>
   系統開啟 ---> 連結
<img src="https://i.imgur.com/wuwY3Ow.jpg" alt="連結_APP推薦" width="278" height="344">

5. 使用者點選 ---> 【圖文選單】中的【股市新聞】<br>
   系統開啟 ---> 連結
<img src="https://i.imgur.com/0IWtddg.jpg" alt="連結_股市新聞" width="278" height="545">

6. 使用者點選 ---> 【圖文選單】中的【名師專欄】<br>
   系統開啟 ---> 連結
<img src="https://i.imgur.com/mRyrPp8.jpg" alt="連結_名師專欄" width="278" height="538">

7. 使用者點選 ---> 【圖文選單】中的【nStock】<br>
   系統開啟 ---> 連結
<img src="https://i.imgur.com/QHrh05Y.jpg" alt="連結_nStock" width="278" height="550">

## 帶入文字到對話框

1. 使用者點選 ---> 台積電股價(代碼P)<br>
   系統帶入 ---> 對話框-文字【台積電股價】
<img src="https://i.imgur.com/jSk3CbJ.jpg" alt="對話框-文字【台積電股價】" width="278" height="394">

2. 使用者點選 ---> 台積電K線(代碼K)<br>
   系統帶入 ---> 對話框-文字【台積電日K】
<img src="https://i.imgur.com/L6i6LNL.jpg" alt="對話框-文字【台積電日K】" width="278" height="394">

## 產出圖表

1. 使用者輸入 ---> 文字【台積電日K】<br>
   系統回傳 ---> Bubble【台積電日K】
<img src="https://i.imgur.com/eRrvfBm.jpg" alt="Bubble【台積電日K】" width="278" height="354">

2. 使用者輸入 ---> 文字【台積電EPS】<br>
   系統回傳 ---> Bubble【台積電EPS】
<img src="https://i.imgur.com/H1He0M0.jpg" alt="Bubble【台積電EPS】" width="278" height="381">

3. 使用者輸入 ---> 文字【台積電營收】<br>
   系統回傳 ---> Bubble【台積電營收】
<img src="https://i.imgur.com/qH4EwFT.jpg" alt="Bubble【台積電營收】" width="278" height="459">

4. 使用者輸入 ---> 文字【台積電持股】<br>
   系統回傳 ---> Bubble【台積電持股】
<img src="https://i.imgur.com/fRSCEza.jpg" alt="Bubble【台積電持股】" width="278" height="432">
