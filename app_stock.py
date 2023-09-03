# 匯入相關模組
from bs4 import BeautifulSoup
import chardet
import csv
from datetime import datetime, time, timedelta, date
from dateutil.relativedelta import relativedelta
from flask import Flask, request, abort
import json
import locale
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.ticker import FormatStrFormatter
matplotlib.use("Agg")                        # 使用 Agg 渲染器，可以在非交互模式下運行，而無須啟動 GUI
import mplfinance as mpf                     # matplotlib旗下專門用於金融分析的繪圖模組
import numpy as np
import os
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import pyimgur
import re
import requests
import textwrap
import time
import twstock
from urllib.parse import quote
import urllib.request

# 載入 LINE Bot 所需要的模組
from linebot import (LineBotApi, WebhookHandler)
from linebot.exceptions import (InvalidSignatureError)
from linebot.models import *

# 記得進去env.json裡面修改成自己要套用的LINE Bot API
with open("env.json") as f:
    env = json.load(f)

# 抓取 LINE Bot 資訊
line_bot_api = LineBotApi(env["YOUR_CHANNEL_ACCESS_TOKEN"])
handler = WebhookHandler(env["YOUR_CHANNEL_SECRET"])

# ----------------------設定imgur----------------------

client_id = env["YOUR_IMGUR_ID"]
access_token = env["IMGUR_TOKEN"]
headers = {"Authorization": f"Bearer {access_token}"}

# ----------------------設定【Rich Menu】----------------------

# 設定 Rich Menu ID
rich_menu_id = env["YOUR_RICH_MENU_ID"]

# 定義Rich Menu圖片的路徑
rich_menu_image_path = "./nstock_rich_menu.jpg"

# 設定【分享好友】URI
uri = "https://line.me/R/share?text=Hi，我發現這個「查股價」機器人可以快速查詢台股股價、K線、營收、EPS、股利等資訊，分享給你！https://line.me/R/ti/p/@636hvuxg"
# URI設定編碼
encoded_uri = quote(uri, safe=":/?&=%")

# 創建Rich Menu物件
rich_menu = RichMenu(
    size=RichMenuSize(width=2500, height=1272),
    selected=True,                                # 固定開啟主選單
    name="主選單",
    chat_bar_text="查股價選單",
    areas=[
        RichMenuArea(
            bounds=RichMenuBounds(x=0, y=0, width=833, height=636),
            action=MessageAction(label="指令教學", text="指令")
        ),
        RichMenuArea(
            bounds=RichMenuBounds(x=833, y=0, width=833, height=636),
            action=URIAction(label="分享好友", uri=encoded_uri)
        ),
        RichMenuArea(
            bounds=RichMenuBounds(x=1666, y=0, width=833, height=636),
            action=URIAction(label="APP推薦", uri="https://shop.nstock.tw/market/")
        ),
        RichMenuArea(
            bounds=RichMenuBounds(x=0, y=636, width=833, height=636),
            action=URIAction(label="股市新聞", uri="https://www.nstock.tw/news/")
        ),
        RichMenuArea(
            bounds=RichMenuBounds(x=833, y=636, width=833, height=636),
            action=URIAction(label="名師專欄", uri="https://www.nstock.tw/author/")
        ),
        RichMenuArea(
            bounds=RichMenuBounds(x=1666, y=636, width=833, height=636),
            action=URIAction(label="nStock", uri="https://www.nstock.tw/")
        )
        # 可以繼續新增其他的互動區域
        # ...
    ]
)

# 上傳圖片並設定Rich Menu
with open(rich_menu_image_path, "rb") as f:
    rich_menu_id = line_bot_api.create_rich_menu(rich_menu=rich_menu)
    line_bot_api.set_rich_menu_image(rich_menu_id, "image/png", f)

# 將Rich Menu與頻道綁定
line_bot_api.set_default_rich_menu(rich_menu_id)

# ------------------------------------------------------------------

# 製作【股票名稱/代碼.csv】
def stock_name_no_csv():

    # 建立請求的URL(=欲爬取網站)
    url = "https://www.nstock.tw/api/v2/stock-list/data"

    # 讀取網頁內容
    response = urllib.request.urlopen(url)
    data = response.read().decode("utf-8")

    json_data = json.loads(data)

    # 建立一空字典，儲存股票名稱/代碼
    stock_dict = {}

    # 製作字典
    for stock_data in json_data["data"]:
        stock_name = stock_data["股票名稱"]
        stock_code = stock_data["股票代號"]
        stock_dict[stock_name] = stock_code

    # -----------------------儲存為CSV-----------------------

    # 取得今日日期
    now_date = date.today().strftime("%Y%m%d")

    # 設定另存檔案路徑
    csv_file_path = f"stock_name_no_{now_date}.csv"

    # 指定欄位名稱
    fieldnames = ["股票名稱", "股票代碼"]

    # 將字典寫入指定檔案
    with open(csv_file_path, "w", newline="", encoding="utf-8-sig") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for stock_name, stock_code in stock_dict.items():
            writer.writerow({"股票名稱": stock_name, "股票代碼": stock_code})

    text = f"已完成 stock_name_no_{now_date}.csv"
    print(text)

# 確認是否有今日檔案【股票名稱/代碼.csv】
def check_file_stock_name_no_csv():

    # 取得今日時間
    today = date.today()

    # 指定檔名
    filename = f"stock_name_no_{today.strftime('%Y%m%d')}.csv"

    # 如果指定檔不存在
    if not os.path.isfile(filename):
        # 製作【股票名稱/代號.csv】
        stock_name_no_csv()

    # 如果指定檔存在
    else:
        # 印出確認文字
        text = f"無須更新 stock_name_no_{today.strftime('%Y%m%d')}.csv"
        print(text)

# 股票名稱/代碼的轉換
def find_stock_code(stockNo):

    # 取得今日時間
    today = date.today()

    # 指定檔名
    filename = f"stock_name_no_{today.strftime('%Y%m%d')}.csv"

    # 檢測csv編碼
    with open(filename, "rb") as file:
        detection = chardet.detect(file.read())
        encoding = detection["encoding"]

    # 讀取csv並進行判斷
    with open(filename, "r", encoding=encoding, errors="ignore") as file:
        reader = csv.DictReader(file)
        for row in reader:
            # 如果輸入資訊在欄位【股票名稱】
            if stockNo in row["股票名稱"]:
                # 回覆欄位【股票代號】的值
                return row["股票代碼"]
            # 如果輸入資訊在欄位【股票代碼】
            if stockNo in row["股票代碼"]:
                # 回覆輸入資訊(即【股票代碼】)
                return stockNo
    return None

# ----------------------------------------------------------------

# 製作【股票資訊.csv】
def stock_info_csv(stockNo):

    # 取得今日日期
    now_date = datetime.now()

    # 格式化日期為"YYYYMMDD"
    formatted_now_date = now_date.strftime("%Y%m%d")

    # https://www.nstock.tw/api/v2/real-time-quotes/data?stock_id=2330
    url = f"https://www.nstock.tw/api/v2/real-time-quotes/data?stock_id={stockNo}"

    # 讀取網頁內容
    response = urllib.request.urlopen(url)
    data = response.read().decode("utf-8")

    # 解析 JSON 資料
    json_data = json.loads(data)

    # ------------取得【股票名稱】------------

    stock_info = "股票名稱"

    stock_name = json_data["data"][0][stock_info]

    # ------------取得【股票代號】------------

    stock_info = "股票代號"

    stock_no = json_data["data"][0][stock_info]

    stock_no = f"({stock_no})"

    # ---------取得【開盤價】---------

    stock_info = "開盤價"

    # 如果小數點後為0，只保留整數
    if json_data["data"][0][stock_info].split(".")[1] == "0000":
        open_price = json_data["data"][0][stock_info].split(".")[0]

    # 如果小數點後不為0，全部顯示
    else:
        open_price = json_data["data"][0][stock_info]

    # ---------取得【最高價】---------

    stock_info = "最高價"

    # 如果小數點後為0，只保留整數
    if json_data["data"][0][stock_info].split(".")[1] == "0000":
        high_price = json_data["data"][0][stock_info].split(".")[0]

    # 如果小數點後不為0，全部顯示
    else:
        high_price = json_data["data"][0][stock_info]

    # ---------取得【最低價】---------

    stock_info = "最低價"

    # 如果小數點後為0，只保留整數
    if json_data["data"][0][stock_info].split(".")[1] == "0000":
        low_price = json_data["data"][0][stock_info].split(".")[0]

    # 如果小數點後不為0，全部顯示
    else:
        low_price = json_data["data"][0][stock_info]

    # ---------取得【收盤價】---------

    stock_info = "當盤成交價"

    # 如果小數點後為0，只保留整數
    if json_data["data"][0][stock_info].split(".")[1] == "0000":
        close_price = json_data["data"][0][stock_info].split(".")[0]

    # 如果小數點後不為0，全部顯示
    else:
        close_price = json_data["data"][0][stock_info]

    # ----------取得【股價更新時間】----------

    stock_info = "最近交易日期"

    update_date = json_data["data"][0][stock_info]

    # 將日期轉換為datetime
    date_obj = datetime.strptime(update_date, "%Y-%m-%d")

    # 格式化為MMDD
    update_date = date_obj.strftime("%m/%d")

    stock_info = "最近成交時刻"

    update_time = json_data["data"][0][stock_info]

    update_date_time = f"{update_date} {update_time}"

    # -----------取得【漲跌】+【漲跌幅】-----------------

    stock_info = "漲跌"

    price_change = json_data["data"][0][stock_info]

    # 如果小數點後為0，只保留整數再加.00
    if price_change.split(".")[1] == "0000":
        price_change = price_change.split(".")[0] + ".00"

    # 如果小數點後不為0，顯示到小數點第二位(需先字串轉為浮點數)
    else:
        price_change = f"{float(price_change):.2f}"

    stock_info = "漲跌幅"

    # 取得漲跌幅(字串)
    price_change_rate = json_data["data"][0][stock_info]

    # 將漲跌幅(字串)轉為浮點數，用於符號判斷
    price_change_rate_test = float(price_change_rate)

    # 將漲跌幅(字串)轉為浮點數，再轉為字串
    price_change_rate = f"({float(price_change_rate)}%)"

    # 【漲跌價】符號判斷
    # 漲跌幅為正
    if price_change_rate_test > 0:
        price_change = "▲ " + price_change

    # 漲跌幅為0
    elif price_change_rate_test == 0:
        price_change = "- " + price_change

    # 漲跌幅為負
    elif price_change_rate_test < 0:
        price_change = "▼ " + price_change

    # -----------------進行爬蟲【產業名稱】---------------------------------------

    # https://www.nstock.tw/api/v2/basic-info/data?stock_id=2330
    url = f"https://www.nstock.tw/api/v2/basic-info/data?stock_id={stockNo}"

    # 使用 urllib.request 套件來讀取網頁內容
    response = urllib.request.urlopen(url)
    data = response.read().decode("utf-8")

    # 解析 JSON 資料
    json_data = json.loads(data)

    # -----------取得【產業名稱】-------------

    stock_info = "產業名稱"

    industry = json_data["data"][0][stock_info]

    # -----------------進行爬蟲【上市櫃】+【成交量】---------------------------------------

    # 建立請求的URL(=欲爬取網站)
    # https://www.nstock.tw/stock_info?status=1&stock_id=2330
    url = f"https://www.nstock.tw/stock_info?status=1&stock_id={stockNo}"

    # 設定請求標頭(模擬正常的瀏覽器行為)
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.51"
    }

    # 發送請求並獲取回應
    res = requests.get(url, headers=headers)

    # 設定編碼(防止中文亂碼)
    res.encoding = "utf-8"

    # 對獲取到的HTML內容進行解析
    soup = BeautifulSoup(res.text, "html.parser")

    # ----------------取得【上市櫃資訊】----------------
    soup1 = soup.find("span", {"class": "rounded-full border border-gray-400 p-0.5 px-2 overflow-clip"})
    text = soup1.text
    listing_cabinet = text.strip()

    # ------------取得【成交量】------------
    soup1 = soup.find_all("div", {"class": "text-base sm:text-2xl my-auto border-r border-gray-300 px-4 pr-2 sm:px-6"})
    text = soup1[0].find_all("div")[1].text
    deal_amount = text.strip()

    # ---------------將特定變數的值插入到字典中，再將字典轉換為DataFrame，再保存為csv---------------

    # 欄位名稱帶入變數
    var1 = update_date
    var2 = update_date_time
    var3 = stock_no
    var4 = stock_name
    var5 = listing_cabinet
    var6 = industry
    var7 = close_price
    var8 = price_change
    var9 = price_change_rate
    var10 = open_price
    var11 = high_price
    var12 = low_price
    var13 = deal_amount

    # 建立字典(每個值都用列表包起來，是為了將變數值轉換為DataFrame時符合格式要求)
    data = {
        "資料日期": [var1],
        "更新時間": [var2],
        "股票代碼": [var3],
        "股票名稱": [var4],
        "上市櫃": [var5],
        "產業別": [var6],
        "收盤價": [var7],
        "漲跌價": [var8],
        "漲跌幅": [var9],
        "開盤": [var10],
        "最高": [var11],
        "最低": [var12],
        "成交量": [var13]
    }

    # 將字典轉換為DataFrame對象
    df = pd.DataFrame(data)

    # 保存DataFrame為csv
    file_name = f"{stockNo}_info.csv"
    file_path = f"./{stockNo}_info.csv"
    df.to_csv(file_path, index=False, encoding="utf-8-sig")

    # 確認文字
    print("已完成 " + file_name)

# 確認是否有最新檔案【股票資訊.csv】
def check_file_stock_info_csv(stockNo):

    # 取得當前時間
    now_time = datetime.now()

    # 取得當前日期
    today_date = now_time.date()

    # 設定台灣股票每日交易時間
    trade_time = time(hour=13, minute=30)

    # 指定檔案路徑
    file_name = f"./{stockNo}_info.csv"

    # 如果指定檔案存在
    if os.path.exists(file_name):
        # 判斷當天是星期幾
        weekday = today_date.weekday()
        # 如果當天是星期一~五
        if weekday <= 4:
            # 抓取日期不變
            today_date = today_date
            # 如果現在時間介於 13:30~24:00
            if now_time.time() >= trade_time:
                # 讀取檔案
                df = pd.read_csv(file_name)
                # 如果"資料日期"欄位包含當天日期
                if today_date.strftime("%m/%d") in df["資料日期"].values:
                    # 列印確認文字
                    print(f"一~五下午，資料抓今天，無須更新 {stockNo}_info.csv")
                # 如果"資料日期"欄位不包含當天日期
                else:
                    # 製作【股票代碼清單】
                    stock_info_csv(stockNo)
            # 如果現在時間介於 00:00~13:30
            else:
                # 讀取檔案
                df = pd.read_csv(file_name)
                # 取得昨天日期
                yesterday = today_date - timedelta(days=1)
                # 如果"資料日期"欄位包含昨天日期
                if yesterday.strftime("%m/%d") in df["資料日期"].values:
                    # 列印確認文字
                    print(f"一~五上午，資料抓昨天，無須更新 {stockNo}_info.csv.csv")
                # 如果"資料日期"欄位不包含昨天日期
                else:
                    # 製作【股票代碼清單】
                    stock_info_csv(stockNo)
        # 如果當天是星期六~日
        else:
            # 抓取日期調整到星期五
            today_date = today_date - timedelta(days=weekday - 4)
            # 讀取檔案
            df = pd.read_csv(file_name)
            # 如果"資料日期"欄位包含當天日期
            if today_date.strftime("%m/%d") in df["資料日期"].values:
                # 列印確認文字
                print(f"六~日，資料抓周五，無須更新 {stockNo}_info.csv.csv")
            # 如果"資料日期"欄位不包含當天日期
            else:
                # 將日期調整回星期六~日
                today_date = now_time.date()
                # 製作【股票代碼清單】
                stock_info_csv(stockNo)
    # 如果指定檔案不存在
    else:
        # 製作【股票代碼清單】
        stock_info_csv(stockNo)

# ----------------------------------------------------------------

# 製作【每股盈餘-圖表+表格】
def EPS_png(stockNo):

    # ----------------------------------歷季收盤(df)----------------------------------

    # 建立請求的URL(=欲爬取網站)
    # https://goodinfo.tw/tw/ShowK_Chart.asp?STOCK_ID=2330&CHT_CAT2=QUAR
    url = "https://goodinfo.tw/tw/ShowK_Chart.asp?STOCK_ID=" + str(stockNo) + "&CHT_CAT2=QUAR"

    # 設定請求標頭(模擬正常的瀏覽器行為)
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.51"
    }

    # 發送請求並獲取回應
    res = requests.get(url, headers=headers)

    # 設定編碼(防止中文亂碼)
    res.encoding = "utf-8"

    # 對獲取到的HTML內容進行解析
    soup = BeautifulSoup(res.text, "html.parser")

    # 提取表格資料
    table = soup.find("table", {"id": "tblPriceDetail", "class": "b1 p4_0 r0_10 row_bg_2n row_mouse_over"})
    rows = table.find_all("tr", {"align": "center"})

    # 提取列資料
    data = []
    # 找到每列中的所有儲存格，提取儲存格的文字內容，並將每列的資料儲存為列表
    for row in rows:
        tds = row.find_all("td")
        row_data = [td.text for td in tds]
        data.append(row_data)
    # 定義包含表格標題的列表，用於後續創建DataFrame
    title = ["交易季度", "交易日數", "開盤", "最高", "最低", "收盤", "漲跌", "漲跌(%)", "振幅(%)", "成交張數(千張)",
             "成交張數(日均)", "成交金額(億元)", "成交金額(日均)", "法人買賣超(千張)-外資", "法人買賣超(千張)-投信",
             "法人買賣超(千張)-自營", "法人買賣超(千張)-合計", "外資持股(%)", "融資(千張)-增減", "融資(千張)-餘額",
             "融券(千張)-增減", "融券(千張)-餘額", "券資比"]

    # 創建DataFrame
    df = pd.DataFrame(data, columns=title)

    # -------------資料處理和篩選-------------

    # 僅保留前17列
    df = df.head(17)
    # 保留指定欄位
    df = df.loc[:, ["交易季度", "收盤"]]
    # 依欄位"交易季度"排序，由小到大
    df = df.sort_values(by="交易季度")
    # 重新排序序號(=重新設置索引)
    open_df = df.reset_index(drop=True)

    # 定義修改函式
    def modify_season(season):
        # 修改【交易季度】列中的內容
        prefix = "20" if int(season[:2]) < 94 else "19"
        return prefix + season

    # 使用修改函式，對【交易季度】欄位進行內容修改
    open_df["交易季度"] = open_df["交易季度"].apply(modify_season)

    # ----------------------------------每股盈餘(df)----------------------------------

    # 建立請求的URL(=欲爬取網站)
    url = f"https://histock.tw/stock/{stockNo}/%E6%AF%8F%E8%82%A1%E7%9B%88%E9%A4%98"
    # 設定請求標頭(模擬正常的瀏覽器行為)
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.51"
    }
    # 發送請求並獲取回應
    res = requests.get(url, headers=headers)
    # 設定編碼(防止中文亂碼)
    res.encoding = "utf-8"
    # 對獲取到的HTML內容進行解析
    soup = BeautifulSoup(res.text, "html.parser")

    # 提取表格資料
    soup1 = soup.find("table", {"class": "tb-stock text-center tbBasic"})
    # 將HTML表格轉換為DataFrame
    df = pd.read_html(str(soup1))[0]
    # 刪除【總計】該列(最後一列)
    df = df.drop(df.index[-1])

    # 將DataFrame轉換為JSON
    json_data = df.to_json(orient="records")
    # 載入JSON檔
    data = json.loads(json_data)
    # 儲存每個季度的EPS資料
    eps_data = {}
    for item in data:
        quarter = item["季別/年度"]
        eps = {}
        for key, value in item.items():
            if key != "季別/年度" and value != "-":
                eps[key + quarter] = float(value)
        eps_data.update(eps)

    # 按照鍵由小到大排序
    sorted_eps_data = dict(sorted(eps_data.items(), key=lambda x: x[0]))
    # 取得倒數17個鍵值對(字典型別)
    EPS = dict(list(sorted_eps_data.items())[-17:])
    # 將字典轉換為DataFrame
    EPS_df = pd.DataFrame.from_dict(EPS, orient="index", columns=["EPS"])
    EPS_df.index.name = "交易季度"
    EPS_df.reset_index(inplace=True)
    # 將欄位"EPS"改為"每股盈餘(元)"
    EPS_df = EPS_df.rename(columns={"EPS": "每股盈餘(元)"})

    # ----------------------------------合併(歷季收盤 df + EPS每股盈餘 df)----------------------------------

    # 合併兩df
    merged_df = pd.merge(open_df, EPS_df, on="交易季度", how="outer")
    # 依交易季度排序
    sorted_df = merged_df.sort_values("交易季度")
    # 刪除含有NaN的列
    cleaned_df = sorted_df.dropna()
    EPS_open_df = cleaned_df
    # 將收盤價轉換為float類型
    EPS_open_df.loc[:, "收盤"] = EPS_open_df["收盤"].astype(float)

    # ----------------------------------製作圖表----------------------------------

    # 設定圖表大小
    fig, ax1 = plt.subplots(figsize=(10, 4))

    # 繪製折線圖 (左Y軸-季收盤價)
    ax1.plot(EPS_open_df["交易季度"],
             EPS_open_df["收盤"],
             marker="o",
             markerfacecolor="white",
             color="red")  # 修改為空心點
    ax1.set_ylabel("季收盤價(元)",
                   color="red",
                   rotation=0,
                   ha="left",
                   fontsize=14)

    # 創建第二個Y軸
    ax2 = ax1.twinx()

    # 繪製長條圖 (右Y軸-每股盈餘)
    ax2.bar(EPS_open_df["交易季度"],
            EPS_open_df["每股盈餘(元)"],
            color="orange",
            alpha=0.5,
            width=0.7)
    ax2.set_ylabel("每股盈餘(元)",
                   color="orange",
                   rotation=0,
                   ha="right",
                   fontsize=14)

    # 調整ax1的y軸範圍
    y1_min = min(EPS_open_df["收盤"])
    y1_max = max(EPS_open_df["收盤"])
    ax1.set_ylim(y1_min * 0.8,
                 y1_max * 1.05)

    # 調整ax2的y軸範圍
    y2_min = min(EPS_open_df["每股盈餘(元)"])
    y2_max = max(EPS_open_df["每股盈餘(元)"])
    ax2.set_ylim(y2_min * 0.9991,
                 y2_max * 1.002)

    # 對齊二Y軸
    ax2.set_yticks(np.linspace(ax2.get_yticks()[0], ax2.get_yticks()[-1], len(ax1.get_yticks())))
    ax2.set_yticks(np.linspace(ax2.get_yticks()[0], ax2.get_yticks()[-1], len(ax1.get_yticks())))

    y1_ticks = ax1.get_yticks()
    y1_pos = ax1.get_yaxis().get_view_interval()

    # 計算資料的長度
    total_ticks = len(EPS_open_df)

    # 設定x軸標籤旋轉角度，修改刻度顏色及文字大小
    ax1.set_xticks(
        [EPS_open_df["交易季度"].iloc[0],
         EPS_open_df["交易季度"].iloc[total_ticks // 4],
         EPS_open_df["交易季度"].iloc[total_ticks // 2],
         EPS_open_df["交易季度"].iloc[total_ticks * 3 // 4],
         EPS_open_df["交易季度"].iloc[-1]])
    ax1.set_xticklabels(
        [EPS_open_df["交易季度"].iloc[0],
         EPS_open_df["交易季度"].iloc[total_ticks // 4],
         EPS_open_df["交易季度"].iloc[total_ticks // 2],
         EPS_open_df["交易季度"].iloc[total_ticks * 3 // 4],
         EPS_open_df["交易季度"].iloc[-1]],
        color="gray", fontsize=14)

    # 修改左Y軸、右Y軸刻度顏色
    ax1.tick_params(axis="y", colors="red")
    ax2.tick_params(axis="y", colors="orange")

    # 修改刻度對應的值格式(小數點第二位)
    ax1.yaxis.set_major_formatter(FormatStrFormatter("%.2f"))
    ax2.yaxis.set_major_formatter(FormatStrFormatter("%.2f"))

    # 調整y軸標籤位置
    ax1.yaxis.set_label_coords(0, 1.04)
    ax2.yaxis.set_label_coords(1, 1.10)

    # 設定中文字體
    plt.rcParams["font.sans-serif"] = ["Microsoft JhengHei"]
    plt.rcParams["axes.unicode_minus"] = False

    # 設定大標題
    plt.title("")

    # 設定邊框顏色
    ax1.spines["top"].set_color("gray")
    ax1.spines["bottom"].set_color("gray")
    ax1.spines["left"].set_color("gray")
    ax1.spines["right"].set_color("gray")
    ax2.spines["top"].set_color("gray")
    ax2.spines["bottom"].set_color("gray")
    ax2.spines["left"].set_color("gray")
    ax2.spines["right"].set_color("gray")

    # 調整X軸、兩個Y軸的標註文字大小
    ax1.xaxis.set_tick_params(labelsize=14)
    ax1.yaxis.set_tick_params(labelsize=14)
    ax2.yaxis.set_tick_params(labelsize=14)

    # 新增圖例
    ax1.legend(["季收盤價"],
               loc="upper left",
               bbox_to_anchor=(0.29, 1.17),
               fontsize=14,
               frameon=False)  # 第一個圖例文字大小
    ax2.legend(["每股盈餘"],
               loc="upper right",
               bbox_to_anchor=(0.705, 1.17),
               fontsize=14,
               frameon=False)  # 第二個圖例文字大小

    # 繪製右Y軸刻度0的直線
    ax2.axhline(0, color="gray", linestyle="--", linewidth=0.5, zorder=0)

    # 加入格線
    ax1.xaxis.grid(True, color="lightgray", linestyle="--", zorder=0)
    ax2.grid(True, color="lightgray", linestyle="--", zorder=0)

    # 儲存圖片
    plt.savefig(str(stockNo) + "每股盈餘-圖表.png", bbox_inches="tight")

    # ----------------------------------每股盈餘(表格)----------------------------------

    # 建立請求的URL(=欲爬取網站)
    url = f"https://histock.tw/stock/{stockNo}/%E6%AF%8F%E8%82%A1%E7%9B%88%E9%A4%98"
    # 設定請求標頭(模擬正常的瀏覽器行為)
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.51"
    }
    # 發送請求並獲取回應
    res = requests.get(url, headers=headers)
    # 設定編碼(防止中文亂碼)
    res.encoding = "utf-8"
    # 對獲取到的HTML內容進行解析
    soup = BeautifulSoup(res.text, "html.parser")

    # 提取表格資料
    soup1 = soup.find("table", {"class": "tb-stock text-center tbBasic"})
    # 將HTML表格轉換為DataFrame
    df = pd.read_html(str(soup1))[0]

    # 顛倒第2~10欄的順序
    reversed_columns = df.columns[:1].tolist() + df.columns[10:1:-1].tolist()
    df = df[reversed_columns]

    # 只保留前4欄
    df = df.iloc[:, :4]

    # 第一欄名稱改為季
    df = df.rename(columns={df.columns[0]: "季"})
    new_column_name = df.columns[1] + " EPS(元)"
    df = df.rename(columns={df.columns[1]: new_column_name})
    new_column_name = df.columns[2] + " EPS(元)"
    df = df.rename(columns={df.columns[2]: new_column_name})
    new_column_name = df.columns[3] + " EPS(元)"
    df = df.rename(columns={df.columns[3]: new_column_name})

    # 將第二欄、第四欄和第五欄的字串型態轉換為浮點數型態，並將其格式樺為字串，保留小數點後兩位；如果遇到"-"字元則忽略
    df[df.columns[1]] = df[df.columns[1]].apply(lambda x: "{:.2f}".format(float(x)) if x != "-" else x)
    df[df.columns[2]] = df[df.columns[2]].apply(lambda x: "{:.2f}".format(float(x)))
    df[df.columns[3]] = df[df.columns[3]].apply(lambda x: "{:.2f}".format(float(x)))

    # 插入新的列到第三列
    df.insert(2, "年增(%)", "")

    # 計算新列的值
    for index, row in df.iterrows():
        if row[df.columns[1]] == "-":
            df.at[index, "年增(%)"] = ""
        else:
            value = (float(row[df.columns[1]]) - float(row[df.columns[3]])) / float(row[df.columns[3]]) * 100
            df.at[index, "年增(%)"] = f"{value:.2f}%"

    # 獲取"年增(%)"列的第1行的值
    value = df.at[0, "年增(%)"]

    # 將"年增(%)"列的第5行的值改為與第1行相同的值
    df.at[4, "年增(%)"] = value

    # 將第一欄的第5個值更新
    df.iloc[4, 0] = "累計"

    # -----------------------製作表格-----------------------

    # 設定中文字體
    plt.rcParams["font.sans-serif"] = ["Microsoft JhengHei"]
    plt.rcParams["axes.unicode_minus"] = False

    # 繪製圖表
    fig, ax = plt.subplots()
    ax.axis("off")  # 隐藏坐标轴
    table = ax.table(cellText=df.values,
                     colLabels=df.columns,
                     cellLoc="right",
                     loc="center")

    # 指定需要居中对齐的列索引
    center_align_columns = [0]

    # 設置指定欄位的對齊方式為居中
    for row in range(6):
        for column in center_align_columns:
            cell = table._cells[(row, column)]
            cell.set_text_props(ha="center")

    # 設定表頭列的背景顏色
    for i in range(len(df.columns)):
        cell = table[0, i]
        cell.set_facecolor("#F6F6FC")

    # 設定表格框線顏色為灰色
    for key, cell in table.get_celld().items():
        cell.set_linewidth(0.5)
        cell.set_edgecolor("#c3c6c9")

    # 表格格式
    table.auto_set_font_size(False)
    # 文字大小
    table.set_fontsize(20)
    # 調整寬度、高度
    table.scale(2, 3)

    # 【年增(%)】顏色設定
    for row in range(1, 6):
        cell = table._cells[(row, 2)]  # 【年增(%)】列索列為2

        cell_text = cell.get_text().get_text().replace(",", "").replace("%", "")
        # 【年增(%)】的值
        try:
            cell_value = float(cell_text)
        except ValueError:
            cell_value = 0.0  # 或者任何你認為適合的預設值

        # 如果【年增(%)】的值大於0
        if cell_value > 0:
            # 【年增(%)】值改為紅色
            cell.set_text_props(color="red")

        elif cell_value == 0:
            # 【年增(%)】值改為黑色
            cell.set_text_props(color="black")

        # 如果【年增(%)】的值小於0
        else:
            # 【年增(%)】值改為綠色
            cell.set_text_props(color="green")

    # 儲存圖片
    image_path = str(stockNo) + "每股盈餘-表格.png"
    plt.savefig(image_path, bbox_inches="tight", pad_inches=0.05)

    # ----------------------------------表格裁切空白處----------------------------------

    # 讀取圖片
    image_path = str(stockNo) + "每股盈餘-表格.png"
    image = Image.open(image_path)

    # 裁剪圖片
    left = 2  # 裁剪框左上角的x座標
    right = 1001  # 裁剪框右下角的x座標
    top = 33  # 裁剪框左上角的y座標
    bottom = 345  # 裁剪框右下角的y座標
    cropped_image = image.crop((left, top, right, bottom))

    # 儲存裁剪後的圖片
    image_path = str(stockNo) + "每股盈餘-表格-裁剪.png"
    cropped_image.save(image_path)

    # ----------------------------------依比例調整表格----------------------------------

    # -------取得【圖表寬度】-------

    # 讀取圖片
    image_path = str(stockNo) + "每股盈餘-圖表.png"
    image = Image.open(image_path)

    # 取得圖片的寬度
    width_target = image.width

    # -------調整【表格寬度】與【圖表寬度】相同-------

    # 讀取圖片
    image_path = str(stockNo) + "每股盈餘-表格-裁剪.png"
    image = Image.open(image_path)

    # 取得圖片的寬度和高度
    width, height = image.size

    # 計算縮放比例
    scale_factor = width_target / width

    # 計算縮放後的寬度和高度
    new_width = int(width * scale_factor)
    new_height = int(height * scale_factor)

    # 縮放圖片
    resized_image = image.resize((new_width, new_height))

    # 輸出結果
    image_path = str(stockNo) + "每股盈餘-表格-裁剪-調寬度.png"
    resized_image.save(image_path)

    # ----------------------------------合併PNG*2----------------------------------

    # 讀取第一張圖片
    image1_path = str(stockNo) + "每股盈餘-圖表.png"
    image1 = Image.open(image1_path)

    # 讀取第二張圖片
    image2_path = str(stockNo) + "每股盈餘-表格-裁剪-調寬度.png"
    image2 = Image.open(image2_path)

    # 獲取第一張圖片的寬度和高度
    width1, height1 = image1.size

    # 獲取第二張圖片的寬度和高度
    width2, height2 = image2.size

    # 確定合併後圖片的寬度和高度
    merged_width = max(width1, width2)
    merged_height = height1 + height2

    # 創建一個新的空白圖片，尺寸為合併後的寬度和高度
    merged_image = Image.new("RGB", (merged_width, merged_height))

    # 將第一張圖片黏貼到合併圖片的上半部分
    merged_image.paste(image1, (0, 0))

    # 將第二張圖片黏貼到合併圖片的下半部分
    merged_image.paste(image2, (0, height1))

    # 取得當前時間
    now = datetime.now()

    # 提取當前時間的年份和月份
    month = now.month

    # 計算前一個月份
    # 如果當前月份為1月，則回推到前一年的12月
    if month == 1:
        previous_month = 12
    else:
        previous_month = month - 1

    # 儲存合併後的圖片
    image_path = str(stockNo) + "每股盈餘_截至" + str(previous_month) + "月.png"
    merged_image.save(image_path)

# 確認是否有【每股盈餘-圖表+表格】
def check_file_EPS_png(stockNo):

    # 取得當前時間
    now = datetime.now()

    # 提取當前時間的年份和月份
    month = now.month

    # 計算前一個月份
    # 如果當前月份為1月，則回推到前一年的12月
    if month == 1:
        previous_month = 12
    else:
        previous_month = month - 1

    # 指定檔名
    filename = str(stockNo) + "每股盈餘_截至" + str(previous_month) + "月.png"

    # 如果指定檔不存在
    if not os.path.isfile(filename):
        # 製作【股票名稱/代號.csv】
        EPS_png(stockNo)
        # 印出確認文字
        print("已完成 " + filename)

    # 如果指定檔存在
    else:
        # 印出確認文字
        print("無須更新 " + filename)

# EPS-bubble
def EPS(stockNo):

    # 製作【股票資訊.csv】
    stock_info_csv(stockNo)

    # 確認是否有【每股盈餘-圖表+表格】
    check_file_EPS_png(stockNo)

    # 取得當前時間
    now = datetime.now()

    # 提取當前時間的年份和月份
    month = now.month

    # 計算前一個月份
    # 如果當前月份為1月，則回推到前一年的12月
    if month == 1:
        previous_month = 12
    else:
        previous_month = month - 1

    # 指定檔名
    filename = str(stockNo) + "每股盈餘_截至" + str(previous_month) + "月.png"

    # ------------------圖片上傳到Imgur------------------

    # 讀取指定檔案，取得IMGUR_ID
    with open("env.json") as f:
        env = json.load(f)

    # 讀入IMGUR_ID
    CLIENT_ID = env["YOUR_IMGUR_ID"]

    # 指定圖片路徑
    PATH = filename

    # 指定圖片標題
    title = filename

    # 將圖片上傳到IMGUR
    im = pyimgur.Imgur(CLIENT_ID)

    # 帶入圖片路徑及標題
    uploaded_image = im.upload_image(PATH, title=title)

    # ------------------------------讀取stock_info.csv------------------------------

    # 指定檔案路徑
    file_name = f"./{stockNo}_info.csv"

    # 讀取指定檔案
    df = pd.read_csv(file_name)

    # 列出欄位下的值
    update_date = df["資料日期"].tolist()[0]
    update_date_time = df["更新時間"].tolist()[0]
    stock_no = df["股票代碼"].tolist()[0]
    stock_name = df["股票名稱"].tolist()[0]
    listing_cabinet = df["上市櫃"].tolist()[0]
    industry = df["產業別"].tolist()[0]

    close_price = df["收盤價"].tolist()[0]
    # 如果值有小數點，就改為小數點後兩位(字串)
    if len(str(close_price).split(".")) == 2:
        close_price = "{:.2f}".format(close_price)
    # 如果值為整數，就維持(整數)
    else:
        close_price = close_price

    price_change = df["漲跌價"].tolist()[0]
    price_change_rate = df["漲跌幅"].tolist()[0]

    open_price = df["開盤"].tolist()[0]
    # 如果值有小數點，就改為小數點後兩位(字串)
    if len(str(open_price).split(".")) == 2:
        open_price = "{:.2f}".format(open_price)
    # 如果值為整數，就維持(整數)
    else:
        open_price = open_price

    high_price = df["最高"].tolist()[0]
    # 如果值有小數點，就改為小數點後兩位(字串)
    if len(str(high_price).split(".")) == 2:
        high_price = "{:.2f}".format(high_price)
    # 如果值為整數，就維持(整數)
    else:
        high_price = high_price

    low_price = df["最低"].tolist()[0]
    # 如果值有小數點，就改為小數點後兩位(字串)
    if len(str(low_price).split(".")) == 2:
        low_price = "{:.2f}".format(low_price)
    # 如果值為整數，就維持(整數)
    else:
        low_price = low_price

    deal_amount = df["成交量"].tolist()[0]

    # 建立FlexSendMessage
    message = FlexSendMessage(
        alt_text = str(stock_name) + str(stock_no) + " 季EPS",
        contents = {
            "type": "bubble",
            "size": "giga",
            "header": {
                "type": "box",
                "layout": "vertical",
                "contents": [
                    {
                        "type": "box",
                        "layout": "horizontal",
                        "contents": [
                            {
                                "type": "text",
                                "text": "⇧" + " 加好友",
                                "color": "#d90b0b",
                                "weight": "bold",
                                "size": "xl",
                                "action": {
                                    "type": "uri",
                                    "label": "action",
                                    "uri": "https://line.me/R/ti/p/%40636hvuxg"
                                }
                            }
                        ]
                    },
                    {
                        "type": "box",
                        "layout": "horizontal",
                        "contents": [
                            {
                                "type": "text",
                                "text": str(stock_name),
                                "size": "xxl",
                                "weight": "bold",
                                "flex": 0
                            },
                            {
                                "type": "text",
                                "text": str(stock_no),
                                "gravity": "bottom",
                                "flex": 2
                            },
                            {
                                "type": "filler"
                            },
                            {
                                "type": "text",
                                "text": str(close_price),
                                "size": "xxl",
                                "flex": 3,
                                "align": "end"
                            },
                            {
                                "type": "box",
                                "layout": "vertical",
                                "contents": [
                                    {
                                        "type": "text",
                                        "text": str(price_change),
                                        "align": "center",
                                        "size": "xs"
                                    },
                                    {
                                        "type": "text",
                                        "text": str(price_change_rate),
                                        "align": "center",
                                        "size": "xs"
                                    }
                                ],
                                "flex": 0,
                                "margin": "sm"
                            }
                        ]
                    },
                    {
                        "type": "box",
                        "layout": "baseline",
                        "contents": [
                            {
                                "type": "text",
                                "text": str(listing_cabinet) + " / " + str(industry),
                                "flex": 1,
                                "size": "xs",
                                "color": "#C5C4C7"
                            },
                            {
                                "type": "text",
                                "text": "股價更新時間：" + str(update_date_time),
                                "size": "xs",
                                "flex": 2,
                                "align": "end",
                                "color": "#C5C4C7"
                            }
                        ]
                    }
                ],
                "backgroundColor": "#F6F6FC"
            },
            "hero": {
                "type": "image",
                "url": uploaded_image.link,
                "size": "full",
                "aspectRatio": "12:9",
                "aspectMode": "fit",
                "action": {
                    "type": "uri",
                    "label": "action",
                    "uri": "https://www.nstock.tw/stock_info?status=6&stock_id=" + str(stockNo) + "&utm_source=line&utm_medium=line_bot"
                }
            },
            "footer": {
                "type": "box",
                "layout": "horizontal",
                "spacing": "none",
                "contents": [
                    {
                        "type": "text",
                        "text": "選項：",
                        "size": "xs",
                        "color": "#787878"
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "即時",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": "查" + str(stock_name)
                        }
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "K線",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "日K"
                        }
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "法人",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "法人"
                        }
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "EPS",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "EPS"
                        },
                        "color": "#d11111",
                        "weight": "bold"
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "營收",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "營收"
                        }
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "股利",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "股利"
                        }
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "持股",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "持股"
                        }
                    }
                ]
            }
        }
    )

    return message

# ----------------------------------------------------------------

# 製作【持股-圖表+表格】
def shareholder_png(stockNo, filename):

    # ----------------------開始爬蟲----------------------

    # 建立請求的URL(=欲爬取網站)
    # https://norway.twsthr.info/StockHolders.aspx?stock=2330
    url = f"https://norway.twsthr.info/StockHolders.aspx?stock={stockNo}"

    # 設定請求標頭(模擬正常的瀏覽器行為)
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.51"
    }

    # 發送請求並獲取回應
    res = requests.get(url, headers=headers)

    # 設定編碼(防止中文亂碼)
    res.encoding = "utf-8"

    # 對獲取到的HTML內容進行解析
    soup = BeautifulSoup(res.text, "html.parser")

    # ----------取得【資料日期】----------

    dates = []

    soup = BeautifulSoup(res.text, "html.parser")
    soup1 = soup.find_all("tr", {"class": "lLS"})[6]
    match = re.search(r"<td>(\d{8})\s*</td>", str(soup1))
    if match:
        dates.append(match.group(1))

    soup1 = soup.find_all("tr", {"class": "lDS"})[7]
    match = re.search(r"<td>(\d{8})\s*</td>", str(soup1))
    if match:
        dates.append(match.group(1))

    soup1 = soup.find_all("tr", {"class": "lLS"})[5]
    match = re.search(r"<td>(\d{8})\s*</td>", str(soup1))
    if match:
        dates.append(match.group(1))

    soup1 = soup.find_all("tr", {"class": "lDS"})[6]
    match = re.search(r"<td>(\d{8})\s*</td>", str(soup1))
    if match:
        dates.append(match.group(1))

    soup1 = soup.find_all("tr", {"class": "lLS"})[4]
    match = re.search(r"<td>(\d{8})\s*</td>", str(soup1))
    if match:
        dates.append(match.group(1))

    soup1 = soup.find_all("tr", {"class": "lDS"})[5]
    match = re.search(r"<td>(\d{8})\s*</td>", str(soup1))
    if match:
        dates.append(match.group(1))

    soup1 = soup.find_all("tr", {"class": "lLS"})[3]
    match = re.search(r"<td>(\d{8})\s*</td>", str(soup1))
    if match:
        dates.append(match.group(1))

    soup1 = soup.find_all("tr", {"class": "lDS"})[4]
    match = re.search(r"<td>(\d{8})\s*</td>", str(soup1))
    if match:
        dates.append(match.group(1))

    soup1 = soup.find_all("tr", {"class": "lLS"})[2]
    match = re.search(r"<td>(\d{8})\s*</td>", str(soup1))
    if match:
        dates.append(match.group(1))

    soup1 = soup.find_all("tr", {"class": "lDS"})[3]
    match = re.search(r"<td>(\d{8})\s*</td>", str(soup1))
    if match:
        dates.append(match.group(1))

    soup1 = soup.find_all("tr", {"class": "lLS"})[1]
    match = re.search(r"<td>(\d{8})\s*</td>", str(soup1))
    if match:
        dates.append(match.group(1))

    soup1 = soup.find_all("tr", {"class": "lDS"})[2]
    match = re.search(r"<td>(\d{8})\s*</td>", str(soup1))
    if match:
        dates.append(match.group(1))

    soup1 = soup.find_all("tr", {"class": "lLS"})[0]
    match = re.search(r"<td>(\d{8})\s*</td>", str(soup1))
    if match:
        dates.append(match.group(1))

    soup1 = soup.find_all("tr", {"class": "lDS"})[1]
    match = re.search(r"<td>(\d{8})\s*</td>", str(soup1))
    if match:
        dates.append(match.group(1))

    # 調整日期格式
    date_list = []

    for date_str in dates:
        date_obj = datetime.strptime(date_str, "%Y%m%d")
        formatted_date = date_obj.strftime("%Y/%m/%d")
        date_list.append(formatted_date)

    # --------------取得【>1000張 人數】--------------

    # 取得第13個【>1000張人數=大股東人數】
    soup1 = soup.find_all("tr", {"class": "lDS"})[7]
    value = soup1.find_all("td")[12].get_text(strip=True)
    shareholder_num_13 = locale.format_string("%d", int(value),
                                              grouping=True)  # Format the number with comma separator

    # 取得第11個【>1000張人數=大股東人數】
    soup1 = soup.find_all("tr", {"class": "lDS"})[6]
    value = soup1.find_all("td")[12].get_text(strip=True)
    shareholder_num_11 = locale.format_string("%d", int(value),
                                              grouping=True)  # Format the number with comma separator

    # 取得第9個【>1000張人數=大股東人數】
    soup1 = soup.find_all("tr", {"class": "lDS"})[5]
    value = soup1.find_all("td")[12].get_text(strip=True)
    shareholder_num_9 = locale.format_string("%d", int(value),
                                             grouping=True)  # Format the number with comma separator

    # 取得第7個【>1000張人數=大股東人數】
    soup1 = soup.find_all("tr", {"class": "lDS"})[4]
    value = soup1.find_all("td")[12].get_text(strip=True)
    shareholder_num_7 = locale.format_string("%d", int(value),
                                             grouping=True)  # Format the number with comma separator

    # 取得第5個【>1000張人數=大股東人數】
    soup1 = soup.find_all("tr", {"class": "lDS"})[3]
    value = soup1.find_all("td")[12].get_text(strip=True)
    shareholder_num_5 = locale.format_string("%d", int(value),
                                             grouping=True)  # Format the number with comma separator

    # 取得3第個【>1000張人數=大股東人數】
    soup1 = soup.find_all("tr", {"class": "lDS"})[2]
    value = soup1.find_all("td")[12].get_text(strip=True)
    shareholder_num_3 = locale.format_string("%d", int(value),
                                             grouping=True)  # Format the number with comma separator

    # 取得第1個【>1000張人數=大股東人數】
    soup1 = soup.find_all("tr", {"class": "lDS"})[1]
    value = soup1.find_all("td")[12].get_text(strip=True)
    shareholder_num_1 = locale.format_string("%d", int(value),
                                             grouping=True)  # Format the number with comma separator

    # 取得第14個【>1000張人數=大股東人數】
    soup1 = soup.find_all("tr", {"class": "lLS"})[6]
    value = soup1.find_all("td")[12].get_text(strip=True)
    shareholder_num_14 = locale.format_string("%d", int(value),
                                              grouping=True)  # Format the number with comma separator

    # 取得第12個【>1000張人數=大股東人數】
    soup1 = soup.find_all("tr", {"class": "lLS"})[5]
    value = soup1.find_all("td")[12].get_text(strip=True)
    shareholder_num_12 = locale.format_string("%d", int(value),
                                              grouping=True)  # Format the number with comma separator

    # 取得第10個【>1000張人數=大股東人數】
    soup1 = soup.find_all("tr", {"class": "lLS"})[4]
    value = soup1.find_all("td")[12].get_text(strip=True)
    shareholder_num_10 = locale.format_string("%d", int(value),
                                              grouping=True)  # Format the number with comma separator

    # 取得第8個【>1000張人數=大股東人數】
    soup1 = soup.find_all("tr", {"class": "lLS"})[3]
    value = soup1.find_all("td")[12].get_text(strip=True)
    shareholder_num_8 = locale.format_string("%d", int(value),
                                             grouping=True)  # Format the number with comma separator

    # 取得第6個【>1000張人數=大股東人數】
    soup1 = soup.find_all("tr", {"class": "lLS"})[2]
    value = soup1.find_all("td")[12].get_text(strip=True)
    shareholder_num_6 = locale.format_string("%d", int(value),
                                             grouping=True)  # Format the number with comma separator

    # 取得第4個【>1000張人數=大股東人數】
    soup1 = soup.find_all("tr", {"class": "lLS"})[1]
    value = soup1.find_all("td")[12].get_text(strip=True)
    shareholder_num_4 = locale.format_string("%d", int(value),
                                             grouping=True)  # Format the number with comma separator

    # 取得第2個【>1000張人數=大股東人數】
    soup1 = soup.find_all("tr", {"class": "lLS"})[0]
    value = soup1.find_all("td")[12].get_text(strip=True)
    shareholder_num_2 = locale.format_string("%d", int(value),
                                             grouping=True)  # Format the number with comma separator

    shareholder_num_list = []

    shareholder_num_list.extend([
        shareholder_num_14,
        shareholder_num_13,
        shareholder_num_12,
        shareholder_num_11,
        shareholder_num_10,
        shareholder_num_9,
        shareholder_num_8,
        shareholder_num_7,
        shareholder_num_6,
        shareholder_num_5,
        shareholder_num_4,
        shareholder_num_3,
        shareholder_num_2,
        shareholder_num_1
    ])

    # -------------取得【收盤價】-------------

    # 取得第14個【收盤價】
    result = soup.find_all("tr", {"class": "lLS"})[6].find_all("td")[14].text
    close_14 = float(result)  # 轉為浮點數

    # 取得第13個【收盤價】
    result = soup.find_all("tr", {"class": "lDS"})[7].find_all("td")[14].text
    close_13 = float(result)  # 轉為浮點數

    # 取得第12個【收盤價】
    result = soup.find_all("tr", {"class": "lLS"})[5].find_all("td")[14].text
    close_12 = float(result)  # 轉為浮點數

    # 取得第11個【收盤價】
    result = soup.find_all("tr", {"class": "lDS"})[6].find_all("td")[14].text
    close_11 = float(result)  # 轉為浮點數

    # 取得第10個【收盤價】
    result = soup.find_all("tr", {"class": "lLS"})[4].find_all("td")[14].text
    close_10 = float(result)  # 轉為浮點數

    # 取得第9個【收盤價】
    result = soup.find_all("tr", {"class": "lDS"})[5].find_all("td")[14].text
    close_9 = float(result)  # 轉為浮點數

    # 取得第8個【收盤價】
    result = soup.find_all("tr", {"class": "lLS"})[3].find_all("td")[14].text
    close_8 = float(result)  # 轉為浮點數

    # 取得第7個【收盤價】
    result = soup.find_all("tr", {"class": "lDS"})[4].find_all("td")[14].text
    close_7 = float(result)  # 轉為浮點數

    # 取得第6個【收盤價】
    result = soup.find_all("tr", {"class": "lLS"})[2].find_all("td")[14].text
    close_6 = float(result)  # 轉為浮點數

    # 取得第5個【收盤價】
    result = soup.find_all("tr", {"class": "lDS"})[3].find_all("td")[14].text
    close_5 = float(result)  # 轉為浮點數

    # 取得第4個【收盤價】
    result = soup.find_all("tr", {"class": "lLS"})[1].find_all("td")[14].text
    close_4 = float(result)  # 轉為浮點數

    # 取得第3個【收盤價】
    result = soup.find_all("tr", {"class": "lDS"})[2].find_all("td")[14].text
    close_3 = float(result)  # 轉為浮點數

    # 取得第2個【收盤價】
    result = soup.find_all("tr", {"class": "lLS"})[0].find_all("td")[14].text
    close_2 = float(result)  # 轉為浮點數

    # 取得第1個【收盤價】
    result = soup.find_all("tr", {"class": "lDS"})[1].find_all("td")[14].text
    close_1 = float(result)  # 轉為浮點數

    close_list = []

    close_list.extend([
        close_14,
        close_13,
        close_12,
        close_11,
        close_10,
        close_9,
        close_8,
        close_7,
        close_6,
        close_5,
        close_4,
        close_3,
        close_2,
        close_1
    ])

    # -------【>1000張 大股東持有百分比】-------

    # 取得第1個【>1000張 大股東持有百分比】
    soup1 = soup.find_all("tr", {"class": "lDS"})[1]
    shareholder_percent_1 = soup1.find_all("td")[13].text.strip()
    shareholder_percent_1 = float(shareholder_percent_1)

    # 取得第2個【>1000張 大股東持有百分比】
    soup1 = soup.find_all("tr", {"class": "lLS"})[0]
    shareholder_percent_2 = soup1.find_all("td")[13].text.strip()
    shareholder_percent_2 = float(shareholder_percent_2)

    # 取得第3個【>1000張 大股東持有百分比】
    soup1 = soup.find_all("tr", {"class": "lDS"})[2]
    shareholder_percent_3 = soup1.find_all("td")[13].text.strip()
    shareholder_percent_3 = float(shareholder_percent_3)

    # 取得第4個【>1000張 大股東持有百分比】
    soup1 = soup.find_all("tr", {"class": "lLS"})[1]
    shareholder_percent_4 = soup1.find_all("td")[13].text.strip()
    shareholder_percent_4 = float(shareholder_percent_4)

    # 取得第5個【>1000張 大股東持有百分比】
    soup1 = soup.find_all("tr", {"class": "lDS"})[3]
    shareholder_percent_5 = soup1.find_all("td")[13].text.strip()
    shareholder_percent_5 = float(shareholder_percent_5)

    # 取得第6個【>1000張 大股東持有百分比】
    soup1 = soup.find_all("tr", {"class": "lLS"})[2]
    shareholder_percent_6 = soup1.find_all("td")[13].text.strip()
    shareholder_percent_6 = float(shareholder_percent_6)

    # 取得第7個【>1000張 大股東持有百分比】
    soup1 = soup.find_all("tr", {"class": "lDS"})[4]
    shareholder_percent_7 = soup1.find_all("td")[13].text.strip()
    shareholder_percent_7 = float(shareholder_percent_7)

    # 取得第8個【>1000張 大股東持有百分比】
    soup1 = soup.find_all("tr", {"class": "lLS"})[3]
    shareholder_percent_8 = soup1.find_all("td")[13].text.strip()
    shareholder_percent_8 = float(shareholder_percent_8)

    # 取得第9個【>1000張 大股東持有百分比】
    soup1 = soup.find_all("tr", {"class": "lDS"})[5]
    shareholder_percent_9 = soup1.find_all("td")[13].text.strip()
    shareholder_percent_9 = float(shareholder_percent_9)

    # 取得第10個【>1000張 大股東持有百分比】
    soup1 = soup.find_all("tr", {"class": "lLS"})[4]
    shareholder_percent_10 = soup1.find_all("td")[13].text.strip()
    shareholder_percent_10 = float(shareholder_percent_10)

    # 取得第11個【>1000張 大股東持有百分比】
    soup1 = soup.find_all("tr", {"class": "lDS"})[6]
    shareholder_percent_11 = soup1.find_all("td")[13].text.strip()
    shareholder_percent_11 = float(shareholder_percent_11)

    # 取得第12個【>1000張 大股東持有百分比】
    soup1 = soup.find_all("tr", {"class": "lLS"})[5]
    shareholder_percent_12 = soup1.find_all("td")[13].text.strip()
    shareholder_percent_12 = float(shareholder_percent_12)

    # 取得第13個【>1000張 大股東持有百分比】
    soup1 = soup.find_all("tr", {"class": "lDS"})[7]
    shareholder_percent_13 = soup1.find_all("td")[13].text.strip()
    shareholder_percent_13 = float(shareholder_percent_13)

    # 取得第14個【>1000張 大股東持有百分比】
    soup1 = soup.find_all("tr", {"class": "lLS"})[6]
    shareholder_percent_14 = soup1.find_all("td")[13].text.strip()
    shareholder_percent_14 = float(shareholder_percent_14)

    shareholder_percent_list = []

    shareholder_percent_list.extend([
        shareholder_percent_14,
        shareholder_percent_13,
        shareholder_percent_12,
        shareholder_percent_11,
        shareholder_percent_10,
        shareholder_percent_9,
        shareholder_percent_8,
        shareholder_percent_7,
        shareholder_percent_6,
        shareholder_percent_5,
        shareholder_percent_4,
        shareholder_percent_3,
        shareholder_percent_2,
        shareholder_percent_1
    ])

    # ----------------合併資料----------------

    data = {"資料時間": date_list,
            "收盤價": close_list,
            "大股東持有比": shareholder_percent_list,
            "人數": shareholder_num_list}

    df = pd.DataFrame(data)

    # ----------------製作圖表----------------

    # 設定圖表大小
    fig, ax1 = plt.subplots(figsize=(10, 4))

    # 繪製折線圖 (左Y軸-收盤價)
    ax1.plot(df["資料時間"], df["收盤價"], marker="o", markerfacecolor="white", color="red")  # 修改為空心點
    ax1.set_ylabel("收盤價", color="red", rotation=0, ha="left", fontsize=14)  # 修改為灰色，並調大字型大小

    # 創建第二個Y軸
    ax2 = ax1.twinx()

    # 繪製長條圖 (右Y軸-大股東持有比)
    ax2.bar(df["資料時間"], df["大股東持有比"], color="orange", alpha=0.5, width=0.4)
    ax2.set_ylabel("大股東(%)", color="orange", rotation=0, ha="right", fontsize=14)  # 修改為灰色，並調大字型大小

    # 修改刻度對應的值格式(小數點第二位)
    ax1.yaxis.set_major_formatter(FormatStrFormatter("%.2f"))
    ax2.yaxis.set_major_formatter(FormatStrFormatter("%.2f"))

    # 調整ax1的y軸範圍
    ax1.set_ylim(min(df["收盤價"]) * 0.9, max(df["收盤價"]) * 1.01)

    # 調整ax2的y軸範圍
    ax2.set_ylim(min(df["大股東持有比"]) * 0.9991, max(df["大股東持有比"]) * 1.002)

    # 計算資料的長度
    total_ticks = len(df)

    # 設定x軸標籤旋轉角度，修改刻度顏色及文字大小
    ax1.set_xticks(
        [df["資料時間"].iloc[0], df["資料時間"].iloc[total_ticks // 4], df["資料時間"].iloc[total_ticks // 2],
         df["資料時間"].iloc[total_ticks * 3 // 4], df["資料時間"].iloc[-1]])
    ax1.set_xticklabels(
        [df["資料時間"].iloc[0], df["資料時間"].iloc[total_ticks // 4], df["資料時間"].iloc[total_ticks // 2],
         df["資料時間"].iloc[total_ticks * 3 // 4], df["資料時間"].iloc[-1]], color="gray", fontsize=14)
    ax1.tick_params(axis="y", colors="red")  # 修改左Y軸刻度顏色
    ax2.tick_params(axis="y", colors="orange")  # 修改右Y軸刻度顏色

    # 調整y軸標籤位置
    ax1.yaxis.set_label_coords(0, 1.04)
    ax2.yaxis.set_label_coords(1, 1.10)

    # 設定中文字體
    plt.rcParams["font.sans-serif"] = ["Microsoft JhengHei"]
    plt.rcParams["axes.unicode_minus"] = False

    # 設定大標題
    plt.title("")

    # 設定邊框顏色
    ax1.spines["top"].set_color("gray")
    ax1.spines["bottom"].set_color("gray")
    ax1.spines["left"].set_color("gray")
    ax1.spines["right"].set_color("gray")
    ax2.spines["top"].set_color("gray")
    ax2.spines["bottom"].set_color("gray")
    ax2.spines["left"].set_color("gray")
    ax2.spines["right"].set_color("gray")

    # 調整X軸、兩個Y軸的標註文字大小
    ax1.xaxis.set_tick_params(labelsize=14)
    ax1.yaxis.set_tick_params(labelsize=14)
    ax2.yaxis.set_tick_params(labelsize=14)

    # 新增圖例
    ax1.legend(["收盤價"], loc="upper left", bbox_to_anchor=(0.29, 1.17), fontsize=14, frameon=False)  # 第一個圖例文字大小
    ax2.legend(["大股東(%)"], loc="upper right", bbox_to_anchor=(0.705, 1.17), fontsize=14, frameon=False)  # 第二個圖例文字大小

    # 加入格線(收盤價)
    ax1.grid(True, color="lightgray", linestyle="--", zorder=0)

    # 儲存圖片
    plt.savefig(str(stockNo) + "持股-圖表.png", bbox_inches="tight")

    # ----------------修改表格用資料----------------

    # 將【資料時間】格式改為時間(原為字串)
    df["資料時間"] = pd.to_datetime(df["資料時間"])

    # 創建新的DataFrame
    new_df = pd.DataFrame()

    # 設置新表格欄位【日期】
    new_df["日期"] = df["資料時間"]

    # 設置新表格欄位【大股東比】，將原始值除以100後改為百分比
    new_df["大股東比"] = (df["大股東持有比"] / 100).map("{:.2%}".format)

    # 設置新表格欄位【大股東變動】並取到小數點第二位，再加上百分號
    new_df["大股東變動"] = df["大股東持有比"].diff().round(2)

    # 如果數值大於0，在數值前加上"+"號；如果數值等於0，改為+0.00
    new_df["大股東變動"] = new_df["大股東變動"].apply(
        lambda x: "+" + str(x) if x > 0 else "+0.00" if x == 0.0 else str(x))

    # 加上百分號
    new_df["大股東變動"] = new_df["大股東變動"].astype(str) + "%"

    # 設置新表格欄位【大股東人數】
    new_df["大股東人數"] = df["人數"]

    # 將【大股東人數】轉換為整數型態
    new_df["大股東人數"] = new_df["大股東人數"].astype(int)

    # 將【大股東人數】格式改為千位分隔格式
    new_df["大股東人數"] = new_df["大股東人數"].apply(lambda x: "{:,.0f}".format(x))

    # 反轉新表格的順序
    new_df = new_df.iloc[::-1].reset_index(drop=True)

    # 將【日期】格式轉換為MM/DD
    new_df["日期"] = new_df["日期"].dt.strftime("%m/%d")

    # 只取前四筆
    new_df_front_four = new_df.head(4)

    # ----------------製作表格----------------

    # 繪製圖表
    fig, ax = plt.subplots(figsize=(13, 6))

    # 不顯示坐標軸
    ax.axis("off")

    # 將 DataFrame 轉換為表格圖
    table = ax.table(cellText=new_df_front_four.values, colLabels=new_df_front_four.columns, cellLoc="center",
                     loc="center")

    # 設定表格樣式
    table.auto_set_font_size(False)

    # 設定表格文字大小
    table.set_fontsize(10)

    # 調整表格大小(寬度、高度)
    table.scale(0.5, 2.1)

    # 遍歷所有的表格元素，設定背景顏色並移除框線
    for key, cell in table.get_celld().items():
        cell.set_linewidth(0)  # 移除框線
        cell.set_facecolor("white")  # 設定背景顏色

        # 設定標頭欄、第三行、第五行的背景顏色為灰色
        if key[0] == 0 or key[0] == 2 or key[0] == 4:
            cell.set_facecolor("#F6F6FC")

        # 將儲存格內容水平置中對齊
        cell.set_text_props(horizontalalignment="center")

    # 【大股東比】【大股東變動】顏色設定
    for row in range(1, 5):
        cell_index_1 = table._cells[(row, 1)]  # 【大股東比】列索列為1
        cell_index_2 = table._cells[(row, 2)]  # 【大股東變動】列索列為2

        cell_index_2_text = cell_index_2.get_text().get_text().replace(",", "").replace("%", "")
        # 【大股東變動】的值
        cell_index_2_value = float(cell_index_2_text)

        # 如果【大股東變動】的值大於0
        if cell_index_2_value > 0:
            # 【大股東比】【大股東變動】值改為紅色
            cell_index_1.set_text_props(color="red")
            cell_index_2.set_text_props(color="red")

        elif cell_index_2_value == 0:
            # 【大股東比】【大股東變動】值改為紅色
            cell_index_1.set_text_props(color="black")
            cell_index_2.set_text_props(color="black")

        # 如果【大股東變動】的值小於0
        else:
            # 【大股東比】【大股東變動】值改為綠色
            cell_index_1.set_text_props(color="green")
            cell_index_2.set_text_props(color="green")

    # 【大股東人數】顏色設定
    for i in range(1, 5):
        cell = table._cells[(i, 3)]
        # 如果大於前一個值，就設為紅色
        if new_df["大股東人數"].iloc[i - 1] > new_df["大股東人數"].iloc[i]:
            cell.set_text_props(color="red")
        # 如果小於前一個值，就設為綠色
        else:
            cell.set_text_props(color="green")

    # 儲存圖片
    plt.savefig(f"{stockNo}持股-表格.png")

    # ----------------------------------表格裁切空白處----------------------------------

    # 讀取圖片
    image_path = str(stockNo) + "持股-表格.png"
    image = Image.open(image_path)

    # 獲取圖片的寬度和高度
    width, height = image.size

    # 裁剪圖片
    left = 420  # 裁剪框左上角的x座標
    top = 205  # 裁剪框左上角的y座標
    right = 915  # 裁剪框右下角的x座標
    bottom = 405  # 裁剪框右下角的y座標
    cropped_image = image.crop((left, top, right, bottom))

    # 儲存裁剪後的圖片
    cropped_image_path = str(stockNo) + "持股-表格-裁剪.png"
    cropped_image.save(cropped_image_path)

    # ----------------依比例調整表格----------------

    # --------取得【圖表寬度】--------

    # 讀取圖片
    image_path = str(stockNo) + "持股-圖表.png"
    image = Image.open(image_path)

    # 取得圖片的寬度
    width_target = image.width

    # -------調整【表格寬度】與【圖表寬度】相同-------

    # 讀取圖片
    image_path = str(stockNo) + "持股-表格-裁剪.png"
    image = Image.open(image_path)

    # 獲取原始圖片的寬度和高度
    width, height = image.size

    # 計算縮放比例
    scale_factor = width_target / width

    # 計算縮放後的寬度和高度
    new_width = int(width * scale_factor)
    new_height = int(height * scale_factor)

    # 縮放圖片
    resized_image = image.resize((new_width, new_height))

    # 輸出結果
    resized_image_path = str(stockNo) + "持股-表格-裁剪-調寬度.png"
    resized_image.save(resized_image_path)

    # ----------------------------------合併PNG【圖表】+【表格】----------------------------------

    # 讀取第一張PNG圖片
    image1_path = str(stockNo) + "持股-圖表.png"
    image1 = Image.open(image1_path)

    # 讀取第二張PNG圖片
    image2_path = str(stockNo) + "持股-表格-裁剪-調寬度.png"
    image2 = Image.open(image2_path)

    # 獲取第一張圖片的寬度和高度
    width1, height1 = image1.size

    # 獲取第二張圖片的寬度和高度
    width2, height2 = image2.size

    # 確定合併後圖片的寬度和高度
    merged_width = max(width1, width2)
    merged_height = height1 + height2

    # 創建一個新的空白圖片，尺寸為合併後的寬度和高度
    merged_image = Image.new("RGB", (merged_width, merged_height))

    # 將第一張圖片黏貼到合併圖片的上半部分
    merged_image.paste(image1, (0, 0))

    # 將第二張圖片黏貼到合併圖片的下半部分
    merged_image.paste(image2, (0, height1))

    # 儲存合併後的圖片
    merged_image_path = str(stockNo) + "持股-圖表+表格.png"
    merged_image.save(merged_image_path)

    # ---------------------製作【持股備註-圖片】---------------------

    # 創建空白底圖
    image_width = width_target
    image_height = 60
    background_color = (255, 255, 255)  # 白色背景
    image = Image.new("RGB", (image_width, image_height), background_color)
    draw = ImageDraw.Draw(image)

    # 設置文字內容
    text = "條件：大股東張數大於1000    "
    text_color = (74, 74, 74)  # "#4a4a4a"的RGB值

    # 使用字體
    font_size = 26
    font_path = "msjh.ttc"
    font = ImageFont.truetype(font_path, font_size)

    # 將內容拆分為多行
    wrapper = textwrap.TextWrapper(width=image_width, break_long_words=False)
    lines = wrapper.wrap(text)

    # 計算內容大小
    max_width = max(font.getbbox(line)[2] for line in lines)
    text_height = font.getmask(lines[0]).getbbox()[3] * len(lines)  # 假設所有行高度相同

    # 計算內容位置（靠右對齊）
    text_x = image_width - max_width - 50  # 內容右上角的X座標
    text_y = (image_height - text_height) // 3  # 將內容對齊(垂直置中)

    # 在圖像上繪製內容
    for line in lines:
        draw.text((text_x, text_y), line, fill=text_color, font=font)
        text_y += text_height

    # 保存图像
    image.save("持股-備註文字.png")

    # ----------------------------------合併PNG【備註文字】+【圖表+表格】----------------------------------

    # 讀取第一張PNG圖片
    image1_path = "持股-備註文字.png"
    image1 = Image.open(image1_path)

    # 讀取第二張PNG圖片
    image2_path = str(stockNo) + "持股-圖表+表格.png"
    image2 = Image.open(image2_path)

    # 獲取第一張圖片的寬度和高度
    width1, height1 = image1.size

    # 獲取第二張圖片的寬度和高度
    width2, height2 = image2.size

    # 確定合併後圖片的寬度和高度
    merged_width = max(width1, width2)
    merged_height = height1 + height2

    # 創建一個新的空白圖片，尺寸為合併後的寬度和高度
    merged_image = Image.new("RGB", (merged_width, merged_height))

    # 將第一張圖片黏貼到合併圖片的上半部分
    merged_image.paste(image1, (0, 0))

    # 將第二張圖片黏貼到合併圖片的下半部分
    merged_image.paste(image2, (0, height1))

    # 儲存合併後的圖片
    merged_image_path = filename
    merged_image.save(merged_image_path)

    # 印出確認文字
    print("已完成 " + filename)

# 持股-bubble
def shareholder(stockNo):

    # 製作【股票資訊.csv】
    stock_info_csv(stockNo)

    # ------------------------------

    # 取得今天日期
    now = datetime.now()

    # 設定當前日期格式
    now_date = now.strftime("%m%d")

    # 取得當前日期是星期幾(星期一為0，星期日為6)
    current_weekday = now.weekday()

    # 計算距離上個星期五的天數
    days_to_friday = (current_weekday - 4) % 7

    # 計算上個星期五的日期
    last_friday_date = (now - timedelta(days=days_to_friday)).strftime("%m%d")

    # 指定檔名
    filename = str(stockNo) + f"持股_{last_friday_date}.png"

    # 如果指定檔不存在
    if not os.path.isfile(filename):

        # 製作【持股-圖表+表格】
        shareholder_png(stockNo, filename)

    # 如果指定檔存在
    else:
        # 印出確認文字
        print("無須更新 " + filename)

    # ------------------圖片上傳到Imgur------------------

    # 讀取指定檔案，取得IMGUR_ID
    with open("env.json") as f:
        env = json.load(f)

    # 讀入IMGUR_ID
    CLIENT_ID = env["YOUR_IMGUR_ID"]
    # 指定圖片路徑
    PATH = filename
    # 指定圖片標題
    title = filename

    # 將圖片上傳到IMGUR
    im = pyimgur.Imgur(CLIENT_ID)
    # 帶入圖片路徑及標題
    uploaded_image = im.upload_image(PATH, title=title)

    # ------------------------------讀取stock_info.csv------------------------------

    # 指定檔案路徑
    file_name = f"./{stockNo}_info.csv"

    # 讀取指定檔案
    df = pd.read_csv(file_name)

    # 列出欄位下的值
    update_date = df["資料日期"].tolist()[0]
    update_date_time = df["更新時間"].tolist()[0]
    stock_no = df["股票代碼"].tolist()[0]
    stock_name = df["股票名稱"].tolist()[0]
    listing_cabinet = df["上市櫃"].tolist()[0]
    industry = df["產業別"].tolist()[0]

    close_price = df["收盤價"].tolist()[0]
    # 如果值有小數點，就改為小數點後兩位(字串)
    if len(str(close_price).split(".")) == 2:
        close_price = "{:.2f}".format(close_price)
    # 如果值為整數，就維持(整數)
    else:
        close_price = close_price

    price_change = df["漲跌價"].tolist()[0]
    price_change_rate = df["漲跌幅"].tolist()[0]

    open_price = df["開盤"].tolist()[0]
    # 如果值有小數點，就改為小數點後兩位(字串)
    if len(str(open_price).split(".")) == 2:
        open_price = "{:.2f}".format(open_price)
    # 如果值為整數，就維持(整數)
    else:
        open_price = open_price

    high_price = df["最高"].tolist()[0]
    # 如果值有小數點，就改為小數點後兩位(字串)
    if len(str(high_price).split(".")) == 2:
        high_price = "{:.2f}".format(high_price)
    # 如果值為整數，就維持(整數)
    else:
        high_price = high_price

    low_price = df["最低"].tolist()[0]
    # 如果值有小數點，就改為小數點後兩位(字串)
    if len(str(low_price).split(".")) == 2:
        low_price = "{:.2f}".format(low_price)
    # 如果值為整數，就維持(整數)
    else:
        low_price = low_price

    deal_amount = df["成交量"].tolist()[0]

    # 建立FlexSendMessage
    message = FlexSendMessage(
        alt_text = str(stock_name) + str(stock_no) + " 大股東週報",
        contents = {
            "type": "bubble",
            "size": "giga",
            "header": {
                "type": "box",
                "layout": "vertical",
                "contents": [
                    {
                        "type": "box",
                        "layout": "horizontal",
                        "contents": [
                            {
                                "type": "text",
                                "text": "⇧" + " 加好友",
                                "color": "#d90b0b",
                                "weight": "bold",
                                "size": "xl",
                                "action": {
                                    "type": "uri",
                                    "label": "action",
                                    "uri": "https://line.me/R/ti/p/%40636hvuxg"
                                }
                            }
                        ]
                    },
                    {
                        "type": "box",
                        "layout": "horizontal",
                        "contents": [
                            {
                                "type": "text",
                                "text": str(stock_name),
                                "size": "xxl",
                                "weight": "bold",
                                "flex": 0
                            },
                            {
                                "type": "text",
                                "text": str(stock_no),
                                "gravity": "bottom",
                                "flex": 2
                            },
                            {
                                "type": "filler"
                            },
                            {
                                "type": "text",
                                "text": str(close_price),
                                "size": "xxl",
                                "flex": 3,
                                "align": "end"
                            },
                            {
                                "type": "box",
                                "layout": "vertical",
                                "contents": [
                                    {
                                        "type": "text",
                                        "text": str(price_change),
                                        "align": "center",
                                        "size": "xs"
                                    },
                                    {
                                        "type": "text",
                                        "text": str(price_change_rate),
                                        "align": "center",
                                        "size": "xs"
                                    }
                                ],
                                "flex": 0,
                                "margin": "sm"
                            }
                        ]
                    },
                    {
                        "type": "box",
                        "layout": "baseline",
                        "contents": [
                            {
                                "type": "text",
                                "text": str(listing_cabinet) + " / " + str(industry),
                                "flex": 1,
                                "size": "xs",
                                "color": "#C5C4C7"
                            },
                            {
                                "type": "text",
                                "text": "股價更新時間：" + str(update_date_time),
                                "size": "xs",
                                "margin": "none",
                                "flex": 2,
                                "align": "end",
                                "color": "#C5C4C7"
                            }
                        ]
                    }
                ],
                "backgroundColor": "#F6F6FC"
            },
            "hero": {
                "type": "image",
                "url": uploaded_image.link,
                "size": "full",
                "aspectRatio": "19:18",
                "aspectMode": "fit",
                "action": {
                    "type": "uri",
                    "label": "action",
                    "uri": "https://www.nstock.tw/stock_info?status=1&stock_id=" + str(stockNo) + "&utm_source=line&utm_medium=line_bot"
                }
            },
            "footer": {
                "type": "box",
                "layout": "horizontal",
                "spacing": "none",
                "contents": [
                    {
                        "type": "text",
                        "text": "選項：",
                        "size": "xs",
                        "color": "#787878"
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "即時",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": "查" + str(stock_name)
                        }
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "K線",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "日K"
                        }
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "法人",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "法人"
                        }
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "EPS",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "EPS"
                        }
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "營收",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "營收"
                        }
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "股利",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "股利"
                        }
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "持股",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "持股"
                        },
                        "color": "#d11111",
                        "weight": "bold"
                    }
                ]
            }
        }
    )

    return message

# ----------------------------------------------------------------

# 製作【營收-表格】
def revenue_png(stockNo):
    # 欲爬取網站
    # https://histock.tw/stock/3038/%E8%B2%A1%E5%8B%99%E5%A0%B1%E8%A1%A8
    url = "https://histock.tw/stock/" + str(stockNo) + "/%E8%B2%A1%E5%8B%99%E5%A0%B1%E8%A1%A8"

    # 提供額外的請求資訊，確保請求可以正確地發送並獲得回應
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36 Edg/114.0.1823.51"
    }

    res = requests.get(url, headers=headers)

    # 因中文字顯示可能變為亂碼，需編碼設定
    res.encoding = "utf-8"
    soup = BeautifulSoup(res.text, "html.parser")

    # ------------取得【年份-月-營收】------------

    rows = soup.find_all("tr")[1:]

    # 逐一取出數據並加到清單中
    data = []
    for row in rows:
        cells = row.find_all("td")
        if len(cells) > 0:
            month = cells[0].text.strip()
            value = cells[1].text.strip()
            data.append([month, value])

    # 製作DataFrame
    df = pd.DataFrame(data, columns=["年度/月份", "營收(億)"])

    # 去除欄位【營收(億)】的逗號，並轉換為數值型別
    df["營收(億)"] = pd.to_numeric(df["營收(億)"].str.replace(",", ""))

    # 欄位【營收(億)】除以100000
    df["營收(億)"] = df["營收(億)"] / 100000

    # 欄位【營收(億)】保留小數點後二位
    df["營收(億)"] = df["營收(億)"].round(2)

    # 欄位【營收(億)】格式修改為每三位加一逗號
    df["營收(億)"] = df["營收(億)"].apply(lambda x: "{:,.2f}".format(x))

    # 顛倒DataFrame的行順序
    df = df.iloc[::-1].reset_index(drop=True)

    # 從欄位【年度/月份】的後2個字，新增欄位【月】
    df["月"] = df["年度/月份"].str[-2:]

    # 從欄位【年度/月份】的前4個字，新增欄位【年度】
    df["年度"] = df["年度/月份"].str[:4]

    # 調整欄位的順序
    df = df[["年度", "月", "營收(億)"]]

    # ----------------製作欄位【月】----------------

    # 設定目標年份=當前年份-1
    target_year = datetime.now().year - 1

    # 根據篩選條件(欄位【年度】為前1年)，並取出欄位【月】，重新創建新DataFrame
    month_df = df[df["年度"] == str(target_year)][["月"]].reset_index(drop=True)

    # 新增第13行，放入目標內容
    month_df.loc[12] = "累計"

    # ----------------製作欄位【前2年之營收】----------------

    # 設定目標年份=當前年份-2
    target_year = datetime.now().year - 2

    # 根據篩選條件(欄位【年度】為前2年)，並取出欄位【營收(億)】，重新創建新DataFrame
    now_2_df = df[df["年度"] == str(target_year)][["營收(億)"]].reset_index(drop=True)

    # 修改欄位【營收(億)】名稱
    now_2_df = now_2_df.rename(columns={"營收(億)": f"{target_year} 營收(億)"})

    # -------------取得總營收(前2年)-------------

    # 取出所有值，並製作成清單
    revenue_now = now_2_df[now_2_df.columns[0]].tolist()

    # 將字串轉換成浮點數，並刪除-
    revenue_now_float = [float(value.replace(",", "")) for value in revenue_now if value != "-"]

    # 將所有值相加
    total_revenue_now = sum(revenue_now_float)

    # 取到小數點第二位，並加上逗號
    total_revenue_now = "{:,.2f}".format(total_revenue_now)

    # 新增第13行，放入目標內容(該年營收總額)
    now_2_df.loc[12] = total_revenue_now

    # ----------------製作欄位【前1年之營收】----------------

    # 設定目標年份=當前年份-1
    target_year = datetime.now().year - 1

    # 根據篩選條件(欄位【年度】為前1年)，並取出欄位【營收(億)】，重新創建新DataFrame
    now_1_df = df[df["年度"] == str(target_year)][["營收(億)"]].reset_index(drop=True)

    # 修改欄位【營收(億)】名稱
    now_1_df = now_1_df.rename(columns={"營收(億)": f"{target_year} 營收(億)"})

    # -------------取得總營收(前1年)-------------

    # 取出所有值，並製作成清單
    revenue_now = now_1_df[now_1_df.columns[0]].tolist()

    # 將字串轉換成浮點數，並刪除-
    revenue_now_float = [float(value.replace(",", "")) for value in revenue_now if value != "-"]

    # 將所有值相加
    total_revenue_now = sum(revenue_now_float)

    # 取到小數點第二位
    total_revenue_now = "{:,.2f}".format(total_revenue_now)

    # 新增第13行，放入目標內容(該年營收總額)
    now_1_df.loc[12] = total_revenue_now

    # ----------------製作欄位【今年之營收】----------------

    # 設定目標年份=當前年份
    target_year = datetime.now().year

    # 根據篩選條件(欄位【年度】為今年)，並取出欄位【營收(億)】，重新創建新DataFrame
    now_df = df[df["年度"] == str(target_year)][["營收(億)"]].reset_index(drop=True)

    # 修改欄位【營收(億)】名稱
    now_df = now_df.rename(columns={"營收(億)": f"{target_year} 營收(億)"})

    # 增加行數到12行，並填充內容為-
    now_df = now_df.reindex(range(12)).fillna("-")

    # -------------取得總營收(今年)-------------

    # 取出所有值，並製作成清單
    revenue_now = now_df[now_df.columns[0]].tolist()

    # 將字串轉換成浮點數，並刪除-
    revenue_now_float = [float(value.replace(",", "")) for value in revenue_now if value != "-"]

    # 將所有值相加
    total_revenue_now = sum(revenue_now_float)

    # 取到小數點第二位，並加上逗號
    total_revenue_now = "{:,.2f}".format(total_revenue_now)

    # 新增第13行，放入目標內容(該年營收總額)
    now_df.loc[12] = total_revenue_now

    # -------------合併df-------------

    # 將3個DataFrame按照從左到右的順序進行合併
    df = pd.concat([month_df, now_df, now_1_df, now_2_df], axis=1)

    # 新增欄位【年增(%)】
    df.insert(2, "年增(%)", "")

    # ----------------------------------計算【年增(%)】-----------------------------------------

    # -------------計算【今年的營收】-------------

    # 取出所有值，並製作成清單
    revenue_now = df[df.columns[1]].tolist()

    # 將字串轉換成浮點數，並刪除-
    revenue_now_float = [float(value.replace(",", "")) if value != "-" else value for value in revenue_now]

    # 將累計值改為-
    revenue_now_float[-1] = "-"

    # -------------計算【前1年的營收】-------------

    # 取出所有值，並製作成清單
    revenue_now = df[df.columns[3]].tolist()

    # 將字串轉換成浮點數，並刪除-
    revenue_now_float_1 = [float(value.replace(",", "")) if value != "-" else value for value in revenue_now]

    # -------------比較【今年的營收】與【前1年的營收】-------------

    # 如果今年營收為-，前1年的營收改為-
    for i in range(len(revenue_now_float)):
        if revenue_now_float[i] == "-":
            revenue_now_float_1[i] = "-"

    # -------------計算【今年的營收】與【前1年的營收】之年增(%)-------------

    # 建立一清單
    result = []

    # 計算年增(%)
    for val1, val2 in zip(revenue_now_float, revenue_now_float_1):
        # 如果其一為-，就新增-
        if val1 == "-" or val2 == "-":
            result.append("-")
        # 如果都不為-，進行計算
        else:
            result.append((val1 - val2) / val2)

    # -------------將清單轉為百分比格式-------------

    # 建立一清單
    formatted_result = []

    # 進行格式調整
    for val in result:
        if val == "-":
            formatted_result.append(val)
        else:
            formatted_result.append(f"{val:.1%}")

    # 排除掉-，只保留有數值的部分
    filtered_result = [val for val in formatted_result if val != "-"]

    # ---------取得【當年營收總額】---------

    # 刪除-
    revenue_now_float = [value for value in revenue_now_float if value != "-"]

    # 將數字加總
    total_revenue = sum(revenue_now_float)

    # ---------取得【前1年營收總額】---------

    # 刪除-
    revenue_now_float_1 = [value for value in revenue_now_float_1 if value != "-"]

    # 將數字加總
    total_revenue_1 = sum(revenue_now_float_1)

    # ---------取得【當年營收增額】---------

    # 進行計算
    total_revenue_average = (total_revenue - total_revenue_1) / total_revenue_1

    # 乘以100
    total_revenue_average_percent = total_revenue_average * 100

    # 格式調整
    total_revenue_average = "{:.2f}%".format(total_revenue_average_percent)

    # ----------製作欄位【年增(%)】下的值----------

    # 將清單擴展，共有13個值，空的值為-，最後一個值為當年營收增額
    extended_result = filtered_result + ["-"] * (13 - len(filtered_result) - 1) + [total_revenue_average]

    # 在欄位【年增(%)】插入值
    df["年增(%)"] = extended_result

    # -----------------------製作表格-----------------------

    # 設定中文字體
    plt.rcParams["font.sans-serif"] = ["Microsoft JhengHei"]
    plt.rcParams["axes.unicode_minus"] = False

    # 繪製圖表
    fig, ax = plt.subplots()

    # 隱藏座標軸
    ax.axis("off")

    # 創建表格並填入資料
    table = ax.table(cellText=df.values,
                     colLabels=df.columns,
                     cellLoc="right",
                     loc="center")

    # 指定需要居中對齊的列索引(第一欄)
    center_align_columns = [0]

    # 設置指定欄位的對齊方式為居中
    for row in range(14):
        for column in center_align_columns:
            cell = table._cells[(row, column)]
            cell.set_text_props(ha="center")

    # 設定表頭列的背景顏色
    for i in range(len(df.columns)):
        cell = table[0, i]
        cell.set_facecolor("#F6F6FC")

    # 設定表格框線顏色為灰色
    for key, cell in table.get_celld().items():
        cell.set_linewidth(0.5)
        cell.set_edgecolor("#c3c6c9")

    # 設定表格的樣式
    table.auto_set_font_size(False)

    # 文字大小
    table.set_fontsize(11)

    # 表格寬度、高度
    table.scale(2, 2)

    # 自動調整表格寬度
    table.auto_set_column_width(list(range(len(df.columns))))

    # 根据条件设置文字颜色
    for row in range(1, 14):
        cell = table._cells[(row, 2)]  # "年增(%)"字段所在的列索引为4
        cell_text = cell.get_text().get_text().replace(",", "").replace("%", "")
        if cell_text != "-":
            cell_value = float(cell_text)
            if cell_value > 0:
                cell.set_text_props(color="red")
            else:
                cell.set_text_props(color="green")

    # 儲存圖片
    plt.savefig(str(stockNo) + "營收-待調整.png", bbox_inches="tight", pad_inches=0.05)

    # ----------------------------------表格裁切空白處----------------------------------

    # 讀取圖片
    image_path = str(stockNo) + "營收-待調整.png"
    image = Image.open(image_path)

    # 獲取圖片的寬度和高度
    width, height = image.size

    # 裁剪圖片
    left = 26  # 裁剪框左上角的x座標
    top = 0  # 裁剪框左上角的y座標
    right = 482  # 裁剪框右下角的x座標
    bottom = 478  # 裁剪框右下角的y座標
    cropped_image = image.crop((left, top, right, bottom))

    # 取得當前時間
    now = datetime.now()

    # 提取當前時間的年份和月份
    month = now.month

    # 計算前一個月份
    # 如果當前月份為1月，則回推到前一年的12月
    if month == 1:
        previous_month = 12
    else:
        previous_month = month - 1

    # 儲存裁剪後的圖片
    cropped_image_path = str(stockNo) + "營收_" + str(previous_month) + "月.png"
    cropped_image.save(cropped_image_path)

# 確認是否有【營收-表格】
def check_file_revenue_png(stockNo):

    # 取得當前時間
    now = datetime.now()

    # 提取當前時間的年份和月份
    month = now.month

    # 計算前一個月份
    # 如果當前月份為1月，則回推到前一年的12月
    if month == 1:
        previous_month = 12
    else:
        previous_month = month - 1

    # 指定檔名
    filename = str(stockNo) + "營收_" + str(previous_month) + "月.png"

    # 如果指定檔不存在
    if not os.path.isfile(filename):
        # 製作【股票名稱/代號.csv】
        revenue_png(stockNo)
        # 印出確認文字
        print("已完成 " + str(stockNo) + "營收_" + str(previous_month) + "月.png")

    # 如果指定檔存在
    else:
        # 印出確認文字
        print("無須更新 " + str(stockNo) + "營收_" + str(previous_month) + "月.png")

# 營收-bubble
def revenue(stockNo):

    # 製作【股票資訊.csv】
    stock_info_csv(stockNo)

    # 確認【營收表格】是否存在
    check_file_revenue_png(stockNo)

    # 取得當前時間
    now = datetime.now()

    # 提取當前時間的年份和月份
    month = now.month

    # 計算前一個月份
    # 如果當前月份為1月，則回推到前一年的12月
    if month == 1:
        previous_month = 12
    else:
        previous_month = month - 1

    # ------------------圖片上傳到Imgur------------------

    # 讀取指定檔案，取得IMGUR_ID
    with open("env.json") as f:
        env = json.load(f)

    # 讀入IMGUR_ID
    CLIENT_ID = env["YOUR_IMGUR_ID"]
    # 指定圖片路徑
    PATH = str(stockNo) + "營收_" + str(previous_month) + "月.png"  # A Filepath to an image on your computer
    # 指定圖片標題
    title = str(stockNo) + "營收_" + str(previous_month) + "月.png"

    # 將圖片上傳到IMGUR
    im = pyimgur.Imgur(CLIENT_ID)
    # 帶入圖片路徑及標題
    uploaded_image = im.upload_image(PATH, title=title)

    # ------------------------------讀取stock_info.csv------------------------------

    # 指定檔案路徑
    file_name = f"./{stockNo}_info.csv"

    # 讀取指定檔案
    df = pd.read_csv(file_name)

    # 列出欄位下的值
    update_date = df["資料日期"].tolist()[0]
    update_date_time = df["更新時間"].tolist()[0]
    stock_no = df["股票代碼"].tolist()[0]
    stock_name = df["股票名稱"].tolist()[0]
    listing_cabinet = df["上市櫃"].tolist()[0]
    industry = df["產業別"].tolist()[0]

    close_price = df["收盤價"].tolist()[0]
    # 如果值有小數點，就改為小數點後兩位(字串)
    if len(str(close_price).split(".")) == 2:
        close_price = "{:.2f}".format(close_price)
    # 如果值為整數，就維持(整數)
    else:
        close_price = close_price

    price_change = df["漲跌價"].tolist()[0]
    price_change_rate = df["漲跌幅"].tolist()[0]

    open_price = df["開盤"].tolist()[0]
    # 如果值有小數點，就改為小數點後兩位(字串)
    if len(str(open_price).split(".")) == 2:
        open_price = "{:.2f}".format(open_price)
    # 如果值為整數，就維持(整數)
    else:
        open_price = open_price

    high_price = df["最高"].tolist()[0]
    # 如果值有小數點，就改為小數點後兩位(字串)
    if len(str(high_price).split(".")) == 2:
        high_price = "{:.2f}".format(high_price)
    # 如果值為整數，就維持(整數)
    else:
        high_price = high_price

    low_price = df["最低"].tolist()[0]
    # 如果值有小數點，就改為小數點後兩位(字串)
    if len(str(low_price).split(".")) == 2:
        low_price = "{:.2f}".format(low_price)
    # 如果值為整數，就維持(整數)
    else:
        low_price = low_price

    deal_amount = df["成交量"].tolist()[0]

    # 建立FlexSendMessage
    message = FlexSendMessage(
        alt_text = str(stock_name) + str(stock_no) + " 月營收",
        contents = {
            "type": "bubble",
            "size": "giga",
            "header": {
                "type": "box",
                "layout": "vertical",
                "contents": [
                    {
                        "type": "box",
                        "layout": "horizontal",
                        "contents": [
                            {
                                "type": "text",
                                "text": "⇧" + " 加好友",
                                "color": "#d90b0b",
                                "weight": "bold",
                                "size": "xl",
                                "action": {
                                    "type": "uri",
                                    "label": "action",
                                    "uri": "https://line.me/R/ti/p/%40636hvuxg"
                                }
                            }
                        ]
                    },
                    {
                        "type": "box",
                        "layout": "horizontal",
                        "contents": [
                            {
                                "type": "text",
                                "text": str(stock_name),
                                "size": "xxl",
                                "weight": "bold",
                                "flex": 0
                            },
                            {
                                "type": "text",
                                "text": str(stock_no),
                                "gravity": "bottom",
                                "flex": 2
                            },
                            {
                                "type": "filler"
                            },
                            {
                                "type": "text",
                                "text": str(close_price),
                                "size": "xxl",
                                "flex": 3,
                                "align": "end"
                            },
                            {
                                "type": "box",
                                "layout": "vertical",
                                "contents": [
                                    {
                                        "type": "text",
                                        "text": str(price_change),
                                        "align": "center",
                                        "size": "xs"
                                    },
                                    {
                                        "type": "text",
                                        "text": str(price_change_rate),
                                        "align": "center",
                                        "size": "xs"
                                    }
                                ],
                                "flex": 0,
                                "margin": "sm"
                            }
                        ]
                    },
                    {
                        "type": "box",
                        "layout": "baseline",
                        "contents": [
                            {
                                "type": "text",
                                "text": str(listing_cabinet) + " / " + str(industry),
                                "flex": 1,
                                "size": "xs",
                                "color": "#C5C4C7"
                            },
                            {
                                "type": "text",
                                "text": "股價更新時間：" + str(update_date_time),
                                "size": "xs",
                                "margin": "none",
                                "flex": 2,
                                "align": "end",
                                "color": "#C5C4C7"
                            }
                        ]
                    }
                ],
                "backgroundColor": "#F6F6FC"
            },
            "hero": {
                "type": "image",
                "url": uploaded_image.link,
                "size": "full",
                "aspectRatio": "20:21",
                "aspectMode": "fit"
            },
            "footer": {
                "type": "box",
                "layout": "horizontal",
                "spacing": "none",
                "contents": [
                    {
                        "type": "text",
                        "text": "選項：",
                        "size": "xs",
                        "color": "#787878"
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "即時",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": "查" + str(stock_name)
                        }
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "K線",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "日K"
                        }
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "法人",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "法人"
                        }
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "EPS",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "EPS"
                        }
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "營收",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "營收"
                        },
                        "color": "#d11111",
                        "weight": "bold"
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "股利",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "股利"
                        }
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "持股",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "持股"
                        }
                    }
                ]
            }
        }
    )

    return message

# ----------------------------------------------------------------

# 製作【日K所需資料】
def day_candlestick_data_csv(stockNo):

    # 抓取月數
    month_num = 4

    for i in range(month_num):
        # 使用twstock模組
        stock = twstock.Stock(str(stockNo))

        # 取得今天時間
        now_time = datetime.now()

        # 取得今天日期
        today_date = now_time.date()

        # 計算目標日期
        target_date = today_date - relativedelta(months=i)

        # 執行資料抓取
        price = stock.fetch(target_date.year, target_date.month)

        # 確認用
        print("爬取資料中，請稍後...")
        print("-----------------------------------")

        import time

        # 在請求之後等待6秒鐘
        time.sleep(6)

        # 確認用
        print(f"已完成【{stockNo}-前{i}個月-{target_date.year}{target_date.month:02d}】的資料爬取")

        # 設定資料表的表頭：日期 總成交股數 總成交金額(Volume) 開 高 低 收 漲跌幅 成交量
        name_attribute = ["Date", "Capacity", "Turnover", "Open", "High", "Low", "Close", "Change", "Transcation"]

        # 指定要放入資料表的資料來源，將清單轉換成資料表格式
        df = pd.DataFrame(columns=name_attribute, data=price)

        # csv檔案路徑
        file_name = f"./{stockNo}-{target_date.year}{target_date.month:02d}.csv"

        # 將資料表保存成.csv檔案
        df.to_csv(file_name)

        # 確認用
        print(f"已完成【{stockNo}-前{i}個月-{target_date.year}{target_date.month:02d}】的csv檔")
        print("-----------------------------------")

    # --------製作【檔案名稱】--------
    # 當月
    file_name_0 = f"{stockNo}-{today_date.year}{today_date.month:02d}.csv"
    # 前1個月
    file_name_1 = f"{stockNo}-{(today_date - relativedelta(months=1)).year}{(today_date - relativedelta(months=1)).month:02d}.csv"
    # 前2個月
    file_name_2 = f"{stockNo}-{(today_date - relativedelta(months=2)).year}{(today_date - relativedelta(months=2)).month:02d}.csv"
    # 前3個月
    file_name_3 = f"{stockNo}-{(today_date - relativedelta(months=3)).year}{(today_date - relativedelta(months=3)).month:02d}.csv"

    # 新增一空的DataFrame已儲存數據
    merged_data = pd.DataFrame()

    # 合併檔案
    for file_name in [file_name_3, file_name_2, file_name_1, file_name_0]:
        # 檔案路徑
        file_path = os.path.join(".", file_name)
        df = pd.read_csv(file_path)
        merged_data = pd.concat([merged_data, df], ignore_index=True)

    # 儲存.csv
    merged_file_path = os.path.join(".",
                                    f"{stockNo}-{(today_date - relativedelta(months=3)).year}{(today_date - relativedelta(months=3)).month:02d}-{today_date.year}{today_date.month:02d}.csv")
    merged_data.to_csv(merged_file_path, index=False)

    print(
        f"已完成 {stockNo}-{(today_date - relativedelta(months=3)).year}{(today_date - relativedelta(months=3)).month:02d}-{today_date.year}{today_date.month:02d}.csv")

# 製作【日K線圖】
def day_candlestick_png(stockNo):

    # 取得今天時間
    now_time = datetime.now()

    # 取得今天日期
    current_date = now_time.date()

    # 讀取.csv
    df = pd.read_csv(
        f"./{stockNo}-{(current_date - relativedelta(months=3)).year}0{(current_date - relativedelta(months=3)).month}-{current_date.year}0{current_date.month}.csv",
        parse_dates=True, index_col=1)

    # ----------------製作圖表----------------

    # 調整表頭(Turnover>Volume)，因mplfinance模組中對於交易量的辨認是Volume該字
    df.rename(columns={"Turnover": "Volume"}, inplace=True)

    # 由於mplfinance內建的漲/跌標記顏色是美國的版本(綠漲紅跌)，
    # 所以要先使用mplfinance中自訂圖表外觀功能
    # mpf.make_marketcolors()將漲/跌顏色改為台灣版本(紅漲綠跌)，
    # 接著再將這個設定以mpf.make_mpf_style()功能保存為自訂的外觀

    mc = mpf.make_marketcolors(up="r", down="g", inherit=True)
    s = mpf.make_mpf_style(base_mpf_style="yahoo", marketcolors=mc)

    # 圖表設定
    # 因使用candlestick類型的圖表設定，無法直接在該設定中調整長條度的粗度
    kwargs = dict(type="candle",  # 指定圖表的類型
                  mav=(5, 20, 60),  # 均線的設定
                  volume=True,  # 指示是否顯示交易量的設定
                  figratio=(10, 6),  # 指定圖表的寬高比
                  figscale=1.5,  # 指定圖表的縮放比例
                  title="",  # 圖表的標題
                  style=s,  # 套用的圖表外觀風格
                  xrotation=0,  # x軸標籤的旋轉角度
                  datetime_format="%Y/%m/%d",  # 日期時間的格式設定
                  ylabel=""  # Y軸標籤
                  )

    # 選擇df資料表為資料來源，帶入kwargs參數，畫出目標股票的走勢圖
    fig, axlist = mpf.plot(df, **kwargs, returnfig=True)

    # --------------儲存圖片--------------

    # 儲存路徑
    save_path = f"./{stockNo}-{(current_date - relativedelta(months=3)).year}0{(current_date - relativedelta(months=3)).month}-{current_date.year}0{current_date.month}.png"

    # 儲存圖片
    fig.savefig(save_path, bbox_inches="tight")
    print(f"已完成 {stockNo}日K-圖表.png")

    # ------------取得【5MA, 20MA, 60MA】------------

    # 計算5日移動平均線（5MA）、20日移動平均線（20MA）、60日移動平均線（60MA）
    df["5MA"] = df["Close"].rolling(window=5).mean()
    df["20MA"] = df["Close"].rolling(window=20).mean()
    df["60MA"] = df["Close"].rolling(window=60).mean()

    # 列出最新的5MA、20MA和60MA數值（取到小數點第二位）
    latest_5MA = round(df["5MA"].iloc[-1], 2)
    latest_20MA = round(df["20MA"].iloc[-1], 2)
    latest_60MA = round(df["60MA"].iloc[-1], 2)

    # -----------製作【日K-備註文字】-----------

    # 讀取圖表的寬度
    file_path = f"./{stockNo}-{(current_date - relativedelta(months=3)).year}0{(current_date - relativedelta(months=3)).month}-{current_date.year}0{current_date.month}.png"
    image = Image.open(file_path)
    target_width = image.width

    # 創建空白圖像
    image_width = target_width
    image_height = 60
    background_color = (255, 255, 255)  # 白色背景
    image = Image.new("RGB", (image_width, image_height), background_color)
    draw = ImageDraw.Draw(image)

    # 設置內容
    text = "日K" + "5MA" + str(latest_5MA) + "20MA" + str(latest_20MA) + "60MA" + str(latest_60MA)

    # 設置字體顏色
    text_color_gray = (74, 74, 74)  # 4a4a4a的RGB值
    text_color_orange = (255, 165, 0)  # 橘色的RGB值
    text_color_red = (255, 0, 0)  # 紅色的RGB值
    text_color_blue = (0, 0, 255)  # 藍色的RGB值

    # 設置字體
    font_size = 28
    font_path = "msjh.ttc"
    font = ImageFont.truetype(font_path, font_size)

    # 計算內容大小和位置（左側）
    text_bbox = draw.textbbox((0, 0), text, font=font)
    text_width = text_bbox[2] - text_bbox[0]
    text_height = text_bbox[3] - text_bbox[1]
    text_x = 10  # 內容左上角的X座標
    text_y = (image_height - text_height) // 3  # 內容垂直置中對齊

    # 在圖像上繪製內容
    draw.text((text_x + 40, text_y), "日K", fill=text_color_gray, font=font)
    draw.text((text_x + 230, text_y), "5MA", fill=text_color_blue, font=font)
    draw.text((text_x + 310, text_y), str(latest_5MA), fill=text_color_blue, font=font)
    draw.text((text_x + 560, text_y), "20MA", fill=text_color_red, font=font)
    draw.text((text_x + 650, text_y), str(latest_20MA), fill=text_color_red, font=font)
    draw.text((text_x + 900, text_y), "60MA", fill=text_color_orange, font=font)
    draw.text((text_x + 990, text_y), str(latest_60MA), fill=text_color_orange, font=font)

    # 保存图像
    image_path = str(stockNo) + "日K-備註文字.png"
    image.save(image_path)

    # 印出確認文字
    text = "已完成 " + str(stockNo) + "日K-備註文字.png"
    print(text)

    # ------------------------------讀取stock_info.csv------------------------------

    # 檔案路徑
    file_name = f"./{stockNo}_info.csv"

    # 讀取檔案
    df = pd.read_csv(file_name)

    # 列出欄位下的值
    update_date = df["資料日期"].tolist()[0]
    update_date_time = df["更新時間"].tolist()[0]
    stock_no = df["股票代碼"].tolist()[0]
    stock_name = df["股票名稱"].tolist()[0]
    listing_cabinet = df["上市櫃"].tolist()[0]
    industry = df["產業別"].tolist()[0]

    close_price = df["收盤價"].tolist()[0]
    # 如果值有小數點，就改為小數點後兩位(字串)
    if len(str(close_price).split(".")) == 2:
        close_price = "{:.2f}".format(close_price)
    # 如果值為整數，就維持(整數)
    else:
        close_price = close_price

    price_change = df["漲跌價"].tolist()[0]
    price_change_rate = df["漲跌幅"].tolist()[0]

    open_price = df["開盤"].tolist()[0]
    # 如果值有小數點，就改為小數點後兩位(字串)
    if len(str(open_price).split(".")) == 2:
        open_price = "{:.2f}".format(open_price)
    # 如果值為整數，就維持(整數)
    else:
        open_price = open_price

    high_price = df["最高"].tolist()[0]
    # 如果值有小數點，就改為小數點後兩位(字串)
    if len(str(high_price).split(".")) == 2:
        high_price = "{:.2f}".format(high_price)
    # 如果值為整數，就維持(整數)
    else:
        high_price = high_price

    low_price = df["最低"].tolist()[0]
    # 如果值有小數點，就改為小數點後兩位(字串)
    if len(str(low_price).split(".")) == 2:
        low_price = "{:.2f}".format(low_price)
    # 如果值為整數，就維持(整數)
    else:
        low_price = low_price

    deal_amount = df["成交量"].tolist()[0]

    # -------------取得【跌幅】----------------

    price_change_pic = price_change.split()[0] + " " + str(float(price_change.split()[1]))

    # --------------取得【今日日期】--------------

    # 將字符串解析為日期時間對象
    date_time_obj = datetime.strptime(update_date_time, "%m/%d %H:%M:%S")

    # 取得今天日期時間
    current_date = datetime.now()

    # 將解析的日期的年份設置為今年年份
    date_time_obj = date_time_obj.replace(year=current_date.year)

    # 提取日期部分
    update_date_pic = date_time_obj.strftime("%Y/%m/%d")

    # --------------製作【日K-小圖】--------------

    # 待輸出文字
    output_text = [
        " " + update_date_pic,
        "開：" + str(open_price),
        "高：" + str(high_price),
        "低：" + str(low_price),
        "收：" + str(close_price),
        "漲：" + price_change_pic,
        "幅：" + price_change_rate,
        "量：" + str(deal_amount)
    ]

    # 待儲存圖片路徑
    output_image_path = str(stockNo) + "日K-小圖.png"

    # 設定圖片大小和背景顏色
    image_width = 175
    image_height = 300
    background_color = (255, 255, 255)  # 白色

    # 創建一個白色背景的新圖片
    image = Image.new("RGB", (image_width, image_height), background_color)

    # 創建繪圖對象
    draw = ImageDraw.Draw(image)

    # 載入提供的中文字體
    font_size = 26
    font_path = "msjh.ttc"
    font = ImageFont.truetype(font_path, font_size)

    # 設定初始的垂直位置用於文字渲染
    y_position = 10

    # 將每行文字渲染到圖片上
    for line in output_text:
        draw.text((10, y_position), line, font=font, fill=(0, 0, 0))  # 黑色文字
        y_position += 35                          # 調整此值以改變行之間的垂直間距

    # 繪製灰色邊框
    border_color = (192, 192, 192)  # 灰色
    border_thickness = 1
    border_rectangle = [(0, 0), (image_width - 1, image_height - 1)]
    draw.rectangle(border_rectangle, outline=border_color, width=border_thickness)

    # 將圖片保存到指定路徑
    image.save(output_image_path)

    # 印出確認文字
    print("已完成 " + str(stockNo) + "日K-小圖.png")

    # ---------合併PNG【K線_小圖】+【圖表】----------------------------------

    # 讀取第一張PNG圖片
    image1_path = str(stockNo) + "日K-小圖.png"
    image1 = Image.open(image1_path)

    # 讀取第二張PNG圖片
    image2_path = f"./{stockNo}-{(current_date - relativedelta(months=3)).year}0{(current_date - relativedelta(months=3)).month}-{current_date.year}0{current_date.month}.png"
    image2 = Image.open(image2_path)

    # # 確保第一張圖片大小不超過第二張圖片的大小
    # if image1.size[0] > image2.size[0] or image1.size[1] > image2.size[1]:
    #     raise ValueError("第一張圖片的尺寸大於第二張圖片，無法粘貼在左上方。")

    # 在第二張圖片的左上角位置黏貼上第一張圖片
    image2.paste(image1, (10, 10))

    # 儲存圖片
    image_path = str(stockNo) + "日K-小圖+圖表.png"
    image2.save(image_path)

    # 印出確認文字
    print("已完成 " + str(stockNo) + "日K-小圖+圖表.png")

    # --------------合併PNG【日K-備註文字】+【日K-小圖+圖表】----------------------------------

    # 讀取第一張PNG圖片
    image1_path = str(stockNo) + "日K-備註文字.png"
    image1 = Image.open(image1_path)

    # 讀取第二張PNG圖片
    image2_path = str(stockNo) + "日K-小圖+圖表.png"
    image2 = Image.open(image2_path)

    # 獲取第一張圖片的寬度和高度
    width1, height1 = image1.size

    # 獲取第二張圖片的寬度和高度
    width2, height2 = image2.size

    # 確定合併後圖片的寬度和高度
    merged_width = max(width1, width2)
    merged_height = height1 + height2

    # 創建一個新的空白圖片，尺寸為合併後的寬度和高度
    merged_image = Image.new("RGB", (merged_width, merged_height))

    # 將第一張圖片黏貼到合併圖片的上半部分
    merged_image.paste(image1, (0, 0))

    # 將第二張圖片黏貼到合併圖片的下半部分
    merged_image.paste(image2, (0, height1))

    # 儲存合併後的圖片
    merged_image_path = str(stockNo) + "日K.png"
    merged_image.save(merged_image_path)

    # 印出確認文字
    print("已完成 " + str(stockNo) + "日K.png")

# 確認是否有最新【日K所需資料】+ 同時製作【日K所需資料】、【日K線圖】
def check_file_day_candlestick_data_csv(stockNo):

    # 使用twstock模組
    stock = twstock.Stock(str(stockNo))

    # 獲取當前時間
    now_time = datetime.now()

    # 獲取當前日期
    today_date = now_time.date()

    from datetime import time

    # 設定台灣股票每日交易時間
    trade_time = time(hour=13, minute=30)

    # 檔案路徑
    file_name_all = f"./{stockNo}-{(today_date - relativedelta(months=3)).year}{(today_date - relativedelta(months=3)).month:02d}-{today_date.year}{today_date.month:02d}.csv"

    # 如果檔案存在
    if os.path.exists(file_name_all):

        # 判斷當天是星期幾
        weekday = today_date.weekday()

        # 如果當天是星期一
        if weekday == 0:

            # 抓取日期不變
            today_date = today_date

            # 如果現在時間介於 00:00~13:30
            if now_time.time() < trade_time:
                # 抓取日期調整到星期五(即為3天前)
                today_date = today_date - timedelta(days=3)

                # 讀取檔案
                df = pd.read_csv(file_name_all)

                # 如果"Date"欄位包含當天日期
                if today_date.strftime("%Y-%m-%d") in df["Date"].values:
                    # 列印確認文字
                    print(
                        f"不須更新 {stockNo}-{(today_date - relativedelta(months=3)).year}{(today_date - relativedelta(months=3)).month:02d}-{today_date.year}{today_date.month:02d}.csv")

                # 如果"Date"欄位不包含當天日期
                else:
                    # 重新抓再合併
                    day_candlestick_data_csv(stockNo)
                    day_candlestick_png(stockNo)

            # 如果現在時間介於 13:30~24:00
            else:
                # 讀取檔案
                df = pd.read_csv(file_name_all)

                # 如果"Date"欄位包含當天日期
                if today_date.strftime("%Y-%m-%d") in df["Date"].values:
                    # 列印確認文字
                    text = f"不須更新 {stockNo}-{(today_date - relativedelta(months=3)).year}{(today_date - relativedelta(months=3)).month:02d}-{today_date.year}{today_date.month:02d}.csv"
                    print(text)

                # 如果"Date"欄位不包含當天日期
                else:
                    # 重新抓再合併
                    day_candlestick_data_csv(stockNo)
                    day_candlestick_png(stockNo)


        # 如果當天是星期二~五
        elif weekday <= 4:

            # 抓取日期不變
            today_date = today_date

            # 如果現在時間介於 13:30~24:00
            if now_time.time() >= trade_time:

                # 讀取檔案
                df = pd.read_csv(file_name_all)

                # 如果"Date"欄位包含當天日期
                if today_date.strftime("%Y-%m-%d") in df["Date"].values:
                    # 列印確認文字
                    text = f"不須更新 {stockNo}-{(today_date - relativedelta(months=3)).year}{(today_date - relativedelta(months=3)).month:02d}-{today_date.year}{today_date.month:02d}.csv"
                    print(text)

                # 如果"Date"欄位不包含當天日期
                else:
                    # 重新抓再合併
                    day_candlestick_data_csv(stockNo)
                    day_candlestick_png(stockNo)

            # 如果現在時間介於 00:00~13:30
            else:
                # 讀取檔案
                df = pd.read_csv(file_name_all)

                # 取得昨天日期
                yesterday = today_date - timedelta(days=1)

                # 如果"Date"欄位包含昨天日期
                if yesterday.strftime("%Y-%m-%d") in df["Date"].values:
                    # 列印確認文字
                    text = f"不須更新 {stockNo}-{(today_date - relativedelta(months=3)).year}{(today_date - relativedelta(months=3)).month:02d}-{today_date.year}{today_date.month:02d}.csv"
                    print(text)

                # 如果"Date"欄位不包含昨天日期
                else:
                    # 重新抓再合併
                    day_candlestick_data_csv(stockNo)
                    day_candlestick_png(stockNo)

        # 如果當天是星期六~日
        else:
            # 抓取日期調整到星期五
            today_date = today_date - timedelta(days=weekday - 4)

            # 讀取檔案
            df = pd.read_csv(file_name_all)

            # 如果"Date"欄位包含當天日期
            if today_date.strftime("%Y-%m-%d") in df["Date"].values:
                # 列印確認文字
                text = f"不須更新 {stockNo}-{(today_date - relativedelta(months=3)).year}{(today_date - relativedelta(months=3)).month:02d}-{today_date.year}{today_date.month:02d}.csv"
                print(text)

            # 如果"Date"欄位不包含當天日期
            else:
                # 重新抓再合併
                day_candlestick_data_csv(stockNo)
                day_candlestick_png(stockNo)

    # 如果檔案不存在
    else:
        # 重新抓再合併
        day_candlestick_data_csv(stockNo)
        day_candlestick_png(stockNo)

# 日K-bubble
def day_candlestick(stockNo):

    # 製作【股票資訊.csv】
    stock_info_csv(stockNo)

    # 確認【營收表格】是否存在
    check_file_day_candlestick_data_csv(stockNo)

    # ------------------圖片上傳到Imgur------------------

    # 讀取指定檔案，取得IMGUR_ID
    with open("env.json") as f:
        env = json.load(f)

    # 讀入IMGUR_ID
    CLIENT_ID = env["YOUR_IMGUR_ID"]
    # 指定圖片路徑
    PATH = str(stockNo) + "日K.png"
    # 指定圖片標題
    title = str(stockNo) + "日K.png"

    # 將圖片上傳到IMGUR
    im = pyimgur.Imgur(CLIENT_ID)
    # 帶入圖片路徑及標題
    uploaded_image = im.upload_image(PATH, title=title)

    # ------------------------------讀取stock_info.csv------------------------------

    # 指定檔案路徑
    file_name = f"./{stockNo}_info.csv"

    # 讀取指定檔案
    df = pd.read_csv(file_name)

    # 列出欄位下的值
    update_date = df["資料日期"].tolist()[0]
    update_date_time = df["更新時間"].tolist()[0]
    stock_no = df["股票代碼"].tolist()[0]
    stock_name = df["股票名稱"].tolist()[0]
    listing_cabinet = df["上市櫃"].tolist()[0]
    industry = df["產業別"].tolist()[0]

    close_price = df["收盤價"].tolist()[0]
    # 如果值有小數點，就改為小數點後兩位(字串)
    if len(str(close_price).split(".")) == 2:
        close_price = "{:.2f}".format(close_price)
    # 如果值為整數，就維持(整數)
    else:
        close_price = close_price

    price_change = df["漲跌價"].tolist()[0]
    price_change_rate = df["漲跌幅"].tolist()[0]

    open_price = df["開盤"].tolist()[0]
    # 如果值有小數點，就改為小數點後兩位(字串)
    if len(str(open_price).split(".")) == 2:
        open_price = "{:.2f}".format(open_price)
    # 如果值為整數，就維持(整數)
    else:
        open_price = open_price

    high_price = df["最高"].tolist()[0]
    # 如果值有小數點，就改為小數點後兩位(字串)
    if len(str(high_price).split(".")) == 2:
        high_price = "{:.2f}".format(high_price)
    # 如果值為整數，就維持(整數)
    else:
        high_price = high_price

    low_price = df["最低"].tolist()[0]
    # 如果值有小數點，就改為小數點後兩位(字串)
    if len(str(low_price).split(".")) == 2:
        low_price = "{:.2f}".format(low_price)
    # 如果值為整數，就維持(整數)
    else:
        low_price = low_price

    deal_amount = df["成交量"].tolist()[0]

    # 建立FlexSendMessage
    message = FlexSendMessage(
        alt_text = str(stock_name) + str(stock_no) + " 日K線圖",
        contents = {
            "type": "bubble",
            "size": "giga",
            "header": {
                "type": "box",
                "layout": "vertical",
                "contents": [
                    {
                        "type": "box",
                        "layout": "horizontal",
                        "contents": [
                            {
                                "type": "text",
                                "text": "⇧" + " 加好友",
                                "color": "#d90b0b",
                                "weight": "bold",
                                "size": "xl",
                                "action": {
                                    "type": "uri",
                                    "label": "action",
                                    "uri": "https://line.me/R/ti/p/%40636hvuxg"
                                }
                            }
                        ]
                    },
                    {
                        "type": "box",
                        "layout": "horizontal",
                        "contents": [
                            {
                                "type": "text",
                                "text": str(stock_name),
                                "size": "xxl",
                                "weight": "bold",
                                "flex": 0
                            },
                            {
                                "type": "text",
                                "text": str(stock_no),
                                "gravity": "bottom",
                                "flex": 2
                            },
                            {
                                "type": "filler"
                            },
                            {
                                "type": "text",
                                "text": str(close_price),
                                "size": "xxl",
                                "flex": 3,
                                "align": "end"
                            },
                            {
                                "type": "box",
                                "layout": "vertical",
                                "contents": [
                                    {
                                        "type": "text",
                                        "text": str(price_change),
                                        "align": "center",
                                        "size": "xs"
                                    },
                                    {
                                        "type": "text",
                                        "text": str(price_change_rate),
                                        "align": "center",
                                        "size": "xs"
                                    }
                                ],
                                "flex": 0,
                                "margin": "sm"
                            }
                        ]
                    },
                    {
                        "type": "box",
                        "layout": "baseline",
                        "contents": [
                            {
                                "type": "text",
                                "text": str(listing_cabinet) + " / " + str(industry),
                                "flex": 1,
                                "size": "xs",
                                "color": "#C5C4C7"
                            },
                            {
                                "type": "text",
                                "text": "股價更新時間：" + str(update_date_time),
                                "size": "xs",
                                "margin": "none",
                                "flex": 2,
                                "align": "end",
                                "color": "#C5C4C7"
                            }
                        ]
                    }
                ],
                "backgroundColor": "#F6F6FC"
            },
            "hero": {
                "type": "image",
                "url": uploaded_image.link,
                "size": "full",
                "aspectRatio": "20:13",
                "aspectMode": "fit",
                "action":{
                    "type": "uri",
                    "uri": "https://www.nstock.tw/stock_info?stock_id=2330&utm_source=line&utm_medium=line_bot"
                }
            },
            "footer": {
                "type": "box",
                "layout": "horizontal",
                "spacing": "none",
                "contents": [
                    {
                        "type": "text",
                        "text": "選項：",
                        "size": "xs",
                        "color": "#787878"
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "即時",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": "查" + str(stock_name)
                        }
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "K線",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "日K"
                        },
                        "color": "#d11111",
                        "weight": "bold"
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "法人",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "法人"
                        }
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "EPS",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "EPS"
                        }
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "營收",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "營收"
                        }
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "股利",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "股利"
                        }
                    },
                    {
                        "type": "separator"
                    },
                    {
                        "type": "text",
                        "text": "持股",
                        "size": "xs",
                        "align": "center",
                        "action": {
                            "type": "message",
                            "label": "action",
                            "text": str(stock_name) + "持股"
                        }
                    }
                ]
            }
        }
    )

    return message

# ------------------------------------------------------------

# 查股價指令教學
def nstock_code_list():

    list = ["台積電股價", "台積電日K", "台積電法人", "台積電EPS", "台積電營收", "台積電股利", "台積電持股", "大盤即時", "櫃買即時", "大盤K線", "櫃買K線" ]

    encoded_list = word_to_utf8(list)

    # linebot ID設定
    linebot_id = "@636hvuxg"

    message = FlexSendMessage(
        alt_text = "查股價指令教學",
        contents = {
            "type": "bubble",
            "size": "giga",
            "hero": {
            "type": "image",
            "url": "https://i.imgur.com/6fhie9m.png",
            "aspectMode": "fit",
            "aspectRatio": "15:2",
            "size": "full",
            "backgroundColor": "#F6F6FC"
            },
            "body": {
            "type": "box",
            "layout": "vertical",
            "contents": [
              {
                "type": "text",
                "text": "以查詢台積電(2330)為例：",
                "size": "sm"
              },
              {
                "type": "box",
                "layout": "baseline",
                "contents": [
                  {
                    "type": "text",
                    "text": "台積電股價(代碼P)",
                    "size": "sm",
                    "decoration": "underline",
                    "flex": 3,
                    "action": {
                      "type": "uri",
                      "label": "action",
                      "uri": f"https://line.me/R/oaMessage/{linebot_id}/?" + encoded_list[list.index("台積電股價")]
                    }
                  },
                  {
                    "type": "text",
                    "text": "➔",
                    "size": "sm",
                    "align": "center"
                  },
                  {
                    "type": "text",
                    "text": "看即時行情",
                    "size": "sm",
                    "flex": 4
                  }
                ],
                "margin": "md"
              },
              {
                "type": "box",
                "layout": "baseline",
                "contents": [
                  {
                    "type": "text",
                    "text": "台積電K線(代碼K)",
                    "size": "sm",
                    "decoration": "underline",
                    "flex": 3,
                    "action": {
                      "type": "uri",
                      "label": "action",
                      "uri": f"https://line.me/R/oaMessage/{linebot_id}/?" + encoded_list[list.index("台積電日K")]
                    }
                  },
                  {
                    "type": "text",
                    "text": "➔",
                    "size": "sm",
                    "align": "center"
                  },
                  {
                    "type": "text",
                    "text": "看K線與成交量",
                    "size": "sm",
                    "flex": 4
                  }
                ],
                "margin": "md"
              },
              {
                "type": "box",
                "layout": "baseline",
                "contents": [
                  {
                    "type": "text",
                    "text": "台積電法人(代碼T)",
                    "size": "sm",
                    "decoration": "underline",
                    "flex": 3,
                    "action": {
                      "type": "uri",
                      "label": "action",
                      "uri": f"https://line.me/R/oaMessage/{linebot_id}/?" + encoded_list[list.index("台積電法人")]
                    }
                  },
                  {
                    "type": "text",
                    "text": "➔",
                    "size": "sm",
                    "align": "center"
                  },
                  {
                    "type": "text",
                    "text": "看三大法人買賣超",
                    "size": "sm",
                    "flex": 4
                  }
                ],
                "margin": "md"
              },
              {
                "type": "box",
                "layout": "baseline",
                "contents": [
                  {
                    "type": "text",
                    "text": "台積電EPS(代碼E)",
                    "size": "sm",
                    "decoration": "underline",
                    "flex": 3,
                    "action": {
                      "type": "uri",
                      "label": "action",
                      "uri": f"https://line.me/R/oaMessage/{linebot_id}/?" + encoded_list[list.index("台積電EPS")]
                    }
                  },
                  {
                    "type": "text",
                    "text": "➔",
                    "size": "sm",
                    "align": "center"
                  },
                  {
                    "type": "text",
                    "text": "看每股盈餘",
                    "size": "sm",
                    "flex": 4
                  }
                ],
                "margin": "md"
              },
              {
                "type": "box",
                "layout": "baseline",
                "contents": [
                  {
                    "type": "text",
                    "text": "台積電營收(代碼F)",
                    "size": "sm",
                    "decoration": "underline",
                    "flex": 3,
                    "action": {
                      "type": "uri",
                      "label": "action",
                      "uri": f"https://line.me/R/oaMessage/{linebot_id}/?" + encoded_list[list.index("台積電營收")]
                    }
                  },
                  {
                    "type": "text",
                    "text": "➔",
                    "size": "sm",
                    "align": "center"
                  },
                  {
                    "type": "text",
                    "text": "看每月營收",
                    "size": "sm",
                    "flex": 4
                  }
                ],
                "margin": "md"
              },
              {
                "type": "box",
                "layout": "baseline",
                "contents": [
                  {
                    "type": "text",
                    "text": "台積電股利(代碼D)",
                    "size": "sm",
                    "decoration": "underline",
                    "flex": 3,
                    "action": {
                      "type": "uri",
                      "label": "action",
                      "uri": f"https://line.me/R/oaMessage/{linebot_id}/?" + encoded_list[list.index("台積電股利")]
                    }
                  },
                  {
                    "type": "text",
                    "text": "➔",
                    "size": "sm",
                    "align": "center"
                  },
                  {
                    "type": "text",
                    "text": "看股利、殖利率、配發率",
                    "size": "sm",
                    "flex": 4
                  }
                ],
                "margin": "md"
              },
              {
                "type": "box",
                "layout": "baseline",
                "contents": [
                  {
                    "type": "text",
                    "text": "台積電持股(代碼H)",
                    "size": "sm",
                    "decoration": "underline",
                    "flex": 3,
                    "action": {
                      "type": "uri",
                      "label": "action",
                      "uri": f"https://line.me/R/oaMessage/{linebot_id}/?" + encoded_list[list.index("台積電持股")]
                    }
                  },
                  {
                    "type": "text",
                    "text": "➔",
                    "size": "sm",
                    "align": "center"
                  },
                  {
                    "type": "text",
                    "text": "看大股東持股變化",
                    "size": "sm",
                    "flex": 4
                  }
                ],
                "margin": "md"
              },
              {
                "type": "box",
                "layout": "vertical",
                "contents": [
                  {
                    "type": "text",
                    "text": "\"台積電\"也可以換成股號查詢 例如:2330股價, 2330K線, P2330",
                    "size": "xxs"
                  }
                ],
                "margin": "md"
              },
              {
                "type": "box",
                "layout": "vertical",
                "contents": [
                  {
                    "type": "separator"
                  },
                  {
                    "type": "text",
                    "text": "另可查詢台股大盤貨櫃買指數：",
                    "size": "sm",
                    "margin": "md"
                  }
                ],
                "margin": "md"
              },
              {
                "type": "box",
                "layout": "baseline",
                "contents": [
                  {
                    "type": "text",
                    "text": "大盤即時(P大盤)",
                    "size": "sm",
                    "decoration": "underline",
                    "flex": 3,
                    "action": {
                      "type": "uri",
                      "label": "action",
                      "uri": f"https://line.me/R/oaMessage/{linebot_id}/?" + encoded_list[list.index("大盤即時")]
                    }
                  },
                  {
                    "type": "text",
                    "text": "➔",
                    "size": "sm",
                    "align": "center",
                    "action": {
                      "type": "uri",
                      "label": "action",
                      "uri": "http://linecorp.com/"
                    }
                  },
                  {
                    "type": "text",
                    "text": "看即時行情",
                    "size": "sm",
                    "flex": 4
                  }
                ],
                "margin": "md"
              },
              {
                "type": "box",
                "layout": "baseline",
                "contents": [
                  {
                    "type": "text",
                    "text": "櫃買即時(P櫃買)",
                    "size": "sm",
                    "decoration": "underline",
                    "flex": 3,
                    "action": {
                      "type": "uri",
                      "label": "action",
                      "uri": f"https://line.me/R/oaMessage/{linebot_id}/?" + encoded_list[list.index("櫃買即時")]
                    }
                  },
                  {
                    "type": "text",
                    "text": "➔",
                    "size": "sm",
                    "align": "center"
                  },
                  {
                    "type": "text",
                    "text": "看即時行情",
                    "size": "sm",
                    "flex": 4
                  }
                ],
                "margin": "md"
              },
              {
                "type": "box",
                "layout": "baseline",
                "contents": [
                  {
                    "type": "text",
                    "text": "大盤K線(K大盤)",
                    "size": "sm",
                    "decoration": "underline",
                    "flex": 3,
                    "action": {
                      "type": "uri",
                      "label": "action",
                      "uri": f"https://line.me/R/oaMessage/{linebot_id}/?" + encoded_list[list.index("大盤K線")]
                    }
                  },
                  {
                    "type": "text",
                    "text": "➔",
                    "size": "sm",
                    "align": "center"
                  },
                  {
                    "type": "text",
                    "text": "看K線與成交量",
                    "size": "sm",
                    "flex": 4
                  }
                ],
                "margin": "md"
              },
              {
                "type": "box",
                "layout": "baseline",
                "contents": [
                  {
                    "type": "text",
                    "text": "櫃買K線(K櫃買)",
                    "size": "sm",
                    "decoration": "underline",
                    "flex": 3,
                    "action": {
                      "type": "uri",
                      "label": "action",
                      "uri": f"https://line.me/R/oaMessage/{linebot_id}/?" + encoded_list[list.index("櫃買K線")]
                    }
                  },
                  {
                    "type": "text",
                    "text": "➔",
                    "size": "sm",
                    "align": "center"
                  },
                  {
                    "type": "text",
                    "text": "看K線與成交量",
                    "size": "sm",
                    "flex": 4
                  }
                ],
                "margin": "md"
              }
            ]
          }
        }
    )

    return message

# 將文字清單轉換為utf8
def word_to_utf8(list):

    encoded_list = [quote(item.encode("utf-8")) for item in list]

    return encoded_list

# ------------------------------------------------------------------

app = Flask(__name__)

# 監聽所有來自 /callback 的 Post Request
@app.route("/callback", methods=["POST"])
def callback():
    # get X-Line-Signature header value
    signature = request.headers["X-Line-Signature"]

    # get request body as text
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)

    # handle webhook body
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        print("Invalid signature. Please check your channel access token/channel secret.")
        abort(400)
    except Exception as e:
        print("Error occurred while handling webhook: ", e)
        abort(500)

    return "OK"

# 訊息傳遞區塊(根據訊息內容，做處理)

# 告訴LINE Bot，當使用者輸入訊息，且訊息是文字時，我們就執行以下的程式碼，
@handler.add(MessageEvent, message=TextMessage)

def handle_message(event):
    message = event.message.text

    # 輸入：EPS台積電、台積電EPS、EPS2330、2330EPS
    # 輸出：Bubble
    if "EPS" in event.message.text:

        # 只保留 股票名稱/代號
        stockNo = message.replace("EPS", "").replace(" ", "")

        # 確認股票代碼清單檔案是否存在
        check_file_stock_name_no_csv()

        try:
            # 轉換股票名稱/代碼
            stockNo_result = find_stock_code(stockNo)

            # 使用函數
            message = EPS(stockNo_result)

            # 使用API
            line_bot_api.reply_message(
                event.reply_token,
                message
            )

        except Exception as e:
            # 如果發生錯誤，回覆指定錯誤訊息
            text = "請確認股票名稱、股票代號、指令是否有誤。"
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text)
            )

    # 輸入：E台積電、台積電E、E2330、2330E
    # 輸出：Bubble
    if "E" in event.message.text:

        # 只保留 股票名稱/代號
        stockNo = message.replace("E", "").replace(" ", "")

        # 確認股票代碼清單檔案是否存在
        check_file_stock_name_no_csv()

        try:
            # 轉換股票名稱/代碼
            stockNo_result = find_stock_code(stockNo)

            # 使用函數
            message = EPS(stockNo_result)

            # 使用API
            line_bot_api.reply_message(
                event.reply_token,
                message
            )

        except Exception as e:
            # 如果發生錯誤，回覆指定錯誤訊息
            text = "請確認股票名稱、股票代號、指令是否有誤。"
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text)
            )

    # 輸入：持股台積電、台積電持股、持股2330、2330持股
    # 輸出：Bubble
    elif "持股" in event.message.text:

        # 只保留 股票名稱/代號
        stockNo = message.replace("持股", "").replace(" ", "")

        # 確認股票代碼清單檔案是否存在
        check_file_stock_name_no_csv()

        try:
            # 轉換股票名稱/代碼
            stockNo_result = find_stock_code(stockNo)

            # 使用函數
            message = shareholder(stockNo_result)

            # 使用API
            line_bot_api.reply_message(
                event.reply_token,
                message
            )

        except Exception as e:
            # 如果發生錯誤，回覆指定錯誤訊息
            text = "請確認股票名稱、股票代號、指令是否有誤。"
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text)
            )

    # 輸入：H台積電、台積電H、H2330、2330H
    # 輸出：Bubble
    elif "H" in event.message.text:

        # 只保留 股票名稱/代號
        stockNo = message.replace("H", "").replace(" ", "")

        # 確認股票代碼清單檔案是否存在
        check_file_stock_name_no_csv()

        try:
            # 轉換股票名稱/代碼
            stockNo_result = find_stock_code(stockNo)

            # 使用函數
            message = shareholder(stockNo_result)

            # 使用API
            line_bot_api.reply_message(
                event.reply_token,
                message
            )

        except Exception as e:
            # 如果發生錯誤，回覆指定錯誤訊息
            text = "請確認股票名稱、股票代號、指令是否有誤。"
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text)
            )

    # 輸入：營收台積電、台積電營收、營收2330、2330營收
    # 輸出：Bubble
    elif "營收" in event.message.text:

        # 只保留 股票名稱/代號
        stockNo = message.replace("營收", "").replace(" ", "")

        # 確認【股票代碼清單】是否存在
        check_file_stock_name_no_csv()

        try:
            # 轉換股票名稱/代碼
            stockNo_result = find_stock_code(stockNo)

            # 使用函數
            message = revenue(stockNo_result)

            # 使用API
            line_bot_api.reply_message(
                event.reply_token,
                message
            )

        except Exception as e:
            # 如果發生錯誤，回覆指定錯誤訊息
            text = "請確認股票名稱、股票代號、指令是否有誤。"
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text)
            )

    # 輸入：F台積電、台積電F、F2330、2330F
    # 輸出：Bubble
    elif "F" in event.message.text:

        # 只保留 股票名稱/代號
        stockNo = message.replace("F", "").replace(" ", "")

        # 確認【股票代碼清單】是否存在
        check_file_stock_name_no_csv()

        try:
            # 轉換股票名稱/代碼
            stockNo_result = find_stock_code(stockNo)

            # 使用函數
            message = revenue(stockNo_result)

            # 使用API
            line_bot_api.reply_message(
                event.reply_token,
                message
            )

        except Exception as e:
            # 如果發生錯誤，回覆指定錯誤訊息
            text = "請確認股票名稱、股票代號、指令是否有誤。"
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text)
            )

    # 輸入：日K台積電、台積電日K、日K2330、2330日K
    # 輸出：Bubble
    elif "日K" in event.message.text:

        # 只保留 股票名稱/代號
        stockNo = message.replace("日K", "").replace(" ", "")

        # 確認【股票代碼清單】是否存在
        check_file_stock_name_no_csv()

        try:
            # 轉換股票名稱/代碼
            stockNo_result = find_stock_code(stockNo)

            # 使用函數
            message = day_candlestick(stockNo_result)

            # 使用API
            line_bot_api.reply_message(
                event.reply_token,
                message
            )

        except Exception as e:
            # 如果發生錯誤，回覆指定錯誤訊息
            text = "請確認股票名稱、股票代號、指令是否有誤。"
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text)
            )

    # 輸入：K台積電、台積電K、K2330、2330K
    # 輸出：Bubble
    elif "K" in event.message.text:

        # 只保留 股票名稱/代號
        stockNo = message.replace("K", "").replace(" ", "")

        # 確認【股票代碼清單】是否存在
        check_file_stock_name_no_csv()

        try:
            # 轉換股票名稱/代碼
            stockNo_result = find_stock_code(stockNo)

            # 使用函數
            message = day_candlestick(stockNo_result)

            # 使用API
            line_bot_api.reply_message(
                event.reply_token,
                message
            )

        except Exception as e:
            # 如果發生錯誤，回覆指定錯誤訊息
            text = "請確認股票名稱、股票代號、指令是否有誤。"
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text)
            )

   # 查股價指令教學
    elif event.message.text == "指令":

        # 使用公式
        message = nstock_code_list()

        # 使用API
        line_bot_api.reply_message(
            event.reply_token,
            message
        )

    # 輸入內容，機器人就會原封不動的回傳給一模一樣的內容
    else:
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(message))

if __name__ == "__main__":
    app.run()
