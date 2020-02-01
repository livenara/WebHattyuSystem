"""

発注専用ウェブサイトの自動入力システム

基幹システムから注文拠点ごとにデータを修正してエクセルへ貼付。
注文システム起動。
住所を確認してエンター完了。

注文完了後Gsuite,GMailメールアドレスへ注文確認メールが届くので
深夜、注文当日のGoogleAppsScriptでスプレッドシートへ注文内容をまとめて
エクセルへ出力したのち指定のフォルダへ移動後に担当者へメールで処理後に報告。

"""

from selenium import webdriver
from selenium.webdriver.support.ui import Select # 選択画面
from bs4 import BeautifulSoup
import pandas as pd 
import csv
import time
import jaconv # 半角ｶﾅ文字対応
import pyodbc # DB

# ChromeDriver
chromedriverPath = "chromedriver.exe"

# ログイン
login_page_url = "https://*****.bcart.jp/login.php"
login_id = "****"
login_pass = "****"

# 基本データ
ExcelFile = "tmp_SmileBS_修正登録ファイル.xls" # 基幹システムからエクスポートしたデータを入れる。
CsvFileName = 'tmp_Web_登録用データ.csv'           # 登録用の一時ファイルCSVを作成する。

# 商品コードに該当商品の注文用URLを入れる。
SyoHinCode = {}
SyoHinCode['46990002'] = "https://*****.bcart.jp/product.php?id=3" # ポテトチップス(スタンダード)
SyoHinCode['46990102'] = "https://*****.bcart.jp/product.php?id=6" # ポテトチップス(コンソメ)

# 得意先データベースから個別の得意先データを取り出す。
def TokuiSakiData(TokuiSakiBanGo):

    TokuiSakiBanGo = str(TokuiSakiBanGo).zfill(8) # 8桁ゼロ埋め

    ConfigDic = {}
    ConfigDic['instance'] = "***.***.***.***\SMILEBS" # インスタンス
    ConfigDic['user'] = "***"                         # ユーザー
    ConfigDic['pasword'] = "***"                      # パスワード
    ConfigDic['db'] = "****_DB"                       # DB #######がテストDB

    connection = "DRIVER={SQL Server};SERVER=" + ConfigDic['instance'] + ";uid=" + ConfigDic['user'] + \
                ";pwd=" + ConfigDic['pasword'] + ";DATABASE=" + ConfigDic['db']
    con = pyodbc.connect(connection)

    TABLE = "****_T_TOKUISAKI_MST"  # 得意先マスターテーブル
    cur = con.cursor()
    sql = "select * FROM " + TABLE + " WHERE TOK_CD = " + str(TokuiSakiBanGo)
    cur.execute(sql)
    record = cur.fetchone()
    cur.close()
    con.close()

    TokuiSakiDataDic = {}
    TokuiSakiDataDic['得意先番号'] = int(record[0])
    TokuiSakiDataDic['得意先名'] = record[1].strip()
    TokuiSakiDataDic['郵便番号'] = record[3].strip()
    TokuiSakiDataDic['住所'] = record[4].strip() + record[5].strip()
    TokuiSakiDataDic['電話番号'] = record[6].strip()
    TokuiSakiDataDic['ルート'] = int(record[15])     # 注文済みデータ収集用
    TokuiSakiDataDic['請求書拠点'] = int(record[80]) # 注文済みデータ収集用

    return TokuiSakiDataDic

# ChromeDriverのパスを引数に指定しChromeを起動
driver = webdriver.Chrome(chromedriverPath)

# BeatifulSoupパーサー
def BsParse(source):
    return BeautifulSoup(source, 'html.parser')

# 登録修正ファイルのデータをCSV化
def FileMake(ExcelFile, CsvFile):
    KobetsuNum = 0
    df = pd.read_excel(ExcelFile, skiprows=1, header=1)
    df_check = df[ df['個別発注番号'] > KobetsuNum ]
    df_check.to_csv(CsvFile, header=0, index=0)

# 発注用にデータを作成したCSVをリスト化
def HattyuDataCsv(CsvFileName):
    HattyuList = []
    with open(CsvFileName,"r",encoding="utf-8")as f:
        file = csv.reader(f)
        for x in file:
            HattyuList.append(x)

    # 昇順ソートで確認リストの順に処理ができる。
    HattyuList.sort(key=lambda x: x[3], reverse=False)
    return HattyuList

#ログイン
def Login(login_page_url,login_id,login_pass):

    #ログインページへ
    driver.get(login_page_url)

    Xpath_loginidbox = "/html/body/div[1]/div/div/form/section[1]/table/tbody/tr[1]/td/input"
    driver.find_element_by_xpath(Xpath_loginidbox).send_keys(login_id)

    Xpath_loginpassbox = "/html/body/div[1]/div/div/form/section[1]/table/tbody/tr[2]/td/input"
    driver.find_element_by_xpath(Xpath_loginpassbox).send_keys(login_pass)

    Xpath_loginbutton = "/html/body/div[1]/div/div/form/section[2]/input"
    driver.find_element_by_xpath(Xpath_loginbutton).click()

# メイン
def SyoHinPageData(HattyuList):

    for h in HattyuList:

        driver.get(SyoHinCode[h[8]]) # 商品ページ遷移
        source = driver.page_source
        soup = BsParse(source)

        TokuiNum = h[0] # 得意先番号
        #TokuiName = jaconv.h2z(HattyuList[1],digit=False, ascii=False)
        TokuiRyaku = h[1]
        KobetsuBanGo = str(int(float(h[4])))
        TyakaBi = h[7]
        TyuMonKoSu = str(int(h[10])) # 注文個数
        print("*** 発注情報 *****************************************************")
        print("得意先番号: ",str(TokuiNum))
        print("得意先略称: ",str(TokuiRyaku))
        print("★ 着荷日　: ",str(TyakaBi))
        print("★ 個別番号: ",str(KobetsuBanGo))
        print("★ 注文個数: ",str(TyuMonKoSu),"\n")

        # 着荷日順番把握
        title_text = soup.find_all('h2') # 全着荷日箇所 ["[**]]YYYYmmdd着荷商品","[**]YYYYmmdd着荷商品"~
        TyakaBi_result = TyakaBi.split("-")
        TyakaBiText = TyakaBi_result[0] + TyakaBi_result[1] + TyakaBi_result[2].zfill(2)

        for x in title_text:
            result = x.text.replace("[**]","")
            result = result.replace("着荷商品","")

            if result == TyakaBiText:
                #print(result,TyakaBiText)
                Xpath_index = int(title_text.index(x)) + 1

        # 注文個数入力
        Xpath_KonyuSu = "/html/body/div[1]/div/div/form/section[1]/table/tbody/tr[" + str(Xpath_index) + "]/td[3]/div[2]/div[2]/input"
        driver.find_element_by_xpath(Xpath_KonyuSu).send_keys(TyuMonKoSu)

        # カートに入れるボタンを押す
        Xpath_CurtButton = "/html/body/div[1]/div/div/form/section[2]/button"
        driver.find_element_by_xpath(Xpath_CurtButton).click()

        time.sleep(2) #カートに入れるポップアップ後インターバルがないとカートに商品が入らないことがある

        # カートを見る
        driver.get("https://*****.bcart.jp/cart.php")

        # 注文へ進む
        Xpath_TyuMonButton = '//*[@id="cartForm1"]/div[4]/ul/li[4]/button/span'
        driver.find_element_by_xpath(Xpath_TyuMonButton).click()

        # 別住所へ配送する。
        Xpath_BetsuHaisouButton = '/html/body/div[1]/div/div/form/section[3]/div/table[1]/tbody/tr/td/label[2]'
        driver.find_element_by_xpath(Xpath_BetsuHaisouButton).click()

        # 配送先 会社名 に得意先番号入れる
        Xpath_KaisyaName = "/html/body/div[1]/div/div/form/section[3]/div/table[2]/tbody[2]/tr[1]/td/input"
        driver.find_element_by_xpath(Xpath_KaisyaName).send_keys(TokuiNum) # 得意先番号

        # 発注番号
        Xpath_HaisouSakiSelect = '/html/body/div[1]/div/div/form/section[6]/div/table/tbody/tr/td/input'
        items = driver.find_element_by_xpath(Xpath_HaisouSakiSelect).send_keys(KobetsuBanGo)
    
        # コンソールに中も詳細を載せる。
        TokuiSakiDataDic = TokuiSakiData(TokuiNum) # DB接続
        PostCodeA = str(TokuiSakiDataDic['郵便番号'][:3])
        PostCodeB = str(TokuiSakiDataDic['郵便番号'][4:])
        TokuiName = jaconv.h2z(TokuiSakiDataDic['得意先名']) # h2z 半角to全角
        PhoneNo = TokuiSakiDataDic['電話番号']

        print("(" + str(int(TokuiNum)) + ")",TokuiName)
        print("郵便 :", TokuiSakiDataDic['郵便番号'])
        print("住所 :", TokuiSakiDataDic['住所'])
        address = TokuiSakiDataDic['住所']
        print("★ TEL:" + PhoneNo)
        print("★ 個別番号: ",str(KobetsuBanGo),"\n")

        # 郵便番号 まえ3桁入力
        Xpath_YubinA = "/html/body/div[1]/div/div/form/section[3]/div/table[2]/tbody[2]/tr[4]/td/input[1]"
        driver.find_element_by_xpath(Xpath_YubinA).send_keys(PostCodeA)

        # 郵便番号 うしろ4桁入力
        Xpath_YubinB = "/html/body/div[1]/div/div/form/section[3]/div/table[2]/tbody[2]/tr[4]/td/input[2]"
        driver.find_element_by_xpath(Xpath_YubinB).send_keys(PostCodeB)

        time.sleep(2) # 郵便番号入力後の自動表示間のインターバル時間

        # 番地枠のデータ取得する。 東京都杉並区 "神田"←　～
        Xpath_BanChi = "/html/body/div[1]/div/div/form/section[3]/div/table[2]/tbody[2]/tr[7]/td/input"
        BanChi = driver.find_element_by_xpath(Xpath_BanChi).get_attribute("value")

        # 番地枠の住所で得意先住所を分割する。
        address_result = address.split(BanChi)

        # 後半住所入力
        Xpath_AddressC = "/html/body/div[1]/div/div/form/section[3]/div/table[2]/tbody[2]/tr[8]/td/input"
        driver.find_element_by_xpath(Xpath_AddressC).send_keys(jaconv.h2z(address_result[-1]))

        # 電話番号 ハイフンが足りなかったり無かったりするとウェブ上でエラーになる
        Xpath_PhoneNo = '/html/body/div[1]/div/div/form/section[3]/div/table[2]/tbody[2]/tr[9]/td/input'
        driver.find_element_by_xpath(Xpath_PhoneNo).send_keys(PhoneNo)

        #配送先 担当者の枠に得意先様名を入れる。
        Xpath_TantoSyaName = "/html/body/div[1]/div/div/form/section[3]/div/table[2]/tbody[2]/tr[3]/td/input"
        driver.find_element_by_xpath(Xpath_TantoSyaName).send_keys(TokuiName) # 得意先様名
        
        CheckAddress = input("配送先住所を確認してください\n\n\n") # 得意先住所が郵便番号で出る住所に従っていないため確認後エンター
        print("----------------------------------------------------------------------------")

        # 早く遷移しすぎるとエラーになる。
        Xpath_KakuninButton = '//*[@id="__js-submit"]'
        driver.find_element_by_xpath(Xpath_KakuninButton).click() # 注文"確認"画面へ

        time.sleep(2) # コケたので2秒インターバル

        Xpath_TyumonKakutei = '//*[@id="orderForm"]/section[6]/button/span'
        driver.find_element_by_xpath(Xpath_TyumonKakutei).click() # 注文"確定"画面へ

        time.sleep(2) # コケたので2秒インターバル


if __name__ == "__main__":

    FileMake(ExcelFile, CsvFileName)          # "tmp_SmileBS_修正登録ファイル.xls"から"tmp_Web_登録用データ.csv"作成
    HattyuList = HattyuDataCsv(CsvFileName)   # 発注用に作成したデータをリスト化
    Login(login_page_url,login_id,login_pass) # ログイン
    SyoHinPageData(HattyuList)      # 注文個数入力から発注確定画面まで
