// GoogleAppScript用コード
//
// Gmailに届いた注文サイトからの注文確認メールを保存していく。
// 主に注文履歴を調べる用。
// 
// Function一覧
// function MailSave()    注文確認メールの内容をGoogleスプレッドシートに記述。
// function CreateXlsx()  更新がなくてもスプレッドシートをグーグルドライブにエクセル形式でエクスポート。
// function FileRemove()  エクスポートをしたエクセルを指定フォルダへ移動する。
// 
// トリガーで処理時間を設定する。
// function MailSave()    毎日 AM0:00~AM1:00
// function CreateXlsx()  毎日 AM1:00~AM2:00
// function FileRemove()  毎日 AM2:00~AM3:00
// 

function MailSave() {

    //メールアドレス + 本日の日付 + 翌日の日付でメール検索ワードを作成していく。
    var now = new Date();
    var today = Utilities.formatDate( now, 'Asia/Tokyo', 'yyyy/M/d');
    var NextDay = new Date(now.getYear(), now.getMonth(), now.getDate() + 1); //+1日 翌日の日付
    var tomorrow = Utilities.formatDate( NextDay, 'Asia/Tokyo', 'yyyy/M/d');
    var SearchWord = "shopmaster@*****.jp after:" + today + " before:" + tomorrow; // 検索ワード
    //Logger.log(SearchWord);
    var myThreads = GmailApp.search(SearchWord, 0, 100);
    var myMsgs = GmailApp.getMessagesForThreads(myThreads);
    
    //Logger.log(myThreads);
    //Logger.log(myMsgs);
    //Logger.log(myMsgs[0].length) // メッセージの数
    
    if (myMsgs.length) { //メッセージ0件で処理しない

      for ( var threadIndex = 0 ; threadIndex < myMsgs[0].length ; threadIndex++ ) {
        
        //最終行を取得
        var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('注文メール');
        var LastRow = sheet.getLastRow();
        
        var DataList = [];
        
        // メールから本文全文を抜き出す
        var mailBody = myMsgs[0][threadIndex].getPlainBody();
        
        // メール本文から情報を抽出する。
        var ChuMonBi = mailBody.match(/注文日時：.*/);
        ChuMonBi = ChuMonBi.toString().replace("注文日時：", "");
        DataList.push(ChuMonBi);
        
        var KobeTsuBanGo = mailBody.match(/発注番号：[0-9]+/);
        KobeTsuBanGo = KobeTsuBanGo.toString().replace("発注番号：", "");
        DataList.push(KobeTsuBanGo);
        
        var TyakuBi = mailBody.match(/セット名：.*/);
        TyakuBi = TyakuBi.toString().replace("セット名：[G7]", "");
        TyakuBi = TyakuBi.toString().replace("着荷商品", "");
        DataList.push(TyakuBi);
        
        var ChuMonSu = mailBody.match(/注文数：[0-9]+/);
        ChuMonSu = ChuMonSu.toString().replace("注文数：", "");
        DataList.push(ChuMonSu);
        
        var TokuiSakiBanGo = ""
        //メールに「配送先 会社名：」メッセージが無いエラーが出たので修正
        try {
          TokuiSakiBanGo = mailBody.match(/配送先 会社名：[0-9]+/);
          TokuiSakiBanGo = TokuiSakiBanGo.toString().replace("配送先 会社名：", "");
          DataList.push(TokuiSakiBanGo);
        }catch(e){
          TokuiSakiBanGo = "";
          DataList.push(TokuiSakiBanGo);
        }
        
        var TokuiSaki = mailBody.match(/配送先 担当者：.*/);
        TokuiSaki = TokuiSaki.toString().replace("配送先 担当者：", "");
        TokuiSaki = TokuiSaki.toString().replace(" 様", "");
        DataList.push(TokuiSaki);
        
        var YuBinBanGo = mailBody.match(/配送先 郵便番号：.*/);
        YuBinBanGo = YuBinBanGo.toString().replace("配送先 郵便番号：", "");
        DataList.push(YuBinBanGo);
        
        var HaiSouSaki = mailBody.match(/配送先 住所：.*/);
        HaiSouSaki = HaiSouSaki.toString().replace("配送先 住所：", "");
        DataList.push(HaiSouSaki);
        
        var DenWaBanGo = mailBody.match(/配送先 電話番号：.*/);
        DenWaBanGo = DenWaBanGo.toString().replace("配送先 電話番号：", "");
        DataList.push(DenWaBanGo);
        
        var GoKeiKinGaKu = mailBody.match(/小計：.*/);
        GoKeiKinGaKu = GoKeiKinGaKu.toString().replace("小計：", "");
        GoKeiKinGaKu = GoKeiKinGaKu.toString().replace(",", "");
        GoKeiKinGaKu = GoKeiKinGaKu.toString().replace("円", "");
        DataList.push(GoKeiKinGaKu);
        
        var TyuMonBanGo = mailBody.match(/注文番号：[0-9]+/);
        TyuMonBanGo = TyuMonBanGo.toString().replace("注文番号：", "");
        DataList.push(TyuMonBanGo);
        
        var ShoHinName = mailBody.match(/商品名：.*/);
        ShoHinName = ShoHinName.toString().replace("商品名：", "");
        DataList.push(ShoHinName);
        
        //Logger.log(mailBody) //メール内容
        //Logger.log(DataList)
        
        //メールメッセージのリストをスプレッドにセットしていく
        for ( var i = 0 ; i <= DataList.length - 1 ; i++ ) {
          sheet.getRange(LastRow+1, i+1).setValue(DataList[i]);
        }
        
      }
      
    }else{
      Logger.log("データなし")
    } //if
    
  }
  
  
  //SpreadsheetをExcelファイルに変換してドライブに保存、Fileを返す
  function CreateXlsx() {
    
    var spreadsheet_id = "***************************************";
    var new_file;
    var url = "https://docs.google.com/spreadsheets/d/" + spreadsheet_id + "/export?format=xlsx";
    var options = {
      method: "get",
      headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
      muteHttpExceptions: true
    };
    var res = UrlFetchApp.fetch(url, options);
    if (res.getResponseCode() == 200) {
      var ss = SpreadsheetApp.openById(spreadsheet_id);
      //new_file = DriveApp.createFile(res.getBlob()).setName(ss.getName() + ".xlsx"); //そのままのファイル名を使用する場合
      new_file = DriveApp.createFile(res.getBlob()).setName("WEB注文履歴.xlsx");
    }
    
    //メール送信
    var mail  = "*****@*****.co.jp"
    var title = "’エクセル'出力 完了"
    var body  = ""
    GmailApp.sendEmail(mail,title,body);

  }
  
  
  function FileRemove() {
    
    //ファイル移動前に古いファイルを消す
    var FileInFolder = DriveApp.getFolderById("***************************************");  //ディレクトリのIDを入力
    var file = FileInFolder.getFilesByName("WEB注文履歴.xlsx").next();
    FileInFolder.removeFile(file);
  
    //var file = DriveApp.getFileById(file_id);
    //DriveApp.getFolderById(folder_id).addFile(file);
    //DriveApp.getRootFolder().removeFile(file);
  
    //移動「前」のディレクトリ取得
    var INPUT_dir = DriveApp.getRootFolder(); //ディレクトリのIDを入力
    
    //移動「前」のファイル名
    var INPUT_file_name = "WEB注文履歴.xlsx"; //移動させたいファイルの名前を入力
    
    //移動「後」のディレクトリ取得
    var OUTPUT_dir = DriveApp.getFolderById("***************************************");  //ディレクトリのIDを入力
    
    //ファイルオブジェクトの取得
    var file = INPUT_dir.getFilesByName(INPUT_file_name).next();
    
    //ファイルの移動
    OUTPUT_dir.addFile(file);
    INPUT_dir.removeFile(file);
  
    //メール送信
    var mail  = "*****@*****.co.jp"
    var title = "'エクセル'移動 完了"
    var body  = ""
    GmailApp.sendEmail(mail,title,body);
    
  }
  
  
  