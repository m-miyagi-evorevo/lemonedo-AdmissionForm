//======================
// 入稿フォーム（https://forms.gle/vPaAshLH9MRen2kT6）から送信されたデータを取得し整形した状態で
// メーリングリスト（lemonedo@evorevo.co.jp）へメール送信する関数
// ※function名は必ず「OnformeEdit」でなければならない
//======================
function OnFormeEdit(e) {//参考ページ→https://vba-gas.info/gas-googleform-2
  try{
    //フォームから送信された情報を取得
    var TimeStamp      = e.namedValues['タイムスタンプ'];　//フォームからタイムスタンプを取得
    var Company        = e.namedValues['会社名をご記入ください。'];　//フォームから会社名を取得
    var StaffName      = e.namedValues['ご担当者様のご氏名をご記入ください。'];　//フォームから担当者様名を取得
    var PhoneNumber    = e.namedValues['ご担当者様へ連絡可能な電話番号をご記入ください（ハイフン無しで半角数値をご記入お願いします）'];　//担当者の電話番号を取得
    var ImgURL         = e.namedValues['ステッカー画像のアップロードをお願いします'];　//ドライブに格納されたステッカー画像のURLを取得
    var MovURL         = e.namedValues['QRコードの遷移先として流したい動画がある場合、動画をアップロードお願いします。動画の長さの推奨は30秒です。（最大10分まで可能。ファイルサイズが大きいとアップロードに時間がかかる場合があります。）'];　//ドライブに格納sれた動画のURLを取得
    console.log("Company: " + Company);
    console.log("StaffName: " + StaffName);
    console.log("PhoneNumber: " + PhoneNumber);
    console.log("ImgURL: " + String(ImgURL));
    console.log("MovURL: " + String(MovURL));
    
    // 前処理
    var SubmitTime     = Utilities.formatDate(new Date(TimeStamp), 'Asia/Tokyo','yyyyMMdd_HHmmss');// 入稿日時をyyyyMMdd_HHmmss形式に変換
    var FolderId       = '1vPGDvZ57YfuMs_z0IaGvCfaSFnQeQJSN';　//保存先フォルダのIDを指定
    var Folder         = DriveApp.getFolderById(FolderId);// 保存先フォルダを取得
    var ss             = SpreadsheetApp.getActiveSpreadsheet();// スプレッドシートを取得
    var sh             = ss.getSheetByName('フォームの回答 1');// 「フォームの回答 1」シートを取得
    var MaxRow         = sh.getLastRow();// 最終行を取得
    
    // 同名の会社名で過去に入稿があったか調査し、無ければ新たにフォルダを生成して、画像と動画を格納する処理
    var CompanyFolderExist = false;// 会社名のフォルダ存在フラグの初期値をfalseに設定
    var SubFolders = Folder.getFolders();// 入稿フォルダ内のサブフォルダを全て取得
    var SubFolderCnt = 0;// サブフォルダ数の初期値を設定（デフォルトで全画像が格納されるフォルダと全動画が格納されるフォルダが存在しているので−2とする）
    while(SubFolders.hasNext() && CompanyFolderExist == false){// SubFolderを一つずつ処理（会社名のサブフォルダが見つかった場合はループを抜ける）
      SubFolderCnt++;// SubFolderCntをカウントアップ
      var SubFolder = SubFolders.next();// 処理対象のフォルダを取り出す
      var SubFolderName = SubFolder.getName();// サブフォルダ名を取得
      if(SubFolderName.indexOf(Company + "_") != -1){// サブフォルダ名に企業名が含まれいている場合
        CompanyFolderExist = true;// 会社名のフォルダ存在フラグをtrueに変更
        var CompanyFolder = SubFolder;// CompanyFolderにSubFolderを代入する
      }
    }
    
    if(CompanyFolderExist == false){// 入稿会社名のフォルダが存在しなかった場合
      var SubFolderNo = ("0000" + (SubFolderCnt + 1)).slice(-4);// SubFolder数に+1して、4桁でゼロパンディングする
      var CompanyFolder = Folder.createFolder(SubFolderNo + "." + Company + "_" + SubmitTime);// 企業フォルダを生成（命名規則「フォルダNo.会社名_入稿日時」）
    }
    
    // 画像ファイルの情報取得とフォルダ移動処理
    var reg = /(?<=id=).*/;// 正規表現でURLからID部分を指定する内容を定義
    var ImgResult = String(ImgURL).match(reg);// 画像URLからID部分を抽出
    var ImgFileId = ImgResult[0];// ImgFileIdにファイルIDを格納
    var ImgFile = DriveApp.getFileById(ImgFileId);// 画像ファイルを取得
    var ImgFileName = ImgFile.getName();// ファイル名を取得
    var ImgExtensionExist = ImgFileName.lastIndexOf(".");// ファイル名に「.」が含まれるかファイル名の最後尾からチェック（存在しない場合は−1が返る）
    if(ImgExtensionExist != -1){// 拡張子が存在する場合
      var ImgExtension = ImgFileName.slice(ImgFileName.lastIndexOf(".") + 1);// ファイルの拡張子を取得
    }else{
      var ImgExtension = "拡張子なし"
      }
    
    var NewImgURL = ImgFile.makeCopy("画像" + Company + "_" + StaffName + SubmitTime + ImgExtension, CompanyFolder).getUrl();// 会社名フォルダにImgFileを生成しURLを取得
    
    // 動画ファイルの情報取得とフォルダ移動処理
    if(String(MovURL) != ""){// MovieURLが空白以外の場合（入稿がある場合）
      var reg = /(?<=id=).*/;// 正規表現でURLからID部分を指定する内容を定義
      var MovResult = String(MovURL).match(reg);// 画像URLからID部分を抽出
      var MovFileId = MovResult[0];// MovFileIdにファイルIDを格納
      var MovFile = DriveApp.getFileById(MovFileId);// 画像ファイルを取得
      var MovFileName = MovFile.getName();// ファイル名を取得
      var MovExtension = MovFileName.slice(MovFileName.lastIndexOf(".") + 1);// ファイルの拡張子を取得
      var NewMovURL = MovFile.makeCopy("動画" + Company + "_" + StaffName + SubmitTime + MovExtension, CompanyFolder).getUrl();// 会社名フォルダにMovFileを生成しURLを取得
      var MovMailBody = 
          '動画ファイル拡張子：' + MovExtension + '\n' + 
            '動画ファイルURL　：' + NewMovURL + '\n';
    }else{
      var MovFileName = "なし";
      var MovMailBody = '動画ファイル：なし\n';
    }
    
    // 入稿お知らせメールを送信
    GmailApp.sendEmail(
      'lemonedo@evorevo.co.jp', // 
      Company + ' より入稿がありました', 
      '入稿がありました。以下の内容を確認し、入稿データのチェックを開始してください。\n\n' + 
      '会社名　　　　　　：' + Company + '\n' + 
      '画像ファイル拡張子：' + ImgExtension + '\n' + 
      '画像ファイルURL　：' + NewImgURL + '\n' + 
      MovMailBody + 
      '企業用フォルダ\n' + CompanyFolder.getUrl() + 
      '\n\n内容に問題がなければスプレッドシートA列の承認ステータスを変更してください。\n\n' + ss.getUrl() + '#gid=1130205328&range=A' + MaxRow,
        {from: 'm-miyagi@evorevo.co.jp',name: '入稿お知らせメール'});
    
  }catch(error){
    GmailApp.sendEmail(
    'lemonedo@evorevo.co.jp', // 
    '★エラー発生_入稿ファイル処理', 
    '入稿ファイルの処理でエラーが発生しました。エラーの詳細を確認してください。\n\nエラー内容：' + error + 
    '\n\nGAS：\nhttps://script.google.com/a/evorevo.co.jp/d/1OtzbFJ5Hhiu52JdVxeLRBTE6aw8PNReZVfw1YmL3fgn3quAXQQaaLB7y/edit?uiv=2&mid=ACjPJvFgv5_KYLhCgTQuLUc2nRQODq82ttZzoTrae1kyp6IynABlo-TYc97WZa5WOBzgz02_o5nZLw823HFYJjsF2aUL4CqzpCTMLpu13oHQLbEvgvYpVOy6LnMEVFCsDBra9l4DYTJ7W95aeKeB3CKjlo40lYkAgzVMo5W952d-gKQdzZg4R_bQ4fxChg&splash=yes',
      {from: 'm-miyagi@evorevo.co.jp',name: '入稿お知らせメール'});
  }
}