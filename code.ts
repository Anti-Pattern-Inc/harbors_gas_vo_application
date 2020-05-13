// https://github.com/Anti-Pattern-Inc/harbors_gas_vo_application

function main() {
  console.log('process start');

  // 連携しているスプレッドシートを取得
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 「契約プラン選択フォームの回答内容」シートを取得
  const InputResultSheet = spreadsheet.getActiveSheet(); 

  // シートデータを二次元配列で取得
  let dataList = InputResultSheet.getDataRange().getValues();

  // ヘッダー名に対応する列番号を保持する配列
  let columnIndex = {  
    timeStamp: 0,             // タイムスタンプ
    userName: 0,              // 契約者ご本人氏名
    mailAddress: 0,           // メールアドレス
    invoiceMailAddress: 0,    // 請求先メールアドレス
    invoiceName: 0,           // 請求先氏名
    contractType: 0,          // 契約種別
    companyName: 0,           // 会社名,
    representativeName: 0,    // 代表者名
    zipCode: 0,               // 郵便番号
    address: 0,               // 都道府県・市区町村・番地
    addressDetails: 0,        // 建物名・部屋・番号など
    phoneNumber: 0,           // 電話番号
    mobilePhoneNumber: 0,     // 携帯電話番号
    corporateAddress: 0,      // 法人住所登記
    post: 0,                  // 専用ポスト
    mailTransfer: 0,          // 郵便物転送
    mailTransferSize: 0,      // 郵便物転送 サイズ
    mailTransferAddress: 0,   // 郵便物転送住所
    callForwarding: 0,        // 電話転送サービス
    callForwardingNumber: 0,  // 電話転送先電話番号
    locker: 0,                // お客様専用ロッカー
    lockerSize: 0,            // お客様専用ロッカーサイズ
    startDate: 0,             // 利用開始日
    contractPeriod: 0,        // 契約期間
    status: 0,                // ステータス
  };

  // オペレーターに送信する情報を持つの列番号を取得　←列がずれても処理に影響が出ないようにするため
  for(let i=1; i<dataList[0].length; i++) {
    // ヘッダー名ごとに列番号を専用配列に格納
    switch(dataList[0][i]) {
      case 'タイムスタンプ':
        columnIndex.timeStamp = i;
        break;
      case '契約者ご本人氏名':
        columnIndex.userName = i;
        break;
      case 'メールアドレス':
        columnIndex.mailAddress = i;
        break;
      case '請求先メールアドレス':
        columnIndex.invoiceMailAddress = i;
        break;
      case '請求先氏名':
        columnIndex.invoiceName = i;
        break;
      case '契約種別':
        columnIndex.contractType = i;
        break;
      case '会社名':
        columnIndex.companyName = i;
        break;
      case '代表者名':
        columnIndex.representativeName = i;
        break;
      case '郵便番号':
        columnIndex.zipCode = i;
        break;
      case '都道府県・市区町村・番地':
        columnIndex.address = i;
        break;
      case '建物名・部屋・番号など':
        columnIndex.addressDetails = i;
        break;
      case '電話番号（任意）':
        columnIndex.phoneNumber = i;
        break;
      case '携帯電話番号':
        columnIndex.mobilePhoneNumber = i;
        break;
      case '法人住所登記（月額）':
        columnIndex.corporateAddress = i;
        break;
      case '専用ポスト（月額）':
        columnIndex.post = i;
        break;
      case '郵便物転送（月額）':
        columnIndex.mailTransfer = i;
        break;
      case '郵便物転送 サイズ（月額）':
        columnIndex.mailTransferSize = i;
        break;
      case '郵便物転送　転送先住所':
        columnIndex.mailTransferAddress = i;
        break;
      case '電話転送サービス（月額）':
        columnIndex.callForwarding = i;
        break;
      case '電話転送サービス 転送先電話番号':
        columnIndex.callForwardingNumber = i;
        break;
      case 'お客様専用ロッカー（月額）':
        columnIndex.locker = i;
        break;
      case 'お客様専用ロッカー サイズ（月額）':
        columnIndex.lockerSize = i;
        break;
      case '利用開始日':
        columnIndex.startDate = i;
        break;
      case '契約期間':
        columnIndex.contractPeriod = i;
        break;
      case 'ステータス':
        columnIndex.status = i;
        break;
    }
  }
  console.log('パラメータ数：%d',dataList.length);
  for(let i=1; i<dataList.length; i++) {
    // ステータスが空欄の場合、オペレータへのメール送信を行う
    if(dataList[i][columnIndex.status] == "" && dataList[i][columnIndex.mailAddress] != ""){
      console.log("%d 行目の回答情報を処理", i+1);
      // メールオプション
      const option = {from: 'contact@harbors.sh', name: 'バーチャルオフィス申込みフォーム'};
      // 件名
      const title = "バーチャルオフィス申込み入力完了通知";
      //　予約完了メールのテンプレートをドキュメントより取得
      const document = DocumentApp.openById('1MQYvINz-DP7YbPQMyAIGSv7xklljJ2gQa3lx8Gz5oN0'); //ドキュメントをIDで取得
      const bodyTemplate = document.getBody().getText();

      // タイムスタンプをセット
      let body = bodyTemplate.replace("%timeStamp%", dataList[i][columnIndex.timeStamp].toLocaleDateString());
      // ご契約者ご本人氏名をセット
      body = body.replace("%userName%", dataList[i][columnIndex.userName]);
      // メールアドレスをセット
      body = body.replace("%mailAddress%", dataList[i][columnIndex.mailAddress]);
      // 会社名をセット
      body = body.replace("%companyName%", dataList[i][columnIndex.companyName]);
      // 郵便番号をセット
      body = body.replace("%zipCode%", dataList[i][columnIndex.zipCode]);
      // 住所をセット
      body = body.replace("%address%", dataList[i][columnIndex.address]);
      // 電話番号をセット
      body = body.replace("%phoneNumber%", dataList[i][columnIndex.phoneNumber]);
      // 携帯電話番号をセット
      body = body.replace("%mobilePhoneNumber%", dataList[i][columnIndex.mobilePhoneNumber]);
      // 法人住所登記（月額）プランをセット
      body = body.replace("%corporateAddress%", dataList[i][columnIndex.corporateAddress]);
      // 専用ポスト（月額）プランをセット
      body = body.replace("%post%", dataList[i][columnIndex.post]);
      // 郵便物転送（月額）プランをセット
      body = body.replace("%mailTransfer%", dataList[i][columnIndex.mailTransfer]);
      // 電話転送サービス（月額）プランをセット
      body = body.replace("%callForwarding%", dataList[i][columnIndex.callForwarding]);
      // お客様専用ロッカー（月額）プランをセット
      body = body.replace("%locker%", dataList[i][columnIndex.locker]);
      // 契約期間をセット
      body = body.replace("%contractPeriod%", dataList[i][columnIndex.contractPeriod]);
      // 利用開始日をセット
      body = body.replace("%startDate%", dataList[i][columnIndex.startDate].toLocaleDateString());

      // GmailApp.sendEmail('contact@harbors.sh', title, body, option);
      try {
        try{        
          // slack通知
          // TODO slack 通知の正しいメッセージに変更
          postMessageToContactChannel('<!channel>「アヤンテスト」に申し込みがありました。');
        }catch(error){
          throw new Error('slack送信エラー(' + error + ')');
        }
  
        try{        
          //申し込みお礼のメール送信
          // TODO メール宛先を申請者に変更
          sendCompleteMail('a.hayes@anti-pattern.co.jp');
        }catch(error){
          throw new Error('メール送信エラー(' + error + ')');
        }
      } catch(error) {
        postMessageToContactChannel('<!channel>「アヤンテスト」の申込でエラーが発生しました。\n```エラー内容:' + error + '```');
        dataList[i][columnIndex.status] = 'エラー発生しました'
      }

      console.log(body.toString())
      dataList[i][columnIndex.status] = '申し込み通知済'
    }
  }
  InputResultSheet.getRange(1, 1, dataList.length ,dataList[0].length).setValues(dataList); //データをシートに出力
  console.log('process finish');
}

/** 
 * slackのチャンネルにメッセージを投稿する
 * @param  {string} message 投稿メッセージ
 * @return {void}
 */
function postMessageToContactChannel(message: string): void {
  // #contantへのwebhook URLを取得
  // TODO slack-testからcontactのWebhookURLに切り替える
  const webhookURL = PropertiesService.getScriptProperties().getProperty('WEBHOOK_URL');
  // 投稿に必要なデータを用意
  const jsonData =
  {
      "text" : message  // 投稿メッセージ
  };
  // JSON文字列に変換
  const payload = JSON.stringify(jsonData);

  // 送信オプションを用意
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: "post",
    contentType: "application/json",
    payload: payload
  }
  
  UrlFetchApp.fetch(webhookURL, options);
}

/** 
 * 予約完了メールを送信する
 * @param {string} mailAddress 送信先アドレス
 * @param {string} contactName 予約者者名
 * @param {string} eventName   予約イベント名
 * @param {string} visitDate   予約日利用日
 * @param {string} visitTime   利用開始時間
 * @return void
 */
function sendCompleteMail(mailAddress :string) :void{

  //定義からテンプレートID取得
  const templateId = PropertiesService.getScriptProperties().getProperty('COMPLETE_MAIL_TEMPLATE');

  // メールオプション
  const option = {from: 'contact@harbors.sh', name: 'HarborS運営スタッフ'};
  // 件名
  const title = "申込のお知らせ";
  //　予約完了メールのテンプレートをドキュメントより取得
  const document = DocumentApp.openById(templateId);
  const bodyTemplate = document.getBody().getText();
  // 氏名をセット
  // TODO 正しい値をメール文面に切り替える
  let body = bodyTemplate;
  // let body = bodyTemplate.replace("%contactName%", contactName);
  // // イベントをセット
  // body = body.replace("%eventName%", eventName);
  // // 予約日をセット
  // body = body.replace("%visitDate%", visitDate);
  // // 予約時間をセット
  // body = body.replace("%visitTime%", visitTime);

  GmailApp.sendEmail(mailAddress, title, body, option);

}