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
    timeStamp: 0,         // タイムスタンプ
    userName: 0,          // 契約者ご本人氏名
    mailAddress: 0,       // メールアドレス
    companyName: 0,       // 会社名
    zipCode: 0,           // 郵便番号
    address: 0,           // 住所
    phoneNumber: 0,       // 電話番号
    mobilePhoneNumber: 0, // 携帯電話番号
    corporateAddress: 0,  // 法人住所登記
    post: 0,              // 専用ポスト
    mailTransfer: 0,      // 郵便物転送
    callForwarding: 0,    // 電話転送サービス
    locker: 0,            // お客様専用ロッカー
    startDate: 0,         // 利用開始日
    contractPeriod: 0,    // 契約期間
    status: 0,            // ステータス
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
      case '会社名':
        columnIndex.companyName = i;
        break;
      case '郵便番号':
        columnIndex.zipCode = i;
        break;
      case '住所':
        columnIndex.address = i;
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
      case '電話転送サービス（月額）':
        columnIndex.callForwarding = i;
        break;
      case 'お客様専用ロッカー（月額）':
        columnIndex.locker = i;
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

      GmailApp.sendEmail('contact@harbors.sh', title, body, option);
      console.log(body.toString())
      dataList[i][columnIndex.status] = '確認メール送信済'
    }
  }
  InputResultSheet.getRange(1, 1, dataList.length ,dataList[0].length).setValues(dataList); //データをシートに出力
  console.log('process finish');
}

