function sendCancelAlertMail() {
  
  const ansKyotakuJudo = formToDatabaseKyotakuJudo_();
  const ansKodo = formToDatabaseKodo_();

  Logger.log(ansKyotakuJudo[0][42])
  Logger.log(ansKodo[0][28])

  //2つのフォームデータのいずれかからキャンセルがあった場合メールで送信

  const cancelList1 = []
  const canceList2 = []

  //居宅・重度
  for(let i=0; i<ansKyotakuJudo.length;i++){
    if(ansKyotakuJudo[i][42]!==""){
      cancelList1.push([ansKyotakuJudo[i][2],ansKyotakuJudo[i][4],ansKyotakuJudo[i][6],ansKyotakuJudo[i][8],ansKyotakuJudo[i][42]])
    }
  }

  //行動援護・同行支援・行動支援
  for(let j=0; j<ansKodo.length;j++){
    if(ansKodo[j][28]!==""){
      canceList2.push([ansKodo[j][2],ansKodo[j][4],ansKodo[j][6],ansKodo[j][8],ansKodo[j][28]])
    }
  }

  //2つの二次元配列を結合
  const canceList3 = cancelList1.concat(canceList2);

  Logger.log(canceList3)

  //メール送付の実行
  sendEmail_(cancelAlertMessage_(canceList3))
}

//キャンセル時の通知
//メッセージ内容「〇〇様〇月〇日〇時から開始予定の〇〇サービスでキャンセルがありました」
//value[i][2](利用者氏名),value[i][4](サービス日時),value[i][6](計画開始時刻),value[i][8](サービス区分),value[i][42]

//アラートメッセージの作成関数
function cancelAlertMessage_(cancelList){
  //キャンセルのもののみ出力される
  //	[[鈴木次郎, Sun Jan 21 00:00:00 GMT+09:00 2024, Sat Dec 30 00:00:00 GMT+09:00 1899, 居宅介護, キャンセルテスト], [山田花子, Tue Jan 09 00:00:00 GMT+09:00 2024, Sat Dec 30 05:00:00 GMT+09:00 1899, 居宅介護, キャンセルテスト]]

  const message = []

  for(let i=0; i<cancelList.length; i++){
    const target = new Date(cancelList[i][1]);
    const targetMonth = target.getMonth()+1;
    const targetDate = target.getDate();
    const targetTime = cancelList[i][2];

    const text = `■${cancelList[i][0]}様の${targetMonth}月${targetDate}日${targetTime}から開始予定の${cancelList[i][3]}でキャンセルがありました\n※キャンセル理由：${cancelList[i][4]}`

    message.push(text)
  }

  const formatMessage = message.join('\n');

  return formatMessage
}

//メール送付
function sendEmail_(data) {
  // Gmailでメールを送信する処理を追加
  //スクリプトプロパティに自分のメールアドレスを保存
  const recipient = PropertiesService.getScriptProperties().getProperty('MY_MAIL_ADDRESS'); // 送信先のメールアドレスを指定
  const today = new Date().toLocaleDateString('ja-JP');
  const subject = `キャンセル情報の通知(${today})`;//メールタイトルはここで変更
  const body = data;//メール本文の内容はここで変更


  // メールを送信
  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    body: body,
  });
}