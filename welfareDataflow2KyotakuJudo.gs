//居宅介護・重度訪問
function formToDatabaseKyotakuJudo_(){
  //回答データの取得
  const form = FormApp.openById(formKyotakuJudoId);

  //formの回答者数をリストに格納
  const responses = form.getResponses();
  const today = new Date().toDateString();

  //複数人回答用のリストを準備
  const answersList=[];

  //実行日に回答があったデータを取得
  for (let i =0;i<responses.length;i++){

    //1つ目(1人目)の回答データを取得
    let response = responses[i];

    //回答時間のタイムスタンプの取得
    let timeStamp = response.getTimestamp().toDateString();

    let answerList = [];


    if(timeStamp===today){

      answerList = [timeStamp];

      //各アイテムの質問と回答を取得
      let itemResponses = response.getItemResponses();

      for(const itemResponse of itemResponses){

        //アイテムの回答を取得
        const answer = itemResponse.getResponse();

        //questionList.push(question)
        answerList.push(answer)
      }
    }else{
      continue;
    }
    answersList.push(answerList)
  }

  //Logger.log(answersList)

  return answersList
}

//回答データを書き込む関数

//DB書き込み（居宅・重度）
function answertoDataBaseKyotakuJudo(){
  const dataList = formToDatabaseKyotakuJudo_();

  const ss = SpreadsheetApp.openById(ichiranhyouKyotakuJudo);
  const sheet = ss.getSheetByName('データベースサンプル');

  //最終行のデータを取得
  const lastRow = sheet.getLastRow()

  for(let i=0; i<dataList.length;i++){

    //複数選択時はリスト形式で格納されているので「A,B,C,,」形式で一つのセルに格納できるようにする
    const formatdataList = flattenArray_(dataList[i])

    Logger.log(formatdataList.length)

    // //列方向に書き出す場合はfor文で書く
    for (let j=0; j<formatdataList.length;j++){
      sheet.getRange(lastRow+1+i,j+1).setValue(formatdataList[j])
    }
  }
}


//複数入力項目変換関数(/で区切っている)
function flattenArray_(originalArray) {
  const flattenedArray = [];
  
  for (let i = 0; i < originalArray.length; i++) {
    const item = originalArray[i];
    
    if (Array.isArray(item)) {
      flattenedArray.push(item.join('/'));
    } else {
      flattenedArray.push(item);
    }
  }
  
  return flattenedArray;
}


//質問リストをヘッダーにする関数

function questionToHeader_(){
  //const headerList = formToDatabaseKyotakuJudo_().question;
  const headerList = formToDatabaseKodo_().question;

  //ヘッダーリストを一覧表に書き込む
  //const kodoID = PropertiesService.getScriptProperties().getProperty('KYOTAKUJUDO_SHEET_ID')
  const kodoID = PropertiesService.getScriptProperties().getProperty('KODO_SHEET_ID')

  const ss = SpreadsheetApp.openById(kodoID);
  const sheet = ss.getSheetByName('データベースサンプル');

  //列方向に書き出す場合はfor文で書く
  for (let i=0; i<headerList.length;i++){
    sheet.getRange(1,i+1).setValue(headerList[i])
  }
}


