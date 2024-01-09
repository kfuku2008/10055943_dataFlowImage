//行動援護・同行支援・行動支援
function formToDatabaseKodo_(){
  //回答データの取得
  const form = FormApp.openById(formKodoId);

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
        answerList.push(answer);
      }
    }else{
      continue;
    }
    answersList.push(answerList);

  }

  //Logger.log(answersList)

  return answersList;
}

//回答データを書き込む関数

//DB書き込み（同行・行動・移動）
function answertoDataBaseDokoKodoIdo(){
  const dataList = formToDatabaseKodo_();
  
  const ss = SpreadsheetApp.openById(ichiranhyouKodo);
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