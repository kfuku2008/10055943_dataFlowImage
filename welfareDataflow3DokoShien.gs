//DBからデータを取得する(出力頻度も考慮)
function inputDbDataDokoShien_() {
  //一覧表シートを開く
  const sheet = SpreadsheetApp.openById(ichiranhyouDokoShien).getSheetByName('データベースサンプル');
  const value = sheet.getDataRange().getValues();
  
  //サービス日時を参照して翌月のデータのみを取得する
  //サービス日時は4列目

  //翌月の年月データを取得
  const nextMonth = `${nextYearMonth_().year}/${nextYearMonth_().month}`

  const valueList =[]


  //0はヘッダーのため1から
  for(let i=1;i<value.length;i++){
    if(value[i][4].toLocaleString('ja-JP').slice(0,6)===nextMonth && value[i][8]==='同行支援'){
      valueList.push(value[i])
    }
  }

  //Logger.log(valueList)
  return valueList
}

//スプシを特定する(IDを渡す)
function connectionJissekiSheetDokoShien_(){

  //DB上から利用者IDを取得する(1列目)
  const userId = inputDbDataDokoShien_()[0][1];

  //Logger.log(userId)

  //利用者リストを開く
  const userListSheet = SpreadsheetApp.openById(userIdList).getSheetByName('利用者リスト');

  const userIdColumnValues = userListSheet.getRange("C2:C10").getValues();

  const userIdColumnValuesList = []

  //利用者番号と行数を二次元配列に格納
  for(let i=0 ; i<userIdColumnValues.length; i++){
    if(userIdColumnValues[i][0] !==""){
      const numRows = i+2
      userIdColumnValuesList.push([userIdColumnValues[i][0],numRows])
    }
  }

  const targetIdList = []

  //DB上で取得したIDとリストで取得したIDで一致した場合、そのスプシIDを取得する
  for(let j =0; j<userIdColumnValuesList.length;j++){
    for(let k=0;k<inputDbDataDokoShien_().length;k++){
      if(inputDbDataDokoShien_()[k][1]===userIdColumnValuesList[j][0]){
        const targetRow = userIdColumnValuesList[j][1] ;
        const targetId = userListSheet.getRange(`B${targetRow}`).getValue();
        targetIdList.push(targetId)
      }
    }  
  }

  const targetUniqueIdList = Array.from(new Set(targetIdList))

  return targetUniqueIdList

}

//★書き込むシートを特定・サービス区分を紐づける(トリガー対象関数)
//ランタイム対策コード設置

function decisionSheetDokoShien(){
  //ターゲットとなるスプシを指定
  for (let j=0; j<connectionJissekiSheetDokoShien_().length; j++){

    //スプシを開く
    const ss = SpreadsheetApp.openById(connectionJissekiSheetDokoShien_()[j]);

    //選択したスプシのシート名を取得して実績票（翌月＋該当サービス区分）を取り出す
    const sheetList = ss.getSheets();

    //DBから取得したサービス区分を参照
    for (let k=0;k<inputDbDataDokoShien_().length;k++){
      for (let i=0; i<sheetList.length;i++){
        const sheetName = sheetList[i].getName();
        if(sheetName.includes(`${nextYearMonth_().year}年${nextYearMonth_().month}月_同行支援実績票`)){
          const sheet = ss.getSheetByName(sheetName);
          //dbのサービス日と計画開始時間を取得
          const dbServiceDate = inputDbDataDokoShien_()[k][4].getDate()
          const dbPlanStartTime = `${inputDbDataDokoShien_()[k][6].getHours()},${inputDbDataDokoShien_()[k][6].getMinutes()}`
          const dbCancelMessage = inputDbDataDokoShien_()[k][28];

          const sheetTimes = sheet.getRange("H11:H41").getValues();
          const sheetDate = sheet.getRange("B11:B41").getValues();

          for(let l=0;l<sheetTimes.length;l++){
            if(sheetTimes[l][0]!==""){
              const sheetTime = `${sheetTimes[l][0].getHours()},${sheetTimes[l][0].getMinutes()}`
              if(dbPlanStartTime===sheetTime && dbServiceDate===sheetDate[l][0] && dbCancelMessage === ""){
                sheet.getRange(`T${l+11}`).setValue(inputDbDataDokoShien_()[k][5]);
                sheet.getRange(`X${l+11}`).setValue(inputDbDataDokoShien_()[k][7]);
              }else if(dbPlanStartTime===sheetTime && dbServiceDate===sheetDate[l][0] && dbCancelMessage !== ""){
                sheet.getRange(`AU${l+11}`).setValue(inputDbDataDokoShien_()[k][28]);
              }
            }
          }          
        }
      }
    }
  }
}