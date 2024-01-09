//DBからデータを取得する(出力頻度も考慮)
function inputDbDataKodoshien_() {
  //一覧表シートを開く
  const sheet = SpreadsheetApp.openById(ichiranhyouKodoshien).getSheetByName('データベースサンプル');
  const value = sheet.getDataRange().getValues();
  
  //サービス日時を参照して翌月のデータのみを取得する
  //サービス日時は4列目

  //翌月の年月データを取得
  const nextMonth = `${nextYearMonth_().year}/${nextYearMonth_().month}`

  const valueList =[]


  //0はヘッダーのため1から
  for(let i=1;i<value.length;i++){
    if(value[i][4].toLocaleString('ja-JP').slice(0,6)===nextMonth && value[i][8]==='行動支援'){
      valueList.push(value[i])
    }
  }

  //Logger.log(valueList)
  return valueList
}

//スプシを特定する(IDを渡す)
function connectionJissekiSheetKodoshien_(){

  //DB上から利用者IDを取得する(1列目)
  const userId = inputDbDataKodoshien_()[0][1];

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
    for(let k=0;k<inputDbDataKodoshien_().length;k++){
      if(inputDbDataKodoshien_()[k][1]===userIdColumnValuesList[j][0]){
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

function decisionSheetKodoshien(){
  //ターゲットとなるスプシを指定
  for (let j=0; j<connectionJissekiSheetKodoshien_().length; j++){

    //スプシを開く
    const ss = SpreadsheetApp.openById(connectionJissekiSheetKodoshien_()[j]);

    //選択したスプシのシート名を取得して実績票（翌月＋該当サービス区分）を取り出す
    const sheetList = ss.getSheets();

    //DBから取得したサービス区分を参照
    for (let k=0;k<inputDbDataKodoshien_().length;k++){
   
      //DBのサービス区分に該当するシートを開く
      const dbServiceName = inputDbDataKodoshien_()[k][8];
      //Logger.log(dbServiceName)//重度訪問介護

      for (let i=0; i<sheetList.length;i++){
        const sheetName = sheetList[i].getName();
        if(sheetName.includes(`${nextYearMonth_().year}年${nextYearMonth_().month}月_行動支援実績票`)){
          //シート名、サービス日、計画開始時間、サービス開始・終了時間
          //sheetNameList.push([sheetName,inputDbData_()[k][4],inputDbData_()[k][5],inputDbData_()[k][6],inputDbData_()[k][7]])
          const sheet = ss.getSheetByName(sheetName);

          //dbのサービス日と計画開始時間を取得
          const dbServiceDate = inputDbDataKodoshien_()[k][4].getDate()
          const dbPlanStartTime = `${inputDbDataKodoshien_()[k][6].getHours()},${inputDbDataKodoshien_()[k][6].getMinutes()}`
          const dbCancelMessage = inputDbDataKodoshien_()[k][28];

          const sheetTimes = sheet.getRange("H11:H41").getValues();
          const sheetDate = sheet.getRange("B11:B41").getValues();

          for(let l=0;l<sheetTimes.length;l++){
            if(sheetTimes[l][0]!==""){
              const sheetTime = `${sheetTimes[l][0].getHours()},${sheetTimes[l][0].getMinutes()}`
              if(dbPlanStartTime===sheetTime && dbServiceDate===sheetDate[l][0] && dbCancelMessage === ""){
                sheet.getRange(`T${l+11}`).setValue(inputDbDataKodoshien_()[k][5]);
                sheet.getRange(`X${l+11}`).setValue(inputDbDataKodoshien_()[k][7]);
              }else if(dbPlanStartTime===sheetTime && dbServiceDate===sheetDate[l][0] && dbCancelMessage !== ""){
                sheet.getRange(`AU${l+11}`).setValue(inputDbDataKodoshien_()[k][28]);
              }
            }
          }          
        }
      }
    }
  }
}