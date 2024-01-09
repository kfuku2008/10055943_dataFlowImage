//DBからデータを取得する(出力頻度も考慮)
function inputDbDataJudo_() {
  //一覧表シートを開く(シート名はデータベースサンプル)
  const sheet = SpreadsheetApp.openById(ichiranhyouJudo).getSheetByName('データベースサンプル');
  const value = sheet.getDataRange().getValues();
  
  //サービス日時を参照して翌月のデータのみを取得する
  //サービス日時は4列目

  //翌月の年月データを取得
  const nextMonth = `${nextYearMonth_().year}/${nextYearMonth_().month}`

  const valueList =[]


  //0はヘッダーのため1から
  for(let i=1;i<value.length;i++){
    if(value[i][4].toLocaleString('ja-JP').slice(0,6)===nextMonth && value[i][8]==='重度訪問介護'){
      valueList.push(value[i])
    }
  }


  return valueList
}

//スプシを特定する(IDを渡す)(時間がかかりそう、、)
function connectionJissekiSheetJudo_(){

  //利用者リストを開く(シート名は利用者リスト)
  const userListSheet = SpreadsheetApp.openById(userIdList).getSheetByName('利用者リスト');

  //修正必要あり(最終行で取れるように加工)
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
    for(let k=0;k<inputDbDataJudo_().length;k++){
      if(inputDbDataJudo_()[k][1]===userIdColumnValuesList[j][0]){
        const targetRow = userIdColumnValuesList[j][1] ;
        const targetId = userListSheet.getRange(`B${targetRow}`).getValue();
        targetIdList.push(targetId)
      }
    }  
  }

  const targetUniqueIdList = Array.from(new Set(targetIdList))

  return targetUniqueIdList
}

//書き込むシートを特定・サービス区分を紐づける
function decisionSheetKyotakuJudo(){

  //ターゲットとなるスプシを指定
  for (let j=0; j<connectionJissekiSheetJudo_().length; j++){

    //スプシを開く
    const ss = SpreadsheetApp.openById(connectionJissekiSheetJudo_()[j]);

    //選択したスプシのシート名を取得して実績票（翌月＋該当サービス区分）を取り出す
    const sheetList = ss.getSheets();

    //DBから取得したサービス区分を参照
    for (let k=0;k<inputDbDataJudo_().length;k++){
   
      //DBのサービス区分に該当するシートを開く
      for (let i=0; i<sheetList.length;i++){
        const sheetName = sheetList[i].getName();
        if(sheetName.includes(`${nextYearMonth_().year}年${nextYearMonth_().month}月_重度訪問介護実績票`)){
          //シート名、サービス日、計画開始時間、サービス開始・終了時間
          //sheetNameList.push([sheetName,inputDbData_()[k][4],inputDbData_()[k][5],inputDbData_()[k][6],inputDbData_()[k][7]])
          const sheet = ss.getSheetByName(sheetName);

          //dbのサービス日と計画開始時間を取得
          const dbServiceDate = inputDbDataJudo_()[k][4].getDate()
          const dbPlanStartTime = `${inputDbDataJudo_()[k][6].getHours()},${inputDbDataJudo_()[k][6].getMinutes()}`
          const dbCancelMessage = inputDbDataJudo_()[k][42];

          const sheetTimes = sheet.getRange("N12:N46").getValues();
          const sheetDate = sheet.getRange("D12:D46").getValues();

          for(let l=0;l<sheetTimes.length;l++){
            if(sheetTimes[l][0]!==""){
              const sheetTime = `${sheetTimes[l][0].getHours()},${sheetTimes[l][0].getMinutes()}`
              if(dbPlanStartTime===sheetTime && dbServiceDate===sheetDate[l][0] && dbCancelMessage === ""){
                sheet.getRange(`AB${l+12}`).setValue(inputDbDataJudo_()[k][5]);
                sheet.getRange(`AF${l+12}`).setValue(inputDbDataJudo_()[k][7]);
              }else if(dbPlanStartTime===sheetTime && dbServiceDate===sheetDate[l][0] && dbCancelMessage !== ""){
                sheet.getRange(`BJ${l+12}`).setValue(inputDbDataJudo_()[k][42]);
              }
            }
          }          
        }
      }
    }
  }
}
