//スプレッドシート③をシート②へ記載
function formDataOutput() {
  //記載元データの取得
  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheetByName('シート③(テスト様_google formデータ)');

  //11月のデータのみ取得
  const range = sheet.getRange("A2:C5");
  const value = range.getValues();

  //Logger.log(value)
  const list = []

  for(let i=0; i<value.length; i++){
    //11月のデータのみ取得する(該当月-1)
    if(value[i][1].getMonth()===10.0){
      const service = convertServiceToNumberSheet3(value[i][0]);
      list.push([service,value[i][1].toLocaleString('ja-JP'),value[i][2].toLocaleTimeString('ja-JP').slice(0,-3)])
    }
  }

  //Logger.log(list[0][1])

  //記載先のシート名を指定
  const outputSheet = ss.getSheetByName("シート②(テスト様_11月)");

  const dateRange = outputSheet.getRange("A2:A31");
  const dateValue = dateRange.getValues();

  //Logger.log(dateValue[9]===list[0][1])

  
  for(let v=0; v<dateValue.length;v++){
    const strDate = dateValue[v].toLocaleString('ja-JP')

    // 記入元と記入先の日時が一致した場合、そのサービス名の列に記載
    for (let w = 0; w < list.length; w++) {
      if (list[w][1] === strDate) {
        const setRange = outputSheet.getRange(v+2, list[w][0]);
        setRange.setValue(list[w][2]);
      }
    }
  } 


}

//シート③のサービス文字列⇒列番号変換関数
function convertServiceToNumberSheet3(serviceName) {
  const servicesMapping = {7:"Aサービス",8:"Bサービス",9:"Cサービス",10:"Dサービス"};

  for (let key in servicesMapping) {
    if (servicesMapping[key] == serviceName) {
      return key;
    }
  }

  // 該当するサービスが見つからなかった場合の処理
  return null;
}