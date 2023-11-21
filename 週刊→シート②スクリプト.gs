function weekScheduleOutput() {
  //スプレッドシートのidを入力
  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  const ss = SpreadsheetApp.openById(id);

  //記入元のシート名を指定
  const sheet = ss.getSheetByName('シート①(テスト様_週刊計画表)');

  //サービス名とマップリストを作成
  

  //対応列を記載
  const numbers = [2,3,4,5,6]

  const range = sheet.getRange("A2:D6");

  const value = range.getValues();

  const newValueList = [];

  //取得データの加工
  for (let i=0;i<value.length;i++){
    const startTime = value[i][2].toLocaleTimeString('ja-JP');
    const endTime = value[i][3].toLocaleTimeString('ja-JP');
    const servive = convertServiceToNumber(value[i][0]);

    //サービス
    newValueList.push([servive,value[i][1],startTime,endTime])
  }

  //記載先のシート名を指定
  const outputSheet = ss.getSheetByName("シート②(テスト様_11月)");

  const outpuRange = outputSheet.getRange("A2:F31");
  const outputValue = outpuRange.getValues();

  //Logger.log(outputValue)

  
  for(let v=0; v<outputValue.length;v++){

    //記入先の曜日データを取得
    const day = outputValue[v][0].getDay();
    const dayList = ['日','月','火','水','木','金','土'];
    const dayString = dayList[day]+'曜日';
    //Logger.log(dayString);

    // 記入先に曜日の行とサービス名A~Dの列があり、記入元の曜日が一致した場合、曜日の行・サービス名の列に●を記載
    for (let w = 0; w < newValueList.length; w++) {
      if (newValueList[w][1] === dayString) {
        const setRange = outputSheet.getRange(v+2, newValueList[w][0]);
        setRange.setValue(`${newValueList[w][2].slice(0,-3)}~${newValueList[w][3].slice(0,-3)}`);
      }
    }
  }

}

//サービス文字列⇒列番号変換関数
function convertServiceToNumber(serviceName) {
  const servicesMapping = {2:"ア_サービス",3:"イ_サービス",4:"ウ_サービス",5:"エ_サービス",6:"オ_サービス"};

  for (let key in servicesMapping) {
    if (servicesMapping[key] == serviceName) {
      return key;
    }
  }

  // 該当するサービスが見つからなかった場合の処理
  return null;
}


