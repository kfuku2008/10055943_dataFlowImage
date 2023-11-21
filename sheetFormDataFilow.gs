//★ご依頼内容①

//週刊計画表を月次スケージュールへ移行
function weekScheduleOutput() {
  //スプレッドシートのidを入力
  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  const ss = SpreadsheetApp.openById(id);

  //記入元のシート名を指定
  const sheet = ss.getSheetByName('シート①(テスト様_週刊計画表)');  

  //対応列を記載
  const numbers = [2,3,4,5,6];

  const range = sheet.getRange("A2:D6");

  const value = range.getValues();

  const newValueList = [];

  //取得データの加工
  for (let i=0;i<value.length;i++){
    const startTime = value[i][2].toLocaleTimeString('ja-JP');
    const endTime = value[i][3].toLocaleTimeString('ja-JP');
    const servive = convertServiceToNumber(value[i][0]);

    newValueList.push([servive,value[i][1],startTime,endTime]);
  }

  //記載先のシート名を指定
  const outputSheet = ss.getSheetByName("シート②(テスト様_11月)");

  const outpuRange = outputSheet.getRange("A2:F31");
  const outputValue = outpuRange.getValues();
  
  for(let v=0; v<outputValue.length;v++){

    //記入先の曜日データを取得
    const day = outputValue[v][0].getDay();
    const dayList = ['日','月','火','水','木','金','土'];
    const dayString = dayList[day]+'曜日';

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
  //サービス名と記載先シート②の列番号を対応させる
  const servicesMapping = {2:"ア_サービス",3:"イ_サービス",4:"ウ_サービス",5:"エ_サービス",6:"オ_サービス"};

  for (let key in servicesMapping) {
    if (servicesMapping[key] == serviceName) {
      return key;
    }
  }
  // 該当するサービスが見つからなかった場合の処理
  return null;
}

//★ご依頼内容②

//フォームの回答データを取得し、スプレッドシート③へ格納
//フォームのIDをスクリプトプロパティに格納
function getFormData() {
  const formId = PropertiesService.getScriptProperties().getProperty("FORM_ID");
  const form = FormApp.openById(formId);
  const responses = form.getResponses();
  const itemResponses = responses[0].getItemResponses();

  const answerList =[];

  for (const itemResponse of itemResponses){
    const answer = itemResponse.getResponse();

    answerList.push(answer);
  }

  answerList.shift()
  
  //スプレッドシートに記載
  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheetByName('シート③(テスト様_google formデータ)');

  //最終行を取得
  const lastRow = sheet.getLastRow();

  //最終行＋1に記載
  const setRange = sheet.getRange(lastRow+1,1,1,3);
  setRange.setValues([answerList]);
}


//★ご依頼内容③

//スプレッドシート③をシート②へ記載
function formDataOutput() {
  //記載元データの取得
  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheetByName('シート③(テスト様_google formデータ)');

  //11月のデータのみ取得
  const range = sheet.getRange("A2:C5");
  const value = range.getValues();

  const list = []

  for(let i=0; i<value.length; i++){
    //11月のデータのみ取得する(該当月-1)
    if(value[i][1].getMonth()===10.0){
      const service = convertServiceToNumberSheet3(value[i][0]);
      list.push([service,value[i][1].toLocaleString('ja-JP'),value[i][2].toLocaleTimeString('ja-JP').slice(0,-3)]);
    }
  }

  //記載先のシート名を指定
  const outputSheet = ss.getSheetByName("シート②(テスト様_11月)");

  const dateRange = outputSheet.getRange("A2:A31");
  const dateValue = dateRange.getValues();
 
  for(let v=0; v<dateValue.length;v++){
    const strDate = dateValue[v].toLocaleString('ja-JP');

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
  //サービス名と記載先シート②の列番号を対応させる
  const servicesMapping = {7:"Aサービス",8:"Bサービス",9:"Cサービス",10:"Dサービス"};

  for (let key in servicesMapping) {
    if (servicesMapping[key] == serviceName) {
      return key;
    }
  }

  // 該当するサービスが見つからなかった場合の処理
  return null;
}
