//フォームのIDをスクリプトプロパティに格納
function getFormData() {
  const formId = PropertiesService.getScriptProperties().getProperty("FORM_ID");
  const form = FormApp.openById(formId);
  const responses = form.getResponses();
  const itemResponses = responses[0].getItemResponses();

  const answerList =[]

  for (const itemResponse of itemResponses){
    const answer = itemResponse.getResponse();

    answerList.push(answer)
  }

  answerList.shift()
  
  //スプレッドシートに記載
  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheetByName('シート③(テスト様_google formデータ)');

  //最終行を取得
  const lastRow = sheet.getLastRow();

  //最終行＋1に記載
  const setRange = sheet.getRange(lastRow+1,1,1,3)
  setRange.setValues([answerList])
}
