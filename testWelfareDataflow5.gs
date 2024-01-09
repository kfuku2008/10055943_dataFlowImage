function dbCleaningTest() {
  //対象のスプレッドシートを開く
  const id = PropertiesService.getScriptProperties().getProperty('ICHIRANHYOU_KYOTAKUJUDO');
  const ss = SpreadsheetApp.openById(id);
  const sheet = ss.getSheetByName('クリーニング後_テストサンプル');
  const value = sheet.getDataRange().getValues()
  const valueList = []

  for(let k=1; k<value.length;k++){
    const rowNum = k+1;
    valueList.push([rowNum,value[k][1],value[k][4],value[k][6]])
  }

  for (let i = 0; i < valueList.length; i++) {
    // const currentArray = valueList[i];
    // let duplicateFound = false;
    let beforeRowNum = sheet.getLastRow()

    for(let j=i+1;j<valueList.length;j++){
      if(valueList[i][1]===valueList[j][1] && valueList[i][2].toLocaleString('ja-JP')===valueList[j][2].toLocaleString('ja-JP') && valueList[i][3].toLocaleString('ja-JP')===valueList[j][3].toLocaleString('ja-JP')){
        sheet.deleteRow(valueList[i][0])
        break;
      }
    }
    let afterRowNum = sheet.getLastRow()

    //二次元配列格納の行数を更新する
    for(let m=i+1;m<valueList.length;m++){
      if(beforeRowNum !== afterRowNum){
        valueList[m][0] -=1
      }
    }
  }
}