//const targetPeriod = `_${nextYearMonth_().year}年${nextYearMonth_().month}月_`

function calculationTime() {
  //利用者リストのシートを開く
  const id = PropertiesService.getScriptProperties().getProperty('USER_SHEETID_LIST');
  const ss = SpreadsheetApp.openById(id);

  const data = ss.getDataRange().getValues()

  //正規表現パターンを指定(書き込み場所が異なる場合は分ける)
  //居宅AF,重度AJ,行動支援・同行支援・行動援護AB
  const regexPattern1 = new RegExp(".*"+targetPeriod+"居宅介護実績票");
  const regexPattern2 = new RegExp(".*"+targetPeriod+"重度訪問介護実績票");
  const regexPattern3 = new RegExp(".*"+targetPeriod+"(行動援護|行動支援|同行支援)"+"実績票");



  for (let i=1; i<data.length; i++){
    const sheetId = data[i][1];

    const targetSheet = SpreadsheetApp.openById(sheetId)
    const tagetSheetName = targetSheet.getSheets();

    const matchingSheets1 = [];//居宅介護
    const matchingSheets2 = [];//重度訪問
    const matchingSheets3 = [];//行動援護|行動支援|同行支援

    for(let j=0; j<tagetSheetName.length;j++){
      let sheetName = tagetSheetName[j].getName();
      if(regexPattern3.test(sheetName)){
        const kodoSheet = targetSheet.getSheetByName(sheetName)
        const kodoValues = kodoSheet.getRange("T11:T41").getValues();

        for(let m=0; m<kodoValues.length;m++){
          if(kodoValues[m][0]!==""){
            const kodoStartValuesHours = kodoValues[m][0].getHours();
            const kodoStartValuesMinutes = kodoValues[m][0].getMinutes();

            const kodoEndValuesHours = kodoSheet.getRange(`X${m+11}`).getValue().getHours();
            const kodoEndValuesMinutes = kodoSheet.getRange(`X${m+11}`).getValue().getMinutes();

            const startTime = new Date();
            startTime.setHours(kodoStartValuesHours,kodoStartValuesMinutes,0);

            const endTime = new Date();
            endTime.setHours(kodoEndValuesHours,kodoEndValuesMinutes,0);

            const timeDifference = new Date(endTime - startTime);
            const timeDifferenceHours = timeDifference.getUTCHours();
            const timeDifferenceMInutes = timeDifference.getUTCMinutes();

            const formattedtimeDifferenceMinutes = addLeadingZero_(timeDifferenceMInutes);

            kodoSheet.getRange(`AB${m+11}`).setValue(timeDifferenceHours+":"+formattedtimeDifferenceMinutes)
          }else{
            continue
          }
        }           
      }else if(regexPattern1.test(sheetName)){
        const kyotakuSheet = targetSheet.getSheetByName(sheetName)
        const kyotakuValues = kyotakuSheet.getRange("X12:X39").getValues();

        for(let m=0; m<kyotakuValues.length;m++){
          if(kyotakuValues[m][0]!==""){
            const kyotakuStartValuesHours = kyotakuValues[m][0].getHours();
            const kyotakuStartValuesMinutes = kyotakuValues[m][0].getMinutes();

            const kyotakuEndValuesHours = kyotakuSheet.getRange(`AB${m+12}`).getValue().getHours();
            const kyotakuEndValuesMinutes = kyotakuSheet.getRange(`AB${m+12}`).getValue().getMinutes();

            const startTime = new Date();
            startTime.setHours(kyotakuStartValuesHours,kyotakuStartValuesMinutes,0);

            const endTime = new Date();
            endTime.setHours(kyotakuEndValuesHours,kyotakuEndValuesMinutes,0);

            const timeDifference = new Date(endTime - startTime);
            const timeDifferenceHours = timeDifference.getUTCHours();
            const timeDifferenceMInutes = timeDifference.getUTCMinutes();

            const formattedtimeDifferenceMinutes = addLeadingZero_(timeDifferenceMInutes);

            kyotakuSheet.getRange(`AF${m+12}`).setValue(timeDifferenceHours+":"+formattedtimeDifferenceMinutes)
          }else{
            continue
          }
        }   
      }else if(regexPattern2.test(sheetName)){
        const judoSheet = targetSheet.getSheetByName(sheetName)
        const judoValues = judoSheet.getRange("AB12:AB46").getValues();

        for(let m=0; m<judoValues.length;m++){
          if(judoValues[m][0]!==""){
            const judoStartValuesHours = judoValues[m][0].getHours();
            const judoStartValuesMinutes = judoValues[m][0].getMinutes();

            const judoEndValuesHours = judoSheet.getRange(`AF${m+12}`).getValue().getHours();
            const judoEndValuesMinutes = judoSheet.getRange(`AF${m+12}`).getValue().getMinutes();

            //matchingSheets2.push(judoEndValuesHours)

            const startTime = new Date();
            startTime.setHours(judoStartValuesHours,judoStartValuesMinutes,0);

            const endTime = new Date();
            endTime.setHours(judoEndValuesHours,judoEndValuesMinutes,0);

            const timeDifference = new Date(endTime - startTime);
            const timeDifferenceHours = timeDifference.getUTCHours();
            const timeDifferenceMInutes = timeDifference.getUTCMinutes();

            const formattedtimeDifferenceMinutes = addLeadingZero_(timeDifferenceMInutes);

            judoSheet.getRange(`AJ${m+12}`).setValue(timeDifferenceHours+":"+formattedtimeDifferenceMinutes)
          }else{
            continue
          }
        }        
      }
    }    
  }
  //シート内の各シートを参照
  //開始時間終了時間の記載があれば終了-開始時間で計算、キャンセル枠に記載があるもしくは空白の場合は算定しない
}

function addLeadingZero_(number) {
  return (number < 10 ? '0' : '') + number;
}
