//========================================
//グローバル変数(簡略化する)
//シート名に付与する翌月の値(グローバル変数として使用)
const targetName = `${nextYearMonth_().year}年${nextYearMonth_().month}月`;

//実績票シート呼び出し変数
const jissekiGenponSheetNameList = ["居宅介護実績票","重度訪問介護実績票","行動援護実績票","同行支援実績票","行動支援実績票"];

//計画表シート呼び出し変数
const keikakushoList = ['計画表_居宅介護','計画表_重度訪問介護','計画表_行動援護','計画表_同行支援','計画表_行動支援'];

//----------------------------------------
//スクリプトプロパティ（★スクリプトプロパティに値格納されている事を確認）

//利用者シートIDリストのスプシID
const userIdList = PropertiesService.getScriptProperties().getProperty('USER_SHEETID_LIST');

//一覧表のスプレッドシートID（シート内に「居宅介護・重度訪問介護」「行動援護・同行支援・行動支援」の2シート格納）
const ichiranhyo = PropertiesService.getScriptProperties().getProperty('ICHIRANHYO');

//GoogleフォームのID（居宅介護・重度訪問介護）
const formKyotakuJudoId = PropertiesService.getScriptProperties().getProperty('FORM_KYOTAKUJUDO_ID');

//GoogleフォームのID（行動援護・同行支援・行動支援）
const formKodoId = PropertiesService.getScriptProperties().getProperty('FORM_KODO_ID');


//========================================
//機能１（翌月分の実績シート作成&計画表記載項目書き込み）

//★要トリガー設定：翌月の実績記録票のコピーを行う関数
function copyForm(){

  //必要に応じてランタイムエラー対策コード追加

  for(let i=0; i<inputUserData_().length; i++){

    //スプシIDを指定しシートを開く
    const id = inputUserData_()[i][1];
    const ss = SpreadsheetApp.openById(id);

    //ここで5シート分新たに作成する
    for(let j=0;j<jissekiGenponSheetNameList.length;j++){
      //実績票原本のスプシを開く
      const genponSheet = ss.getSheetByName(jissekiGenponSheetNameList[j]);
      const sourceRange = genponSheet.getDataRange();

      const newSheet = ss.insertSheet(`${inputUserData_()[i][0]}様_${targetName}_${jissekiGenponSheetNameList[j]}`);
    
      //データの書き込み(仕様上左から2番目に書き込まれてしまう(1シート目を目次にするなどが良い？))
      sourceRange.copyTo(newSheet.getRange(1,1));      
    }
  }
}

//★要トリガー設定：シートの体裁調整(列幅を調整)
function formatSheet(){
  //複製したフォーマットを開く
  for(let i=0; i<inputUserData_().length; i++){
    const id = inputUserData_()[i][1];
    const ss = SpreadsheetApp.openById(id);

    for(let j=0;j<jissekiGenponSheetNameList.length;j++){
      //実績票原本のスプシを開く
      const jissekiSheet = ss.getSheetByName(`${inputUserData_()[i][0]}様_${targetName}_${jissekiGenponSheetNameList[j]}`);

      const dupSheetRange = jissekiSheet.getDataRange();

      for (let k=1;k<=dupSheetRange.getNumColumns();k++){
          jissekiSheet.setColumnWidth(k,15)//スプシの列幅を調整
      } 
   }
  }
}

//★要トリガー設定：実績記録票への書き出しを行う関数
function outputData(){
  
  //createNextMonthList関数で翌月の[日付,曜日]リストを取得
  const nextMonthDayList = createNextMonthList_()

  //ユーザー数分のシートを開く
  for(let i=0;i<inputUserData_().length; i++){

    const outputList = []//[日付,曜日,サービス名,開始,終了]

    //計画表のデータを取得(サービス名、曜日、開始、終了時間)
    const inputId =inputUserData_()[i][1];
    const inputData = weeklyDataInput_(inputId)//計画表のデータを取り出す

    //Logger.log(inputData)



    //inputData.length=5(計画表のデータ数)
    for(let k=0; k<inputData.length; k++){
      //データが存在する場合のみ処理を実行
      //中間リストを作成
      const mediumList = []

      if(inputData[k].length !==0){               
        for (let j=0; j<nextMonthDayList.length; j++){   
          for(let x=0;x<inputData[k].length;x++){
            //シートに「該当なし」の記載がなければ以下を実行
            //曜日が一致した場合リスト(outputList)にデータを書き込む
            if(nextMonthDayList[j][1]===inputData[k][x][1]){
              mediumList.push([nextMonthDayList[j][0],nextMonthDayList[j][1],inputData[k][x][0],inputData[k][x][2],inputData[k][x][3]])
            }
          }                  
        }
        outputList.push(mediumList)  
      }else{
        outputList.push([])
      }
    }
    // ここで実績票に書き込み  
    // 自動作成したスプレッドシートへ記載する
    // 別途作成したシートに書き出す
    const ss = SpreadsheetApp.openById(inputId);

    //シート内の実績票を開く
    for(let w=0;w<jissekiGenponSheetNameList.length;w++){      
      //outputList(計画表から取り出したデータ：データ数5(5シート分のデータを意味する))
      for(let y=0;y<outputList.length;y++){
          //シート内のシート名を指定する
          let outputSheet = ss.getSheetByName(`${inputUserData_()[i][0]}様_${targetName}_${jissekiGenponSheetNameList[y]}`);
        if(outputList[y].length !==0){
          //y=2,3,4は分岐
          if(y===2||y===3||y===4){
            for(let z=0; z<outputList[y].length;z++){
              //12行目から記載する[C:日付,E:曜日,H:サービス内容,L:開始時間,O:終了時間]
              outputSheet.getRange(`B${11+z}`).setValue(outputList[y][z][0]);
              outputSheet.getRange(`D${11+z}`).setValue(outputList[y][z][1]);
              outputSheet.getRange(`H${11+z}`).setValue(outputList[y][z][3]);
              outputSheet.getRange(`L${11+z}`).setValue(outputList[y][z][4]);
            }
          }else if(y===1){
            for(let z=0; z<outputList[y].length;z++){
              //12行目から記載する[C:日付,E:曜日,H:サービス内容,L:開始時間,O:終了時間]
              outputSheet.getRange(`D${12+z}`).setValue(outputList[y][z][0]);
              outputSheet.getRange(`F${12+z}`).setValue(outputList[y][z][1]);
              outputSheet.getRange(`I${12+z}`).setValue(outputList[y][z][2]);
              outputSheet.getRange(`N${12+z}`).setValue(outputList[y][z][3]);
              outputSheet.getRange(`R${12+z}`).setValue(outputList[y][z][4]);
            }
           }else if(y===0){
            for(let z=0; z<outputList[y].length;z++){
              //12行目から記載する[C:日付,E:曜日,H:サービス内容,L:開始時間,O:終了時間]
              outputSheet.getRange(`C${12+z}`).setValue(outputList[y][z][0]);
              outputSheet.getRange(`E${12+z}`).setValue(outputList[y][z][1]);
              outputSheet.getRange(`H${12+z}`).setValue(outputList[y][z][2]);
              outputSheet.getRange(`L${12+z}`).setValue(outputList[y][z][3]);
              outputSheet.getRange(`O${12+z}`).setValue(outputList[y][z][4]);
            }
          }
        }
      }    
    }
  }
}

//========================================
//機能２（フォーム入力データ→一覧表取り込み）

//★要トリガー設定：フォーム→一覧表（居宅介護・重度訪問介護）
function answertoDataBaseKyotakuJudo(){
  const dataList = formToDatabaseKyotakuJudo_();

  const ss = SpreadsheetApp.openById(ichiranhyo);

  //一覧表のシート名をここで指定
  const sheet = ss.getSheetByName('居宅介護・重度訪問介護');

  //最終行のデータを取得
  const lastRow = sheet.getLastRow()

  for(let i=0; i<dataList.length;i++){

    //複数選択時はリスト形式で格納されているので「A,B,C,,」形式で一つのセルに格納できるようにする
    const formatdataList = flattenArray_(dataList[i])

    //列数確認用
    //Logger.log(formatdataList.length)

    // //列方向に書き出す場合はfor文で書く
    for (let j=0; j<formatdataList.length;j++){
      sheet.getRange(lastRow+1+i,j+1).setValue(formatdataList[j])
    }
  }
}

//★要トリガー設定：フォーム→一覧表（行動援護・同行支援・行動支援）
function answertoDataBaseDokoKodoIdo(){
  const dataList = formToDatabaseKodo_();
  
  const ss = SpreadsheetApp.openById(ichiranhyo);

  //一覧表のシート名をここで指定
  const sheet = ss.getSheetByName('行動援護・同行支援・行動支援');

  //最終行のデータを取得
  const lastRow = sheet.getLastRow()

  for(let i=0; i<dataList.length;i++){

    //複数選択時はリスト形式で格納されているので「A,B,C,,」形式で一つのセルに格納できるようにする
    const formatdataList = flattenArray_(dataList[i])

    //列数確認用
    //Logger.log(formatdataList.length)

    // //列方向に書き出す場合はfor文で書く
    for (let j=0; j<formatdataList.length;j++){
      sheet.getRange(lastRow+1+i,j+1).setValue(formatdataList[j])
    }
  }
}
//========================================
//機能２．５(DBの重複時古いデータの削除実行)

//★要トリガー設定：居宅介護・重度訪問介護
function dbCleaningKyotakuJudo() {
  //対象のスプレッドシートを開く
  const ss = SpreadsheetApp.openById(ichiranhyo);
  const sheet = ss.getSheetByName('居宅介護・重度訪問介護');
  const value = sheet.getDataRange().getValues()

  const valueList = []

  for(let k=1; k<value.length;k++){
    const rowNum = k+1;
    valueList.push([rowNum,value[k][1],value[k][4],value[k][6]])
  }

  for (let i = 0; i < valueList.length; i++) {

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

//★要トリガー設定：行動援護・同行支援・行動支援
function dbCleaningKodo() {
  //対象のスプレッドシートを開く
  const ss = SpreadsheetApp.openById(ichiranhyo);
  const sheet = ss.getSheetByName('行動援護・同行支援・行動支援');
  const value = sheet.getDataRange().getValues()
  const valueList = []

  for(let k=1; k<value.length;k++){
    const rowNum = k+1;
    valueList.push([rowNum,value[k][1],value[k][4],value[k][6]])
  }

  for (let i = 0; i < valueList.length; i++) {

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


//========================================
//機能３（一覧表→実績票反映）

//※本番は翌月実行の前月分書き込みとなるため、指定方法を変更する必要あり

//書き込むシートを特定・サービス区分を紐づける

//★要トリガー設定：居宅介護
function decisionSheetKyotaku(){

  if(connectionJissekiSheetKyotaku_().length !==0){
    //ターゲットとなるスプシを指定
    for (let j=0; j<connectionJissekiSheetKyotaku_().length; j++){

      //スプシを開く
      const ss = SpreadsheetApp.openById(connectionJissekiSheetKyotaku_()[j]);

      //選択したスプシのシート名を取得して実績票（翌月＋該当サービス区分）を取り出す
      const sheetList = ss.getSheets();

      //DBから取得したサービス区分を参照
      for (let k=0;k<inputDbDataKyotaku_().length;k++){
    
        //DBのサービス区分に該当するシートを開く
        for (let i=0; i<sheetList.length;i++){
          const sheetName = sheetList[i].getName();
          if(sheetName.includes(`${nextYearMonth_().year}年${nextYearMonth_().month}月_居宅介護実績票`)){
            //シート名、サービス日、計画開始時間、サービス開始・終了時間
            const sheet = ss.getSheetByName(sheetName);

            //dbのサービス日と計画開始時間を取得
            const dbServiceDate = inputDbDataKyotaku_()[k][4].getDate()
            const dbPlanStartTime = `${inputDbDataKyotaku_()[k][6].getHours()},${inputDbDataKyotaku_()[k][6].getMinutes()}`
            const dbCancelMessage = inputDbDataKyotaku_()[k][46];

            const sheetTimes = sheet.getRange("L12:L46").getValues();
            const sheetDate = sheet.getRange("C12:C46").getValues();

            for(let l=0;l<sheetTimes.length;l++){
              if(sheetTimes[l][0]!==""){
                const sheetTime = `${sheetTimes[l][0].getHours()},${sheetTimes[l][0].getMinutes()}`
                if(dbPlanStartTime===sheetTime && dbServiceDate===sheetDate[l][0] && dbCancelMessage === ""){
                  sheet.getRange(`X${l+12}`).setValue(inputDbDataKyotaku_()[k][5]);
                  sheet.getRange(`AB${l+12}`).setValue(inputDbDataKyotaku_()[k][7]);
                }else if(dbPlanStartTime===sheetTime && dbServiceDate===sheetDate[l][0] && dbCancelMessage !== ""){
                  sheet.getRange(`AZ${l+12}`).setValue(inputDbDataKyotaku_()[k][46]);
                }
              }
            }          
          }
        }
      }
    }
  }
}

//★要トリガー設定：重度訪問介護
//書き込むシートを特定・サービス区分を紐づける
function decisionSheetJudo(){

  if(connectionJissekiSheetJudo_().length !==0){
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
            const sheet = ss.getSheetByName(sheetName);

            //dbのサービス日と計画開始時間を取得
            const dbServiceDate = inputDbDataJudo_()[k][4].getDate()
            const dbPlanStartTime = `${inputDbDataJudo_()[k][6].getHours()},${inputDbDataJudo_()[k][6].getMinutes()}`
            const dbCancelMessage = inputDbDataJudo_()[k][46];

            const sheetTimes = sheet.getRange("N12:N46").getValues();
            const sheetDate = sheet.getRange("D12:D46").getValues();

            for(let l=0;l<sheetTimes.length;l++){
              if(sheetTimes[l][0]!==""){
                const sheetTime = `${sheetTimes[l][0].getHours()},${sheetTimes[l][0].getMinutes()}`
                if(dbPlanStartTime===sheetTime && dbServiceDate===sheetDate[l][0] && dbCancelMessage === ""){
                  sheet.getRange(`AB${l+12}`).setValue(inputDbDataJudo_()[k][5]);
                  sheet.getRange(`AF${l+12}`).setValue(inputDbDataJudo_()[k][7]);
                }else if(dbPlanStartTime===sheetTime && dbServiceDate===sheetDate[l][0] && dbCancelMessage !== ""){
                  sheet.getRange(`BJ${l+12}`).setValue(inputDbDataJudo_()[k][46]);
                }
              }
            }          
          }
        }
      }
    }
  }
}

//★要トリガー設定：行動援護
function decisionSheetKodoengo(){
  //ターゲットとなるスプシを指定
  if(connectionJissekiSheetKodoengo_().length !==0){
  for (let j=0; j<connectionJissekiSheetKodoengo_().length; j++){

    //スプシを開く
    const ss = SpreadsheetApp.openById(connectionJissekiSheetKodoengo_()[j]);

    //選択したスプシのシート名を取得して実績票（翌月＋該当サービス区分）を取り出す
    const sheetList = ss.getSheets();

        //DBから取得したサービス区分を参照
        for (let k=0;k<inputDbDataKodoengo_().length;k++){
      
          for (let i=0; i<sheetList.length;i++){
            const sheetName = sheetList[i].getName();
            if(sheetName.includes(`${nextYearMonth_().year}年${nextYearMonth_().month}月_行動援護実績票`)){
              //シート名、サービス日、計画開始時間、サービス開始・終了時間
              const sheet = ss.getSheetByName(sheetName);

              //dbのサービス日と計画開始時間を取得
              const dbServiceDate = inputDbDataKodoengo_()[k][4].getDate()
              const dbPlanStartTime = `${inputDbDataKodoengo_()[k][6].getHours()},${inputDbDataKodoengo_()[k][6].getMinutes()}`
              const dbCancelMessage = inputDbDataKodoengo_()[k][29];

              const sheetTimes = sheet.getRange("H11:H41").getValues();
              const sheetDate = sheet.getRange("B11:B41").getValues();

              for(let l=0;l<sheetTimes.length;l++){
                if(sheetTimes[l][0]!==""){
                  const sheetTime = `${sheetTimes[l][0].getHours()},${sheetTimes[l][0].getMinutes()}`
                  if(dbPlanStartTime===sheetTime && dbServiceDate===sheetDate[l][0] && dbCancelMessage === ""){
                    sheet.getRange(`T${l+11}`).setValue(inputDbDataKodoengo_()[k][5]);
                    sheet.getRange(`X${l+11}`).setValue(inputDbDataKodoengo_()[k][7]);
                  }else if(dbPlanStartTime===sheetTime && dbServiceDate===sheetDate[l][0] && dbCancelMessage !== ""){
                    sheet.getRange(`AU${l+11}`).setValue(inputDbDataKodoengo_()[k][29]);
                  }
                }
              }          
            }
          }
        }
    }
  }
}

//★要トリガー設定：同行支援
function decisionSheetDokoShien(){

  if(connectionJissekiSheetDokoShien_().length !==0){

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
            const dbCancelMessage = inputDbDataDokoShien_()[k][29];

            const sheetTimes = sheet.getRange("H11:H41").getValues();
            const sheetDate = sheet.getRange("B11:B41").getValues();

            for(let l=0;l<sheetTimes.length;l++){
              if(sheetTimes[l][0]!==""){
                const sheetTime = `${sheetTimes[l][0].getHours()},${sheetTimes[l][0].getMinutes()}`
                if(dbPlanStartTime===sheetTime && dbServiceDate===sheetDate[l][0] && dbCancelMessage === ""){
                  sheet.getRange(`T${l+11}`).setValue(inputDbDataDokoShien_()[k][5]);
                  sheet.getRange(`X${l+11}`).setValue(inputDbDataDokoShien_()[k][7]);
                }else if(dbPlanStartTime===sheetTime && dbServiceDate===sheetDate[l][0] && dbCancelMessage !== ""){
                  sheet.getRange(`AU${l+11}`).setValue(inputDbDataDokoShien_()[k][29]);
                }
              }
            }          
          }
        }
      }
    }
  }
}

//★要トリガー設定：行動支援
function decisionSheetKodoshien(){

  if(connectionJissekiSheetKodoshien_().length !==0){
    //ターゲットとなるスプシを指定
    for (let j=0; j<connectionJissekiSheetKodoshien_().length; j++){

      //スプシを開く
      const ss = SpreadsheetApp.openById(connectionJissekiSheetKodoshien_()[j]);

      //選択したスプシのシート名を取得して実績票（翌月＋該当サービス区分）を取り出す
      const sheetList = ss.getSheets();

      //DBから取得したサービス区分を参照
      for (let k=0;k<inputDbDataKodoshien_().length;k++){
    
        for (let i=0; i<sheetList.length;i++){
          const sheetName = sheetList[i].getName();
          if(sheetName.includes(`${nextYearMonth_().year}年${nextYearMonth_().month}月_行動支援実績票`)){
            //シート名、サービス日、計画開始時間、サービス開始・終了時間
            const sheet = ss.getSheetByName(sheetName);

            //dbのサービス日と計画開始時間を取得
            const dbServiceDate = inputDbDataKodoshien_()[k][4].getDate()
            const dbPlanStartTime = `${inputDbDataKodoshien_()[k][6].getHours()},${inputDbDataKodoshien_()[k][6].getMinutes()}`
            const dbCancelMessage = inputDbDataKodoshien_()[k][29];

            const sheetTimes = sheet.getRange("H11:H41").getValues();
            const sheetDate = sheet.getRange("B11:B41").getValues();

            for(let l=0;l<sheetTimes.length;l++){
              if(sheetTimes[l][0]!==""){
                const sheetTime = `${sheetTimes[l][0].getHours()},${sheetTimes[l][0].getMinutes()}`
                if(dbPlanStartTime===sheetTime && dbServiceDate===sheetDate[l][0] && dbCancelMessage === ""){
                  sheet.getRange(`T${l+11}`).setValue(inputDbDataKodoshien_()[k][5]);
                  sheet.getRange(`X${l+11}`).setValue(inputDbDataKodoshien_()[k][7]);
                }else if(dbPlanStartTime===sheetTime && dbServiceDate===sheetDate[l][0] && dbCancelMessage !== ""){
                  sheet.getRange(`AU${l+11}`).setValue(inputDbDataKodoshien_()[k][29]);
                }
              }
            }          
          }
        }
      }
    }
  }
}

//★要トリガー設定：実績票算定時間の算出
function calculationTime() {
  //利用者リストのシートを開く
  const id = PropertiesService.getScriptProperties().getProperty('USER_SHEETID_LIST');
  const ss = SpreadsheetApp.openById(id);

  const data = ss.getDataRange().getValues()

  //正規表現パターンを指定(書き込み場所が異なる場合は分ける)
  //居宅AF,重度AJ,行動支援・同行支援・行動援護AB
  const regexPattern1 = new RegExp(".*"+targetName+"_居宅介護実績票");
  const regexPattern2 = new RegExp(".*"+targetName+"_重度訪問介護実績票");
  const regexPattern3 = new RegExp(".*"+targetName+"_(行動援護|行動支援|同行支援)"+"実績票");

  for (let i=1; i<data.length; i++){
    const sheetId = data[i][1];

    const targetSheet = SpreadsheetApp.openById(sheetId)
    const tagetSheetName = targetSheet.getSheets();

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


//===============================================
//以下はトリガー不要な関数

//--------------------------------
//機能１
//居宅介護計画書の週間データを取得する関数
function weeklyDataInput_(id) {
  const ss = SpreadsheetApp.openById(id);

  //サービス名とセルの範囲の二次元配列を作成
  const nonEnptyCells = [];

  const targetList = []

  for(let t=0;t<keikakushoList.length;t++){

    const sheet = ss.getSheetByName(keikakushoList[t]);

    const ranges = sheet.getRange("D19:J43");

    const cells = ranges.getValues();

    //Logger.log(cells)

    //セル結合されているデータの先頭セルを取得
    const mergeCellsList = [];
    const mergeRanges = ranges.getMergedRanges();
    for (const mergeRange of mergeRanges){
      mergeCellsList.push(mergeRange.getA1Notation())
    }
    //Logger.log(mergeCellsList)

    //位置情報付きのサービスデータを取得
    //A1Notationはセルのアドレス取得に使用
    const positions = ranges.getA1Notation().split(":");
    const startCell = positions[0];
    const startRow = parseInt(startCell.match(/\d+/)[0]);
    const startCol = startCell.replace(/\d/,'');

    let mediumCellsList = []
    for (let i =0; i<cells.length;i++){
      
      for (let j=0; j<cells[i].length;j++){
        //セルが空でない値を取得
        if (cells[i][j] !== ""){
          let row = startRow + i;
          let col = String.fromCharCode(startCol.charCodeAt(0)+j);
          let position = col + row;

          for (let k=0;k<mergeCellsList.length;k++){
            if(position === mergeCellsList[k].slice(0,3)){
              position = mergeCellsList[k];
            }
          }
            mediumCellsList.push([cells[i][j],position])
        }         
      }
    }
    nonEnptyCells.push(mediumCellsList);
  }

    for (let l=0;l<nonEnptyCells.length;l++){
      if(nonEnptyCells[l].length !==0){
        let mediumCellsList2 = []
        for(let p=0;p<nonEnptyCells[l].length;p++){

          let dateCol = nonEnptyCells[l][p][1].slice(0,1);
          let startTimeRow = nonEnptyCells[l][p][1].slice(1,3);
          let endTimeRow = nonEnptyCells[l][p][1].slice(1,3);

          if(nonEnptyCells[l][p][1].includes(":")){
            dateCol = nonEnptyCells[l][p][1].slice(0,1);
            startTimeRow = nonEnptyCells[l][p][1].slice(1,3);
            endTimeRow = nonEnptyCells[l][p][1].slice(5);
          }
          
          //Logger.log(dateCol)
          let date = dateTransform_(dateCol);
          let startTime = startTimeRow-19 + ":00"
          let endTime = endTimeRow-18 + ":00"

          mediumCellsList2.push([nonEnptyCells[l][p][0],date,startTime,endTime])
        } targetList.push(mediumCellsList2);
      }else{
        targetList.push([])
      }
    }
  
   return targetList;
  //出力イメージ
  //[[], [[通院介助, 火, 5:00, 6:00], [身体介護, 火, 14:00, 17:00]], [], [[同行支援, 火, 8:00, 9:00]], [[行動支援, 水, 14:00, 15:00]]]
}

//居宅介護計画書の列データを曜日変換する関数
function dateTransform_(targetDate){
  const dateMapping = {"月":"D","火":"E","水":"F","木":"G","金":"H","土":"I","日":"J"};

  for (let key in dateMapping){
    if(dateMapping[key]== targetDate){
      return key
    }
  }
  return null;
}

//実績記録票出力時の翌月の日付と曜日リストを返す関数
function createNextMonthList_(){
  //取得したデータを翌月の該当曜日に書き込む
  //翌月のデータを取得
  // 現在の日付を取得
  const today = new Date();

  // 1ヶ月後の日付を計算
  const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1);


  //翌月の曜日リストを取得
  const dayOfWeek = ['日','月','火','水','木','金','土'];

  const firstDayOfWeek = nextMonth.getDay();

  //翌月の日数を取得(日にち部分を0にすると先月の最終日が取れる)
  const daysInNextMonth = new Date(nextMonth.getFullYear(),nextMonth.getMonth()+1,0).getDate();

  //日付と曜日リストを作成
  const list = [];
  
  for (let i=0;i<daysInNextMonth;i++){
    const day = (i+firstDayOfWeek) % 7;//曜日のインデックス
    list.push([i+1,dayOfWeek[day]])
  }

  return list;
}

//翌月の年月を返す関数
function nextYearMonth_(){

  //月のデータは値のセット時-1が必要
  const currentDate =  new Date();

  //翌月の年月を計算
  const nextMonth = new Date(currentDate);
  nextMonth.setMonth(currentDate.getMonth()+1)

  //Logger.log(nextMonth)

  const nextYear = nextMonth.getFullYear();

  //現実の月データが欲しい場合は+1
  const nextMonthNumber = nextMonth.getMonth()+1 ;

  return{
    year: nextYear,
    month: nextMonthNumber
  };
}

//ご利用者様のお名前とスプレッドシートidを取り出す
function inputUserData_(){
  //リストとなるスプレッドシートを開く
  const listId = PropertiesService.getScriptProperties().getProperty("USER_SHEETID_LIST")
  const ss = SpreadsheetApp.openById(listId);
  const sheet = ss.getSheetByName('利用者リスト');

  //最終行を取得
  const lastRow = sheet.getLastRow();
  const range = sheet.getRange(2,1,lastRow-1,2);
  const value = range.getValues();

  return value;
}
//--------------------------------
//機能２
//居宅介護・重度訪問
function formToDatabaseKyotakuJudo_(){
  //回答データの取得
  const form = FormApp.openById(formKyotakuJudoId);

  //formの回答者数をリストに格納
  const responses = form.getResponses();
  const today = new Date()
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate()-1);

  const formattedYesterday = yesterday.toDateString();


  //複数人回答用のリストを準備
  const answersList=[];

  //実行日に回答があったデータを取得
  for (let i =0;i<responses.length;i++){

    //1つ目(1人目)の回答データを取得
    let response = responses[i];

    //回答時間のタイムスタンプの取得
    let timeStamp = response.getTimestamp().toDateString();

    let answerList = [];


    if(timeStamp===formattedYesterday){

      answerList = [timeStamp];

      //各アイテムの質問と回答を取得
      let itemResponses = response.getItemResponses();

      for(const itemResponse of itemResponses){

        //アイテムの回答を取得
        const answer = itemResponse.getResponse();

        answerList.push(answer)
      }
    }else{
      continue;
    }
    answersList.push(answerList)
  }

  //Logger.log(answersList)

  return answersList
}

//行動援護・同行支援・行動支援
function formToDatabaseKodo_(){
  //回答データの取得
  const form = FormApp.openById(formKodoId);

  //formの回答者数をリストに格納
  const responses = form.getResponses();
  const today = new Date()
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate()-1);

  const formattedYesterday = yesterday.toDateString();

  //複数人回答用のリストを準備
  const answersList=[];

  //実行日に回答があったデータを取得
  for (let i =0;i<responses.length;i++){

    //1つ目(1人目)の回答データを取得
    let response = responses[i];

    //回答時間のタイムスタンプの取得
    let timeStamp = response.getTimestamp().toDateString();

    let answerList = [];

    if(timeStamp===formattedYesterday){

      answerList = [timeStamp];

      //各アイテムの質問と回答を取得
      let itemResponses = response.getItemResponses();

      for(const itemResponse of itemResponses){

        //アイテムの回答を取得
        const answer = itemResponse.getResponse();

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


//複数入力項目変換関数(/で区切っている)
function flattenArray_(originalArray) {
  const flattenedArray = [];
  
  for (let i = 0; i < originalArray.length; i++) {
    const item = originalArray[i];
    
    if (Array.isArray(item)) {
      flattenedArray.push(item.join('/'));
    } else {
      flattenedArray.push(item);
    }
  }  
  return flattenedArray;
}

//--------------------------------
//機能３

//居宅介護
function inputDbDataKyotaku_() {
  //一覧表シートを開く(シート名はデータベースサンプル)
  const sheet = SpreadsheetApp.openById(ichiranhyo).getSheetByName('居宅介護・重度訪問介護');
  const value = sheet.getDataRange().getValues();
  
  //サービス日時を参照して翌月のデータのみを取得する
  //サービス日時は4列目

  //翌月の年月データを取得
  const nextMonth = `${nextYearMonth_().year}/${nextYearMonth_().month}`

  const valueList =[]

  //0はヘッダーのため1から
  for(let i=1;i<value.length;i++){
    if(value[i][4].toLocaleString('ja-JP').slice(0,6)===nextMonth && value[i][8]==='居宅介護'){
      valueList.push(value[i])
    }
  }

  return valueList
}

function connectionJissekiSheetKyotaku_(){

  //利用者リストを開く(シート名は利用者リスト)
  const userListSheet = SpreadsheetApp.openById(userIdList).getSheetByName('利用者リスト');

  const lastRow = userListSheet.getLastRow()

  const userIdColumnValues = userListSheet.getRange(`C2:C${lastRow}`).getValues();

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
    for(let k=0;k<inputDbDataKyotaku_().length;k++){
      if(inputDbDataKyotaku_()[k][1]===userIdColumnValuesList[j][0]){
        const targetRow = userIdColumnValuesList[j][1] ;
        const targetId = userListSheet.getRange(`B${targetRow}`).getValue();
        targetIdList.push(targetId)
      }
    }  
  }

  const targetUniqueIdList = Array.from(new Set(targetIdList))

  return targetUniqueIdList
}

//重度訪問介護

function inputDbDataJudo_() {
  //一覧表シートを開く(シート名はデータベースサンプル)
  const sheet = SpreadsheetApp.openById(ichiranhyo).getSheetByName('居宅介護・重度訪問介護');
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

function connectionJissekiSheetJudo_(){

  //利用者リストを開く(シート名は利用者リスト)
  const userListSheet = SpreadsheetApp.openById(userIdList).getSheetByName('利用者リスト');

  const lastRow = userListSheet.getLastRow()

  const userIdColumnValues = userListSheet.getRange(`C2:C${lastRow}`).getValues();

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

//行動援護
function inputDbDataKodoengo_() {
  //一覧表シートを開く
  const sheet = SpreadsheetApp.openById(ichiranhyo).getSheetByName('行動援護・同行支援・行動支援');
  const value = sheet.getDataRange().getValues();
  
  //サービス日時を参照して翌月のデータのみを取得する
  //サービス日時は4列目

  //翌月の年月データを取得
  //★先月を想定して実施が必要
  const nextMonth = `${nextYearMonth_().year}/${nextYearMonth_().month}`

  const valueList =[]


  //0はヘッダーのため1から
  for(let i=1;i<value.length;i++){
    if(value[i][4].toLocaleString('ja-JP').slice(0,6)===nextMonth && value[i][8]==='行動援護'){
      valueList.push(value[i])
    }
  }

  //Logger.log(valueList)
  return valueList
}


function connectionJissekiSheetKodoengo_(){

  //利用者リストを開く
  const userListSheet = SpreadsheetApp.openById(userIdList).getSheetByName('利用者リスト');

  const lastRow = userListSheet.getLastRow()

  const userIdColumnValues = userListSheet.getRange(`C2:C${lastRow}`).getValues();

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
    for(let k=0;k<inputDbDataKodoengo_().length;k++){
      if(inputDbDataKodoengo_()[k][1]===userIdColumnValuesList[j][0]){
        const targetRow = userIdColumnValuesList[j][1] ;
        const targetId = userListSheet.getRange(`B${targetRow}`).getValue();
        targetIdList.push(targetId)
      }
    }  
  }

  const targetUniqueIdList = Array.from(new Set(targetIdList))

  return targetUniqueIdList
}

//同行支援
function inputDbDataDokoShien_() {
  //一覧表シートを開く
  const sheet = SpreadsheetApp.openById(ichiranhyo).getSheetByName('行動援護・同行支援・行動支援');
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

  //利用者リストを開く
  const userListSheet = SpreadsheetApp.openById(userIdList).getSheetByName('利用者リスト');

  const lastRow = userListSheet.getLastRow()

  const userIdColumnValues = userListSheet.getRange(`C2:C${lastRow}`).getValues();

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

//行動支援
function inputDbDataKodoshien_() {
  //一覧表シートを開く
  const sheet = SpreadsheetApp.openById(ichiranhyo).getSheetByName('行動援護・同行支援・行動支援');
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

  //利用者リストを開く
  const userListSheet = SpreadsheetApp.openById(userIdList).getSheetByName('利用者リスト');

  const lastRow = userListSheet.getLastRow()

  const userIdColumnValues = userListSheet.getRange(`C2:C${lastRow}`).getValues();

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

//算定時間算出
function addLeadingZero_(number) {
  return (number < 10 ? '0' : '') + number;
}

//------------------------------
//基本的に使用しないが念のためおいておく
//質問リストをヘッダーにする関数（予備用で基本使用しない）
// function questionToHeader_(){
//   //const headerList = formToDatabaseKyotakuJudo_().question;
//   const headerList = formToDatabaseKodo_().question;

//   //ヘッダーリストを一覧表に書き込む
//   //const kodoID = PropertiesService.getScriptProperties().getProperty('KYOTAKUJUDO_SHEET_ID')
//   const kodoID = PropertiesService.getScriptProperties().getProperty('KODO_SHEET_ID')

//   const ss = SpreadsheetApp.openById(kodoID);
//   const sheet = ss.getSheetByName('データベースサンプル');

//   //列方向に書き出す場合はfor文で書く
//   for (let i=0; i<headerList.length;i++){
//     sheet.getRange(1,i+1).setValue(headerList[i])
//   }
// }
