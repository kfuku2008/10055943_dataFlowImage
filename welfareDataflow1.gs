//グローバル変数(簡略化する)
//作成するシート名(グローバル変数として使用)
const targetName = `${nextYearMonth_().year}年${nextYearMonth_().month}月`;

//実績票原本のスプシID
const jissekiGenponSheetNameList = ["居宅介護実績票","重度訪問介護実績票","行動援護実績票","同行支援実績票","行動支援実績票"];
const keikakushoList = ['計画表_居宅介護','計画表_重度訪問介護','計画表_行動援護','計画表_同行支援','計画表_行動支援'];

//フォームの回答データを取得し、一覧表に反映させる
//スクリプトプロパティからフォームIDを呼び出す
const formKodoId = PropertiesService.getScriptProperties().getProperty('FORM_KODO_ID');
const formKyotakuJudoId = PropertiesService.getScriptProperties().getProperty('FORM_KYOTAKUJUDO_ID');

//一覧表(居宅介護・重度訪問)のスプシID
const ichiranhyouDokoShien = PropertiesService.getScriptProperties().getProperty('ICHIRANHYOU_KODO');

//スクリプトプロパティからスプシIDを呼び出す
//一覧表(居宅介護・重度訪問)のスプシID
const ichiranhyouJudo = PropertiesService.getScriptProperties().getProperty('ICHIRANHYOU_KYOTAKUJUDO');

//利用者リストのスプシID
const userIdList = PropertiesService.getScriptProperties().getProperty('USER_SHEETID_LIST');

//スクリプトプロパティからスプシIDを呼び出す
//一覧表(居宅介護・重度訪問)のスプシID
const ichiranhyouKodoshien = PropertiesService.getScriptProperties().getProperty('ICHIRANHYOU_KODO');

//スクリプトプロパティからスプシIDを呼び出す
//一覧表(居宅介護・重度訪問)のスプシID
const ichiranhyouKodoengo = PropertiesService.getScriptProperties().getProperty('ICHIRANHYOU_KODO');

//スクリプトプロパティからスプシIDを呼び出す
//一覧表(居宅介護・重度訪問)のスプシID
const ichiranhyouKyotaku = PropertiesService.getScriptProperties().getProperty('ICHIRANHYOU_KYOTAKUJUDO')

//算出時間の記入時に使用
const targetPeriod = `_${nextYearMonth_().year}年${nextYearMonth_().month}月_`


//翌月の実績記録票のコピーを行う関数(体裁調整前)※ここもランタイムエラー対策つけた方がいいかも
function copyForm(){

  for(let i=0; i<inputUserData_().length; i++){
    //実績記録票のコピーを作成(新しいシート作成→コピー元のデータを張り付け)

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

//シートの体裁調整(列幅を調整)
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


//実績記録票への書き出しを行う関数
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
  //Logger.log(nonEnptyCells)
  //[[], [[通院介助, E24], [身体介護, E33:E35]], [], [[同行支援, E27]], [[行動支援, F33]]]



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

