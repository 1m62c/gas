// シート情報取得
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sh = ss.getSheetByName('団員名簿');
var lastRow = sh.getLastRow();

// api情報取得
var apiToken = PropertiesService.getScriptProperties().getProperty('apiToken');
var bandKey = PropertiesService.getScriptProperties().getProperty('bandKey');

// 全ての情報取得
var allDataColomn = sh.getRange(1, 1, lastRow, 6);
var allData = allDataColomn.getValues();

// aaaa/bb/cc型に変換した全ての情報取得
var convertedAllData = fromDateToString(allData);


function fromDateToString(data) {
  for (var i = 0; i < data.length; i++) {
    for (var j = 2; j <= 3; j++) {
      var year = data[i][j].getFullYear();
      var month = data[i][j].getMonth() + 1;
      var day = data[i][j].getDate();
      data[i][j] = `${year}/${month}/${day}`;
    }
  }
  return data;
}


function statusJudge(compareYear, compareMonth) {

  // 休団年月日を配列に
  var breakDatesColumn = sh.getRange(1, 3, lastRow);
  var breakDates = breakDatesColumn.getValues();

  // 復団年月日を配列に
  var returnDatesColumn = sh.getRange(1, 4, lastRow);
  var returnDates = returnDatesColumn.getValues();
  
  // 一人ずつ団員を取り出して繰り返す
  for (var i = 1; i <= lastRow; i++) {

    // その団員の休団・復団年月日
    var breakDate = breakDates[i - 1][0];
    var returnDate = returnDates[i - 1][0];

    // 初期状態
    var status = '活動中';

    if (breakDate && returnDate) { // 日付が存在する場合のみ処理（入力漏れ防止）

      // 休団・復団の年と月を変数に入れる
      var breakYear = breakDate.getFullYear();
      var breakMonth = breakDate.getMonth() + 1;
      var returnYear = returnDate.getFullYear();
      var returnMonth = returnDate.getMonth() + 1;

      // 休団判定
      if (breakYear < compareYear) {
        if (returnYear < compareYear) {
          status = '活動中';
        } else if (compareYear === returnYear) {
          if (returnMonth <= compareMonth) {
            status = '活動中';
          } else {
            status = '休団中';
          }
        } else {
          status = '休団中';
        }
      } else if (breakYear === compareYear) {
        if (compareYear === returnYear) {
          if (compareMonth < breakMonth) {
            status = '活動中';
          } else if (breakMonth <= compareMonth && compareMonth < returnMonth) {
             status = '休団中';
          } else {
            status = '活動中';
          }
        } else {
          if (compareMonth < breakMonth) {
            status = '活動中';
          } else {
            status = '休団中';
          }
        }
      }
    }

    // 結果をF列に設定
    sh.getRange(i, 6).setValue(status);

  }

}


function now() {
  // 現在の年月取得
  var now = new Date();
  var nowYear = now.getFullYear();
  var nowMonth = now.getMonth() + 1;

  statusJudge(nowYear, nowMonth);
}


function research() {
  // 入力した年・月を取得
  var researchDateColumn = sh.getRange('I1');
  var researchDate = researchDateColumn.getValue();
  var researchYear = researchDate.getFullYear();
  var researchMonth = researchDate.getMonth() + 1;

  statusJudge(researchYear, researchMonth);
}


function getBand() {
  // band情報取得
  var payload = 
  {
    'access_token' : apiToken
  };

  var options =
  {
    'method' : 'get',
    'payload' : payload
  };

  var getBandRes = UrlFetchApp.fetch('https://openapi.band.us/v2.1/bands', options);
  Logger.log(getBandRes);
  var getBandName = getBandRes.result_data.bands[0].band_key;
  Logger.log(getBandName)
}


function writePost(message) {
  // 投稿
  var payload = 
  {
    'access_token' : apiToken,
    'band_key' : bandKey,
    'content' : message,
    'do_push' : true
  };

  var options =
  {
    'method' : 'post',
    'payload' : payload
  };

  var writePostRes = UrlFetchApp.fetch('https://openapi.band.us/v2.2/band/post/create', options);
  Logger.log(writePostRes);
}


function breakMemberPost() {
  var breakMember = [];

  for (var i = 0; i < convertedAllData.length; i++) {
    if (convertedAllData[i][5] == '休団中') {
      breakMember.push(`${convertedAllData[i][0]} : ${convertedAllData[i][1]} (${convertedAllData[i][3]})`);
    }
  }
  var breakMemberList = breakMember.join('\n');

  var breakMemberNotify = `【今月の休団者と復団予定日一覧】\n${breakMemberList}`;
  Logger.log(breakMemberNotify);
  writePost(breakMemberNotify);
}