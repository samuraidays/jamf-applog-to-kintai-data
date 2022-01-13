function onOpen() {
  SpreadsheetApp
    .getActiveSpreadsheet()
    .addMenu('管理', [{name: '稼働時間を取得', functionName: 'findWorkaholic'},]);
}

const API_URL = 'https://knm.jamfcloud.com/JSSResource/'

function findWorkaholic() {
  // 前月初日から前月末日までを集計
  var start_date = new Date();
  start_date.setMonth(start_date.getMonth()-1);
  start_date.setDate(1);
  var end_date = new Date();
  end_date.setDate(0); // 0を指定すると前月末日になる
  
  var year = start_date.getFullYear();
  var month = ("0"+(start_date.getMonth() + 1)).slice(-2); // 0埋めで2桁に合わせる
  
  // YYYY-MM-DD_YYYY-MM-DD
  var search_range = year + '-' + month + '-' + ('0' + start_date.getDate()).slice(-2) + '_' + year + '-' + month + '-' + ('0' + end_date.getDate()).slice(-2);
  
  // 記録用のスプレッドーシートを開き新しくシートを作成する
  const ss = SpreadsheetApp.openById("XXXXXXXXXXXXXXXXXXXXX");
  var sheet = ss.insertSheet(year + '-' + month);

  var values = new Array();
  values.push(['Name', 'Date', 'Usage(min)']); // First Row
  
  // ClassicAPIはBasic認証なのでCredentialをBase64エンコードする
  // username:passwordはAPIアクセス用のJamfアカウントを作成し指定する
  const auth_data = Utilities.base64Encode('USERNAME:PASSWORD');

  var options = {
    'method' : 'GET',
    'contentType': 'application/json',
    'headers': {'Authorization' : 'Basic ' + auth_data,
                'accept' : 'application/json'},
  };
  
  // 全てのコンピューター一覧を取得
  const response = UrlFetchApp.fetch(API_URL + 'computers', options);
  var cont = JSON.parse(response.getContentText('UTF-8'));
  
  for (var i=0; i<cont.computers.length; i++) {  
    var computer = cont.computers[i];
    
    // 対象のコンピューターの一か月のアプリ使用状況を取得
    const _response = UrlFetchApp.fetch(API_URL + 'computerapplicationusage/id/' + computer.id + '/' + search_range, options);
    var _cont = JSON.parse(_response.getContentText('UTF-8'));
    
    // 個別のコンピュータの情報を取る
    const _comResponse = UrlFetchApp.fetch(API_URL + 'computers/id/' + computer.id, options);
    var _comres = JSON.parse(_comResponse.getContentText('UTF-8'));
   
    // アプリがforeground(アクティブ状態)だった時間を集計する
    var time = 0 // 月ごとに集計したい場合
    for (var j=0; j<_cont.computer_application_usage.length; j++) {
      var day_usage = _cont.computer_application_usage[j];
      //var time = 0; // 日ごとに集計したい場合
      for (var k=0; k<day_usage.apps.length; k++) {
        time += day_usage.apps[k].foreground;
      }
      //values.push([computer.name, day_usage.date, time]); // 日ごとに集計したい場合
    }
    
    // emailを取る
    var email = _comres.computer.location.email_address;

    values.push([email, year + '-' + month, time]); // 月ごとに集計したい場合
  }
  
  // 結果をスプレッドシートに保存
  sheet.getRange(1, 1, values.length, 3).setValues(values);
  
  return;
}

