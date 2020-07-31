  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var arr = sheet.getRange(`A2:D${lastRow}`).getValues();

function myFunction() {
  if(lastRow < 2) throw 'データが空です！1行以上のデータを追加してください';
  var mail = "";
  var companyName = "";
  var rowNum;
  for(var i = 0;i < arr.length;i++){
    if(arr[i][3] == "") {
      companyName = arr[i][1];
      mail = arr[i][2];
      rowNum = i;
      break;
    }else{
      if(i + 2 == lastRow){
        console.log("完了です");
      }
    }
  }  
  var title = "すごいメールを送信致します";
  var body = `
  ${companyName}御中
  
  すごいメール本文です。
  どうぞよろしくおねがいします。
  `  
  GmailApp.sendEmail(mail, title, body);
  sheet.getRange(`D${rowNum + 2}:D${rowNum + 2}`).setValue("送付済");
}

function getFail(){
  //for(var number = 500;number < 1501; number = number + 500){}
  var threads = GmailApp.search("検索したいメール文面", 0, 500);
  var messagesForThreads = GmailApp.getMessagesForThreads(threads);
  console.log(arr);
  var arrays = messagesForThreads.forEach(function(messages){
    if(messages[1]){
      var body = messages[1].getBody();
      var before = `メールを受信できないアドレスであるため、メールは <a style='color:#212121;text-decoration:none'><b>`
      var after = "</b></a> に配信されませんでした。";
      var failAddress = body.split(before)[1].split(after)[0];
      console.log(failAddress == "y-kimur@call.jp");
      for(var i = 0;i < arr.length;i++){
        for(var j = 0;j < arr[i].length;j++){
          if(arr[i][2] == failAddress && sheet.getRange(`E${i + 2}`).getValue() == ""){
            sheet.getRange(`E${i + 2}`).setValue("送信エラー");
          }
        }
      }
    }
  });
}
