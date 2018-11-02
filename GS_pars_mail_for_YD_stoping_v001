var team_mail='to:mail'; //сюда вводить командную почту
var ss = SpreadsheetApp.getActiveSpreadsheet();    
var sheet = ss.getSheets()[0];

var regExp = new RegExp("[^\d;]")

function get_threads_from_yandex(team_mail) {
  return GmailApp.search('(яндекс.директ/показы ) and (приостановлены по дневному ограничению бюджета) and (to:' + team_mail + ')');
  
};


var threads = get_threads_from_yandex();
Logger.log(threads);

var i=0;

for (thread in threads){
  var subject = threads[i].getFirstMessageSubject().replace(/[^\d;]/g, '').replace(".", '');
  var subject2 = subject;
  Logger.log(subject2);
  var datetime = threads[i].getLastMessageDate();
  sheet.getRange("A"+(i+1)).setValue(subject); // в колонку А пишем название
  sheet.getRange("B"+(i+1)).setValue(datetime); // в колонку B пишем дату время письма 
  i+=1}
