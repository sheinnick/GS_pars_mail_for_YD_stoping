/* Как запустить:
1. Копируем этот код
2. создаем новую гугл таблицу
3. открываем скрипт эдитор http://joxi.ru/Drl1OdXiv1ggJm
4. http://joxi.ru/vAWNYB9T1400E2 
4.1. называем проект как угодно
4.2. удаляем тут всё и вставляем этот код
5. заполняем projname  http://joxi.ru/Dr8xnqBH4b88Or — название аккаунта, берем из письма http://joxi.ru/4AkkOWnIy4qpgA можно и пустым оставить
6. жмем на плей, нажимаем review и со всем соглашаемся http://joxi.ru/ZrJZqw9U97BB4r  
7. увидим error http://joxi.ru/YmE6kwBC0Xyy42 — всё ок, нажимем сохранить и запускаем еще раз
8. ждём. когда вот это http://joxi.ru/V2V8ew9ixno5Kr исчезнет
9. success — в таблице появились данные

TODO:
1. Сделать чтобы грузилось больше 500 строк (сейчас это ограничение метода GmailApp.search)
2. сделать, чтобы скрипт запускался автоматически каждый день и дописывал остановки за вчера
*/

var projname ='';        // название аккаунта, берем из письма http://joxi.ru/4AkkOWnIy4qpgA
var ss = SpreadsheetApp.getActiveSpreadsheet();    
var sheet = ss.getSheets()[0];
var header = [["Логин", "Дата","Время","Кампания","ID кампании"]]  //названия колонок
sheet.getRange('A:E').clearContent();  //очищаем лист , на самом деле только колонки A:E
sheet.getRange("A1:E1").setValues(header); //записываем шапку таблицы

function get_threads_from_yandex(projname_arg) {
  return GmailApp.search('(яндекс.директ/показы ) and (приостановлены по дневному ограничению бюджета) and ('+ projname_arg +')')
};


var threads = get_threads_from_yandex(projname);
var i=0;

for (thread in threads)
{
  // var subject = threads[i].getFirstMessageSubject().replace(/[^\d;]/g, '').replace(".", ''); // парсим кампанию из темы, сейчас не используется
  var datetime = threads[i].getLastMessageDate(); // дата из даты письма
  var messages = threads[i].getMessages();        //вытаскиваем все сообщения в ветке
  var message = messages[0].getPlainBody();       //вытаскиваем текст первого сообщения
  var q = message.split('\n', 3)                  //разбиваем письмо на строки (берем всего 4, там вся нужная инфа)
  // ↓ пишем инфу в таблицу
  sheet.getRange("A"+(i+2)).setValue(q[0].match(RegExp("([a-zA-Z0-9\\-\\.]*)!"))[1]); //accname
  sheet.getRange("B"+(i+2)).setValue(datetime.getFullYear()+'-'+ (datetime.getMonth()< 10 ? '0' : '') + datetime.getMonth() + '-' + (datetime.getDate()< 10 ? '0' : '') + datetime.getDate()); //date
  sheet.getRange("C"+(i+2)).setValue(q[2].match(RegExp("\\s[0-9:]*\\s")).toString().replace(" ","")); //time
  sheet.getRange("D"+(i+2)).setValue(q[2].match(RegExp("\\((.*)\\)\\s"))[1]); //campaignname
  sheet.getRange("E"+(i+2)).setValue(q[2].match(RegExp("N[0-9]*")).toString().replace("N","")); //campaignid 
  i+=1
}

