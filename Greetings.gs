/*
Возможны проблемы из-за количества символов в строке createLink и будет ошибка "Превышен лимит: Длина URL в URLFetch"
-вариант: разделить сообщение на 2 части(вступление + именинники; поздравление)

Повторяемость мероприятий в календаре длится до 2023 года. Не вижу инструмента как продлить этот период.
-вариант: в 2023 году перезалить базу клиентов, чтобы у меропритий обновилась длительность ежегодных повторений

*/

function main(){
  getClientsFromCalendar();
  pullGreeting();
  
  //проверка на наличие именинников
  if(clientString != ''){
    sendText(greetingRandomize, clientString);  
    report(sheet, clientString, today);
  } 
}

function getClientsFromCalendar() {
  //подключение к календарю  
  var calendar = CalendarApp.getCalendarById('текст');
     
  //дата сегодня
  today = new Date();
  //today = new Date(2020, 01, 29);//29 февраля 2020 - день когда нет именинников
  
  //массив событий из календаря
  var clientArr = calendar.getEventsForDay(today);
  
  //взять название каждого события и соеденить в строку
   clientString = "";  
   for (var i=0; i<clientArr.length; i++){    
    clientString += clientArr[i].getTitle();    
    if (i < clientArr.length-1){ //если слово не последнее, то добавить зяпятую и пробел
      clientString += ", ";
    }
  }   
  return clientString;  
}

function pullGreeting(){
  sheet = SpreadsheetApp.openById("текст");  
  var list2 = sheet.getSheetByName("Поздравления Telegram");
  
  //перемешивает строки в диапазоне и берет значение первой
  //если строку удалили со временем, то потом её пустая ячейка окажется внизу диапазона
  var lastRow = list2.getLastRow();
  var greeting = list2.getRange(2, 1, lastRow, 1); 
  return greetingRandomize = greeting.randomize().getValue();
}

function sendText(greetingRandomize, clientString) {
  var botID = 'текст';
  var chatID = 'текст'; 
 // var chatID = 'текст'; //личный чат со мной
  var text = encodeURIComponent('Сегодня празднует свой день рождения \uD83C\uDF89 \uD83C\uDF88 \n' + clientString + '\n' + greetingRandomize + '\uD83C\uDF81 \uD83C\uDF82 \uD83C\uDF70');
  var createLink = "https://api.telegram.org/bot" + botID + "/sendMessage?chat_id=" + chatID + "&text=" + text;  
  Logger.log(createLink);
  var loadLink = UrlFetchApp.fetch(createLink);
}


  
