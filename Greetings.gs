//может выкидывать ошибку из-за преышения допустимой длины URL из-за большого кол-ва именинников или длинного сообщения

function main(){
  getClientsFromCalendar();
  pullGreeting();
  sendText(greetingRandomize, clientString);  
}

function getClientsFromCalendar() {
  //подключение к календарю
   var calendar = CalendarApp.getCalendarById('свой ID');
     
  //дата сегодня
  var today = new Date();
  
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
  var sheet = SpreadsheetApp.openById("свой ID");  
  var list2 = sheet.getSheetByName("Поздравления Telegram");
  
  //перемешивает строки в диапазоне и берет значение первой
  //если строку удалили со временем, то потом её пустая ячейка окажется внизу диапазона
  var lastRow = list2.getLastRow();
  var greeting = list2.getRange(2, 1, lastRow, 1); 
  return greetingRandomize = greeting.randomize().getValue();
}

function sendText(greetingRandomize, clientString) {
  var botID = 'свой ID';
  var chatID = 'свой ID'; 
  var text = encodeURIComponent('Сегодня празднует свой день рождения \uD83C\uDF89 \uD83C\uDF88 \n' + clientString + '\n' + greetingRandomize + '\uD83C\uDF81 \uD83C\uDF82 \uD83C\uDF70');
  var createLink = "https://api.telegram.org/bot" + botID + "/sendMessage?chat_id=" + chatID + "&text=" + text;  
  var loadLink = UrlFetchApp.fetch(createLink);
}
