function transferToCalendar() {
  var sheet = SpreadsheetApp.openById("14j0U6OuX63iu2KjZC9IJXL5QpCeQSqG9G5pxE1zM2Qo");
  var list1 = sheet.getSheetByName("База клиентов");
  var calendar = CalendarApp.getCalendarById('70nhdlr1snpb2na5tfetdd3cdc@group.calendar.google.com');
     
  //дата сегодня
  var today = new Date();
  
  //дата вчера
  //var lastDay = new Date(2019,11,4,0,0,0);
  var yesterday = new Date();
     yesterday.setDate(today.getDate()-1);  
  Logger.log(yesterday);
  
  
  
  //получить массив событий
  var eventsLasDay = calendar.getEventsForDay(yesterday);
  
  //поиск нужного события  
  var yesterdayRow;
  var eventID;
  for (var i=0; i<eventsLasDay.length; i++){    
    var titleString = eventsLasDay[i].getTitle();
    var titleArray = titleString.split(" ");
    for (var j=0; j<titleArray.length; j++){
         if (titleArray[j] == "dinamicRow"){           
            yesterdayRow = titleArray[j+1];
            eventID = eventsLasDay[i].getId();
         } 
   }
    
  }  
 // Logger.log(tomorowRow);
 // Logger.log(eventID);  
  
  //взять последнюю строку в таблице на сегодня
  var todayRow = list1.getLastRow(); 
  Logger.log(todayRow);
    
  //проверить изменилось ли число последней строки (есть ли события за вчерашний день)  
  if (yesterdayRow != todayRow){  
    //сформировать диапазон значиний для загрузки в календарь
    var rangeForTransfer = "B" + (1 + (yesterdayRow *1)) + ":C" + todayRow;
    Logger.log(rangeForTransfer);
    
    //взять массив объектов начиная с 32 строки вверх до 27 (5шт)
    var eventsArray = list1.getRange(rangeForTransfer).getValues();
    //Logger.log(eventsArray);
    
    //отправить массив событий в календарь
    pushEvents(eventsArray, calendar);  
  
    //добавить в календарь на сегодня новое событие "dinamicRow **" 
    var dinamicRow = todayRow;
    Logger.log(dinamicRow);  
    var dinamicRowPushToCalendar = calendar.createAllDayEvent("dinamicRow " + dinamicRow, today); 
    //удалить не нужное событие в календаре  
    calendar.getEventById(eventID).deleteEvent();     
    
    } else {
      Logger.log("Количество строк не изменилось");
      //перенести старую запись в календаре на новый день
      calendar.getEventById(eventID).setAllDayDate(today);
    }  
}

function pushEvents(eventsArray, calendar){
  //правило повторения события (ежегодно)
    var recurrence = CalendarApp.newRecurrence().addYearlyRule(); 
  
  //отправка массива данных в календарь
  for (var i=0; i<eventsArray.length; i++){
    //  Logger.log(eventsArray[i]);      
    var event = calendar.createAllDayEventSeries(eventsArray[i][0], eventsArray[i][1], recurrence);
    }
}
