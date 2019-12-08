function transferToCalendar() {
  var sheet = SpreadsheetApp.openById("свой ID");
  var list1 = sheet.getSheetByName("База клиентов");
  var calendar = CalendarApp.getCalendarById('свой ID');
     
  //дата сегодня
  var today = new Date();
  
  //дата вчера  
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
  //взять последнюю строку в таблице на сегодня
  var todayRow = list1.getLastRow(); 
    
  //проверить изменилось ли число последней строки (есть ли события за вчерашний день)  
  if (yesterdayRow != todayRow){  
    //сформировать диапазон значиний для загрузки в календарь
    var rangeForTransfer = "B" + (1 + (yesterdayRow *1)) + ":C" + todayRow;
    var r = 1 + (yesterdayRow *1);
    var c = 3;
    Logger.log(rangeForTransfer);
        
    //взять массив объектов
    var eventsArray = list1.getRange(rangeForTransfer).getValues();
    //Logger.log(eventsArray);
    
    //отправить массив событий в календарь
    pushEvents(eventsArray, calendar, list1, r, c);  
  
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

function pushEvents(eventsArray, calendar, list1, r, c){    
  //правило повторения события (ежегодно)
    var recurrence = CalendarApp.newRecurrence().addYearlyRule();  
    var dayOfWeek = /Sun|Mon|Tue|Wed|Thu|Fri|Sat/;
  
  //отправка массива данных в календарь
  for (var i=0; i<eventsArray.length; i++){
    //проверка ячейки на наличие даты
    var dateString = eventsArray[i][1].toString();    
    if (dateString.match(dayOfWeek) != null){
      var event = calendar.createAllDayEventSeries(eventsArray[i][0], eventsArray[i][1], recurrence);
    } else if (dateString.match(dayOfWeek) == null){         
      Logger.log("Не верные данные в ячейке с датой. Событие не перенесено");
      list1.getRange(r+i, c).setBackground("#ff8c8c");
    } else if (eventsArray[i][1] != ""){         
      Logger.log("Пустая ячейка с датой. Событие не перенесено");
      list1.getRange(r+i, c).setBackground("#ff8c8c");
    }
  }
}
