//скрипт ежедневной автоматической загрузки данных из базы в календарь
//запускать через тригеры проэкта

function transferToCalendar() {
  var sheet = SpreadsheetApp.openById("Свой ID"); //+
  var list1 = sheet.getSheetByName("База Клиентов"); //+
  var calendar = CalendarApp.getCalendarById('Свой ID'); //+
     
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
    var c = 2;
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
    //проверка ячейки на наличие имени и даты
    var dateCell = eventsArray[i][1];
    var dateString = eventsArray[i][1].toString();  
    var title = eventsArray[i][0];
    
    Logger.log(dateString);
    Logger.log(title)    
    
    if ((dateString.match(dayOfWeek) != null)&&(title != "")){
      //корректировка даты из-за ошибки создания события раньше 1991года
     var yearCell = dateCell.getFullYear();       
       Logger.log(yearCell);       
      
      if (yearCell <= 1991){
        var currentDate = new Date(dateCell);
        currentDate.setDate(currentDate.getDate() + 1);    
      } else {
        currentDate = dateCell;
      }    
      Logger.log(currentDate);        
      var event = calendar.createAllDayEventSeries(eventsArray[i][0], currentDate, recurrence);
    } else if ((dateString.match(dayOfWeek) == null)||(eventsArray[i][0] != "")||(eventsArray[i][1] != "")){         
      Logger.log("Не верные данные в ячейках имени или даты. Событие не перенесено");
      list1.getRange(r+i, c, 1, 2).setBackground("#ff8c8c");     
    } 
  }
}
