function autoLoadBase() {
  var sheet = SpreadsheetApp.openById("14j0U6OuX63iu2KjZC9IJXL5QpCeQSqG9G5pxE1zM2Qo");
  var list1 = sheet.getSheetByName("База клиентов");
  var calendar = CalendarApp.getCalendarById('70nhdlr1snpb2na5tfetdd3cdc@group.calendar.google.com');  
  
  //правило повторения события (ежегодно)
  var recurrence = CalendarApp.newRecurrence().addYearlyRule();  
  var dayOfWeek = /Sun|Mon|Tue|Wed|Thu|Fri|Sat/;
  
  var row = 1;     
  //найти последний флажок в колонке и запустить цикл  
  do{
    row++;
  } while (list1.getRange(row, 12).isChecked())  
   Logger.log(row);
  
  //запустить копирование по 1 строке
  for (var i=0; i<10; i++){    
    if (!list1.getRange(row+i, 12).isChecked()){ //нет флажка    
     Logger.log("НЕТ флажка - копирую данные");    
    
    //проверка на наличие даты
    var dateCell = list1.getRange(row+i, 3).getValue();
    var dateString = dateCell.toString();      
    var title = list1.getRange(row+i, 2).getValue();    
    
     if (dateString.match(dayOfWeek) != null){      
       Logger.log("дата есть! - отправить данные!");
       calendar.createAllDayEventSeries(title, dateCell, recurrence); //загрузить данные в календарь      
       list1.getRange(row+i, 12).check(); //поставить флажок 
       
     } else if (dateString.match(dayOfWeek) == null){         
      Logger.log("Не верные данные в ячейке с датой. Событие не перенесено");      
      list1.getRange(row+i, 3).setBackground("#ff8c8c"); //покрасить ячейку если там нет даты  
      list1.getRange(row+i, 12).check(); //поставить флажок 
     }         
  } else {
     Logger.log("Есть флажок - проверяю дальше боксы");    
  }  
}  
}











