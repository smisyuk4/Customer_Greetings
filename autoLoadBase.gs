function autoLoadBase() {
  var sheet = SpreadsheetApp.openById("Свой ID");
  var list1 = sheet.getSheetByName("База клиентов");
  var calendar = CalendarApp.getCalendarById('Свой ID');  
  
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
  for (var i=0; i<20; i++){    
    if (!list1.getRange(row+i, 12).isChecked()){ //нет флажка    
     Logger.log("НЕТ флажка - копирую данные");    
    
    //проверка на наличие даты и имени
    var dateCell = list1.getRange(row+i, 3).getValue();
    var dateString = dateCell.toString();      
    var title = list1.getRange(row+i, 2).getValue();      
    
     if ((dateString.match(dayOfWeek) != null)&&(title != "")){    
       //корректировка даты из-за ошибки часов         
       var cellTime = dateCell.getHours(); //23 (плохо) или 00 (хорошо)
       Logger.log(cellTime);
   
       if (cellTime == 23){
         var currentDate = new Date(dateCell);
         currentDate.setDate(currentDate.getDate());    
       } else {
         currentDate = dateCell;
       }    
       Logger.log(currentDate);
       
       calendar.createAllDayEventSeries(title, currentDate, recurrence); //загрузить данные в календарь      
       list1.getRange(row+i, 12).check(); //поставить флажок 
       
     } else if ((dateString.match(dayOfWeek) == null)||(title == "")){         
      Logger.log("Не верные данные в ячейке с датой. Событие не перенесено");      
      list1.getRange(row+i, 2, 1, 2).setBackground("#ff8c8c"); //покрасить ячейки если там нет даты или имени  
      list1.getRange(row+i, 12).check(); //поставить флажок 
     }         
  } else {
     Logger.log("Есть флажок - проверяю дальше боксы");    
  }  
}  
}











