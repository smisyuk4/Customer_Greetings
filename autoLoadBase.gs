function autoLoadBase() {
  var sheet = SpreadsheetApp.openById("текст"); 
  var list1 = sheet.getSheetByName("База Клиентов");  
  var calendar = CalendarApp.getCalendarById('текст'); 
     
  //правило повторения события (ежегодно)
  var recurrence = CalendarApp.newRecurrence().addYearlyRule();  
  var dayOfWeek = /Sun|Mon|Tue|Wed|Thu|Fri|Sat/;
  
  var row = 2;     
  //найти последний флажок в колонке и запустить цикл  
  do{
    row++;
  } while (list1.getRange(row, 20).isChecked())  
   Logger.log(row);
  
  //запустить копирование по 1 строке
  for (var i=0; i<50; i++){    
    if (!list1.getRange(row+i, 20).isChecked()){ //нет флажка    
     Logger.log("НЕТ флажка - копирую данные");    
    
    //проверка на наличие даты и имени
    var dateCell = list1.getRange(row+i, 3).getValue();
    var dateString = dateCell.toString();      
    var title = list1.getRange(row+i, 2).getValue();         
    
     if ((dateString.match(dayOfWeek) != null)&&(title != "")){  
       //корректировка даты из-за ошибки часов
       var cellTime = dateCell.getHours(); //23 (плохо) или 00 (хорошо) 
       var currentDate;   
    
       if (cellTime == 0){ 
         currentDate = dateCell;
       } else if (cellTime != 0){ 
         currentDate = new Date(dateCell);
         currentDate.setDate(currentDate.getDate()+1);     
       }    
       Logger.log(currentDate);          
       
       Logger.log("дата есть! - отправить данные!");
       calendar.createAllDayEventSeries(title, currentDate, recurrence); //загрузить данные в календарь      
       list1.getRange(row+i, 20).check(); //поставить флажок 
       
     } else if ((dateString.match(dayOfWeek) == null)||(title == "")){         
      Logger.log("Не верные данные в ячейке с датой. Событие не перенесено");      
      list1.getRange(row+i, 2, 1, 2).setBackground("#ff8c8c"); //покрасить ячейки если там нет даты или имени  
      list1.getRange(row+i, 20).check(); //поставить флажок 
     }         
  } else {
     Logger.log("Есть флажок - проверяю дальше боксы");    
  }  
}  
}











