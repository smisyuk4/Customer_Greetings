function manualTransferData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var list1 = sheet.getActiveSheet(); 
  var calendar = CalendarApp.getCalendarById('Свой ID');
  
  //правило повторения события (ежегодно)
  var recurrence = CalendarApp.newRecurrence().addYearlyRule();  
  var dayOfWeek = /Sun|Mon|Tue|Wed|Thu|Fri|Sat/;
  
  
  //спросить пользователя о намерениях
  var ui = SpreadsheetApp.getUi();    
  var response = ui.alert('Копировать строку в календарь?', ui.ButtonSet.YES_NO);  
  
  if (response == ui.Button.YES) {
   //выбрать ячейку курсором - выяснить её номер строки
     var activeRange = list1.getActiveRange();  
     var activeRow = activeRange.getRow(); 
  
   //проверка на наличие даты и имени
    var dateCell = list1.getRange(activeRow, 3).getValue();
    var dateString = dateCell.toString();      
    var title = list1.getRange(activeRow, 2).getValue();    
    
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
    
     if ((dateString.match(dayOfWeek) != null)&&(title != "")){
      //загрузить данные в календарь
       Logger.log("дата есть! - отправить данные!");
       calendar.createAllDayEventSeries(title, currentDate, recurrence);
    } else if ((dateString.match(dayOfWeek) == null)|| (title == "")){         
      Logger.log("Не верные данные в ячейке с датой и имени. Событие не перенесено");
      //покрасить ячейки
      list1.getRange(activeRow, 2, 1, 2).setBackground("#ff8c8c"); 
    }    
  } else {
    Logger.log('Пользователь нажал(а) кнопку "No" или "закрыть"');
  } 
}

