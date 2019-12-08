function manualTransferData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var list1 = sheet.getActiveSheet(); 
  var calendar = CalendarApp.getCalendarById('свой ID');
  
  //правило повторения события (ежегодно)
  var recurrence = CalendarApp.newRecurrence().addYearlyRule();  
  var dayOfWeek = /Sun|Mon|Tue|Wed|Thur|Fri|Sat/;
  
  
  //спросить пользователя о намерениях
  var ui = SpreadsheetApp.getUi();    
  var response = ui.alert('Копировать строку в календарь?', ui.ButtonSet.YES_NO);  
  
  if (response == ui.Button.YES) {
   //выбрать ячейку курсором - выяснить её номер строки
     var activeRange = list1.getActiveRange();  
     var activeRow = activeRange.getRow(); 
  
   //проверка на наличие даты
    var dateCell = list1.getRange(activeRow, 3).getValue();
    var dateString = dateCell.toString();  
    
    var title = list1.getRange(activeRow, 2).getValue();    
    
     if (dateString.match(dayOfWeek) != null){
      //загрузить данные в календарь
       Logger.log("дата есть! - отправить данные!");
       calendar.createAllDayEventSeries(title, dateCell, recurrence);
    } else if (dateString.match(dayOfWeek) == null){         
      Logger.log("Не верные данные в ячейке с датой. Событие не перенесено");
      //покрасить ячейку если там нет даты
      list1.getRange(activeRow, 3).setBackground("#ff8c8c"); 
    }    
  } else {
    Logger.log('Пользователь нажал(а) кнопку "No" или "закрыть"');
  } 
}
