function onOpen() {    
  var ui = SpreadsheetApp.getUi();   
  ui.createMenu("Меню администратора")
  .addItem("Копирование данных в Google Calendar", "autoLoadBase")  
  .addToUi();  
}
