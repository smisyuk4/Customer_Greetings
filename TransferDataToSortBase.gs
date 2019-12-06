function copyActiveRow(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var list1 = sheet.getActiveSheet(); 
  var list3 = sheet.getSheetByName("База клиентов (Sorted)");  
  
  var ui = SpreadsheetApp.getUi();    
  var response = ui.alert('Ты удалила старую запись в базе Sort (если она была там)?', ui.ButtonSet.YES_NO);  //
  
  if (response == ui.Button.YES) {
    //запуск копирования и сортировки
     var activeRange = list1.getActiveRange();  
     var activeRow = activeRange.getRow(); 
  
   //проверка на пустую ячейку с датой
    var dirtyDateCell = list1.getRange(activeRow, 3).getValue();
  if (dirtyDateCell == ""){
    Logger.log("empty cell");
    list1.getRange(activeRow, 3).setBackground("#F499C0").setValue("Не известно");
  } else {
    Logger.log("not empty cell");
    transferDataAndSort (list1, list3, activeRow);
  }  
  } else {
    Logger.log('Пользователь нажал(а) кнопку "No" или "закрыть"');
  }
  
}

function copyLastRow() {
  var sheet = SpreadsheetApp.openById("14j0U6OuX63iu2KjZC9IJXL5QpCeQSqG9G5pxE1zM2Qo");
  var list1 = sheet.getSheetByName("База клиентов"); 
  var list3 = sheet.getSheetByName("База клиентов (Sorted)");   
  
  var dirtylastRow = list1.getLastRow();
  //проверка на пустую ячейку с датой
  var dirtyDateCell = list1.getRange(dirtylastRow, 3).getValue();
  if (dirtyDateCell == ""){
    Logger.log("empty cell");
    list1.getRange(dirtylastRow, 3).setBackground("#F499C0").setValue("Не известно");
  } else {
    Logger.log("not empty cell");
    transferDataAndSort (list1, list3, dirtylastRow);
  }
  
}
  
function transferDataAndSort (list1, list3, dirtylastRow){  
  //взять массив данных
  var dirtyArr = list1.getRange(dirtylastRow, 1, 1, 4).getValues();
  Logger.log(dirtyArr); 
  
  var dirtyDate = dirtyArr[0][2];
  Logger.log("dirtyDate _______" + dirtyDate); //date
  var dirtyD = dirtyDate.getDate();
  var dirtyM = dirtyDate.getMonth() + 1;
  Logger.log("dirtyDate_" + dirtyD + " dirtyMonth_" + dirtyM); 
  
  
  var sortDateStartRow = 2;
  var sortDateLastRow = list3.getLastRow();
  
  //начало цикла проверок
  while (true){
    
   var sortDate = list3.getRange(sortDateStartRow, 3).getValue();
  Logger.log("sortDate " + sortDate);
  var sortDateCorrentFormat = new Date(sortDate);
  var sortD = sortDateCorrentFormat.getDate();
  var sortM = sortDateCorrentFormat.getMonth() + 1;
 //   
  Logger.log("sortDateCorrentFormat _______" + sortDateCorrentFormat);
    
  Logger.log("sortDate_" + sortD + " sortMonth_" + sortM);
  

  if (list3.getRange(sortDateStartRow, 3).getValue() == ""){
    list3.getRange(sortDateStartRow, 1, 1, 4).setValues(dirtyArr); 
    return false;
    
  } else if((dirtyD == sortD) && (dirtyM == sortM)){
    //добавить строку ниже сравниваемой ячейки
    list3.insertRows(sortDateStartRow + 1);  
    list3.getRange(sortDateStartRow + 1, 1, 1, 4).setValues(dirtyArr); 
    return false;
  
  } else if((dirtyD > sortD) && (dirtyM == sortM)){ //проверить
    //перейти на строку ниже   
    sortDateStartRow++;       
  
  } else if((dirtyD < sortD) && (dirtyM == sortM)){
    //добавить строку выше сравниваемой ячейки
    list3.insertRows(sortDateStartRow);  
    list3.getRange(sortDateStartRow, 1, 1, 4).setValues(dirtyArr);
    return false;  
    
  } else if((dirtyD >= sortD) && (dirtyM < sortM)){
    //добавить строку выше сравниваемой ячейки
    list3.insertRows(sortDateStartRow);  
    list3.getRange(sortDateStartRow, 1, 1, 4).setValues(dirtyArr);
    return false; 
  
  } else if((dirtyD <= sortD) && (dirtyM > sortM)){
  //перейти на строку ниже   
    sortDateStartRow++;
  
  } else if((dirtyD >= sortD) && (dirtyM > sortM)){  
  //перейти на строку ниже   
    sortDateStartRow++;
  
  } else if((dirtyD <= sortD) && (dirtyM < sortM)){
  //добавить строку выше сравниваемой ячейки
    list3.insertRows(sortDateStartRow);  
    list3.getRange(sortDateStartRow, 1, 1, 4).setValues(dirtyArr);
    return false;   
    
  } else if((dirtyD > sortD) && (dirtyM == sortM)){ //проверить
   //добавить строку ниже сравниваемой ячейки
    list3.insertRows(sortDateStartRow + 1);  
    list3.getRange(sortDateStartRow + 1, 1, 1, 4).setValues(dirtyArr); 
    return false;
  }
  }
}
