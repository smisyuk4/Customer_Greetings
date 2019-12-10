function report(sheet, clientString, today) { 
  var list4 = sheet.getSheetByName("Отчёт");
  var lastRow = list4.getLastRow();
  list4.getRange(lastRow+1, 1).setValue(today);
  list4.getRange(lastRow+1, 2).setValue(clientString);
}
