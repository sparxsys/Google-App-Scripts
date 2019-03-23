function fetchSelectedRangeValues() {
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheetvalues = sheet.getActiveRange().getValues(); 
  
  Browser.msgBox(sheetvalues);
  
  for(i=0; i<sheetvalues.length; i++) {
   
    Browser.msgBox(sheetvalues[i]);
    
  }
  
  
}
