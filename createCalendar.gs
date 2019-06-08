function create_calendar() {

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();
var index = 2;
var lastRow = sheet.getLastRow();

for (;index <= lastRow; index++){
  var type = sheet.getRange(index, 1, 1, 1).getValue();
  var taskTitle = sheet.getRange(index, 2, 1, 1).getValue();
  var taskDesc = sheet.getRange(index, 3, 1, 1).getValue();
  var onCalendar = sheet.getRange(index, 4, 1, 1).getValue();
  var originalEstimate = sheet.getRange(index, 5, 1, 1).getValue();
  var loggedEffort = sheet.getRange(index, 6, 1, 1).getValue();
  var startDate = sheet.getRange(index, 7, 1, 1).getValue();
  var endDate = sheet.getRange(index, 8, 1, 1).getValue();
  var status = sheet.getRange(index, 9, 1, 1).getValue();
  
  //var sendInvites = true;
  
  /*Date format should be set correctly in both the sheet and the editor.*/
  
  /* How to format the date 
  startDate_proper = new Date(startDate);
  startDate_proper = Utilities.formatDate(startDate_proper, "GMT+0530", "dd-MM-yyyy HH:mm:ss");
  
  endDate_proper = new Date(endDate);
  endDate_proper = Utilities.formatDate(endDate_proper, "GMT+0530", "dd-MM-yyyy HH:mm:ss");
  
  */  
  
  //Browser.msgBox(type+':'+taskTitle+':'+startDate+':'+onCalendar, Browser.Buttons.OK_CANCEL);
  
  if (onCalendar == 'Yes' && startDate && endDate && status != 'Done')
  {
    //   Browser.msgBox(type+':'+taskTitle+':'+startDate+':'+onCalendar, Browser.Buttons.OK_CANCEL);
    //Browser.msgBox(Utilities.formatDate(temp, "GMT+0530", "dd-MM-yyyy HH:mm:ss"), Browser.Buttons.OK_CANCEL);
    
    //  Browser.msgBox(startDate+':'+endDate+':'+CalendarApp.getCalendarsByName("sparxsys_tasks")[0].getEvents(startDate, endDate)  );
    
    var events =  CalendarApp.getCalendarsByName("sparxsys_tasks")[0].getEvents(startDate, endDate);
    delete_events(events);
    var calendar = CalendarApp.getCalendarsByName("sparxsys_tasks")[0].createEvent(taskTitle, startDate,endDate,{description: taskDesc});
    
    // var calendar = CalendarApp.getCalendarById("sparxsys_task").createEvent(taskTitle, startDate,endDate,{description: taskDesc});
    
  }
  
  // var calendar = CalendarApp.getCalendarById("ravisagar@gmail.com").createEvent(taskTitle, startDate_proper,endDate_proper,{description: taskDesc}); 
  // var calendar = CalendarApp.getCalendarById("ravisagar@gmail.com").createEvent(taskTitle, Utilities.formatDate(startDate_proper, "GMT+0530", "dd-MM-yyyy HH:mm:ss"),Utilities.formatDate(endDate_proper, "GMT+0530", "dd-MM-yyyy HH:mm:ss"),{description: taskDesc});
  
  
}// End of for Loop
  
  
  
}// End of CalendarTest Function



function delete_events(events) {
  
  for(var i=0; i<events.length;i++){
      var ev = events[i];
    
   // Browser.msgBox(ev.getTitle());
//      Logger.log(ev.getTitle()); // show event name in log
      ev.deleteEvent();
    }
  
}
