
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu("Calendar");
  var item = menu.addItem("Authorize", "showSidebar");
  item.addToUi();
}


function showSidebar() {
  var html = HtmlService.createTemplateFromFile('HTML_Sidebar').evaluate()
      .setTitle('Log Tools')
      .setWidth(300);
  SpreadsheetApp.getUi() 
      .showSidebar(html);
}

function getAuthUrl() {
  var authInfo,msg;

  authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);

  msg = 'This spreadsheet needs permission to use Apps Script Services.  Click ' +
    'this url to authorize: <br><br>' + 
      '<a href="' + authInfo.getAuthorizationUrl() +
      '">Link to Authorization Dialog</a>' +      
      '<br><br> This spreadsheet needs to either ' +
      'be authorized or re-authorized.';

  //ScriptApp.AuthMode.FULL is the auth mode to check for since no other authorization mode requires
  //that users grant authorization
  if (authInfo.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.REQUIRED) {
    return msg;
  } else {
    return "No Authorization needed";
  }
}

function calendarAutomation() {
  var allCals = CalendarApp.getAllCalendars();
  var targetCal = CalendarApp.getCalendarById("c_j11vdrbnjr6hhn56tdtk31hrrc@group.calendar.google.com");
  var ws = SpreadsheetApp.getActive().getSheetByName("Data");
  var lr = ws.getLastRow();
  var data = ws.getRange(2, 1, lr-1, 2).getValues();
  var today = Utilities.formatDate(new Date(), "GMT", "MM/dd/yyyy");
  for(var i=0; i<allCals.length; i++){
    var id = allCals[i].getId();
    if(id.indexOf('@import.calendar.google.com')+1>0){
      var cal = CalendarApp.getCalendarById(id);
      var calname = getCalendarName(data, allCals[i].getName());
      var evs = cal.getEventsForDay(new Date());
      for(var j=0; j<evs.length; j++){
        var endD = evs[j].getEndTime();
        var endDt = Utilities.formatDate(subDaysFromDate(endD), "GMT", "MM/dd/yyyy");
        if(today == endDt){
          var startTime = setTimeToDate(endD, "start");
          var endTime = setTimeToDate(endD, "end");
          targetCal.createEvent(calname, startTime, endTime);
        }
      }
    }
  }
}


function subDaysFromDate(date){
   var result = new Date(date.getTime()-(24*3600*1000));
   return result;
}

function getCalendarName(data, name){
  for(var l=0; l<data.length; l++){
    if(data[l][0] == name)return data[l][1];
  }
}

function setTimeToDate(dates, flag){
  var date = new Date(dates.getTime()-(24*3600*1000));
  if(flag == "start"){
    date.setHours(11);
    date.setMinutes(0);
    date.setSeconds(0);
    date.setMilliseconds(0);
    return date;
  }else if(flag == "end"){
    date.setHours(14);
    date.setMinutes(0);
    date.setSeconds(0);
    date.setMilliseconds(0);
    return date;
  }
}
