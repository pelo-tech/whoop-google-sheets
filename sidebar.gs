function onOpen() {
   SpreadsheetApp.getUi()  
      .createMenu('📈 Whoop')
      .addItem('🔒 Login', 'showSidebar')
      .addItem('🔄 Reload Data (Legacy)', 'whoop_rebuild_history')
      .addItem('📖 Load Incremental (Legacy)', 'whoop_get_incremental_history')
      .addItem('🔄 Reload Data (⭐ New/V1)', 'v1_rebuild_history')
      .addItem('📖 Load Incremental (⭐ New/V1)', 'v1_get_incremental_history')
      .addItem('🫀 Load Heartrate Data', 'showHeartrateSidebar')
      .addToUi();
}

function handleSidebarSubmit(obj){
  var results={};
   if(!obj.username || obj.username.length < 5 ||
      !obj.password  || obj.password.length <5 ) {
     return {"error":"Username and password are both required"};
   } 
  var results=handleLoginRequest(obj.username,obj.password);
  // for reasons I don't understand, Google has a hard time serializing this remotely 
  // to HTML calling this via google.script.run, but this fixes the issue.
  //    o
  // -\/^\/-
  // Whatever!
  return  JSON.parse(JSON.stringify(results));
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('login.html')
      .setTitle('Whoop Login')
      .setWidth(320);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);
}


function showHeartrateSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('heartrate.html')
      .setTitle('Heart Rate Lookup')
      .setWidth(320);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showSidebar(html);
}



function loadHeartRate(startDate,endDate,granularity){
  var  whoopSheet = SpreadsheetApp.getActive().getSheetByName(HEART_RATE_SHEET_NAME);
  whoopSheet.clear();
 var result=whoop_get_heart_rate(new Date(startDate), new Date(endDate), granularity);
  SpreadsheetApp.getUi().alert("Heart Rate data processed: "+result+ " measurements loaded into the Heart Rate tab");
  return result;
}