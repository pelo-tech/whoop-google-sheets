function whoop_setup(){
  var ui = SpreadsheetApp.getUi();
  var credentials=promptForCredentials("Enter Whoop E-mail","Enter Whoop Password (won't be stored)");
  if (credentials == null){
    console.log("Login Aborted");
    ui.alert("Whoop authentication cancelled.");
    return;
  }
  handleLoginRequest(credentials.username, credentials.password)
  ui.alert("Your whoop tokens have been obtained and are stored on your CONFIG tab.");
}

function handleLoginRequest(username, password){
  var data=whoop_login(username,password);
  processWhoopToken(data);
  return data;
}

function processWhoopToken(data){
  var now=new Date().getTime();
  var sheet = SpreadsheetApp.getActive().getSheetByName(CONFIG_SHEET_NAME);
  sheet.getRange(WHOOP_ACCESS_TOKEN_CELL).setValue(data.access_token); 
  sheet.getRange(WHOOP_REFRESH_TOKEN_CELL).setValue(data.refresh_token); 
  sheet.getRange(WHOOP_ID_CELL).setValue(data.user.id); 
  sheet.getRange(WHOOP_USERNAME_CELL).setValue(data.user.username); 
  var now=new Date().getTime();
  var expires=now + (data.expires_in * 1000);
  console.log("Expires "+expires);
  sheet.getRange(WHOOP_TOKEN_EXPIRY).setValue(expires);
  sheet.getRange(WHOOP_TOKEN_LAST_REFRESH).setValue(now);
  sheet.getRange(WHOOP_TOKEN_EXPIRY_READABLE).setValue(new Date(expires));
  sheet.getRange(WHOOP_TOKEN_LAST_REFRESH_READABLE).setValue(new Date(now));
  console.log("Token expires:"+new Date(expires));
  return data;
}

function whoop_refresh_token_if_needed(force){
  console.log("Checking if whoop token refresh is needed.");
  var config=getConfigDetails();
  var now=new Date().getTime();
  if(now>config.whoop.token_expiry || force){
    console.log("Expired Token (force="+force+") Expires on "+new Date(config.whoop.token_expiry));
    var data=whoop_refresh_token();
    console.log("Got Refreshed Token data:"+JSON.stringify(data));
    processWhoopToken(data);
  } else {
      console.log("Token refresh not needed. Current token expires on "+new Date(config.whoop.token_expiry));
  }
}
  

function whoop_rebuild_history(){
  var config=getConfigDetails();
  if(config.whoop.access_token ==null || config.whoop.access_token == ""){
    SpreadsheetApp.getUi().alert("Please click the Login button before loading data");
    return;
  }
  var  whoopSheet = SpreadsheetApp.getActive().getSheetByName(WHOOP_SHEET_NAME);
  whoopSheet.clear();
  whoopSheet.getRange('A1').setValue("Loading... "+new Date().toString());

  var start=new Date();
  start.setDate(start.getDate()-config.whoop.history_size);
  var end=new Date();
  var rows=whoop_get_history(start, end);
  whoopSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  whoopSheet.autoResizeRows(1,rows.length);
  whoopSheet.autoResizeColumns(1,rows[0].length);
  var  confSheet = SpreadsheetApp.getActive().getSheetByName(CONFIG_SHEET_NAME);
  confSheet.getRange(LAST_UPDATED_CELL).setValue(new Date());
  confSheet.getRange(RECORD_COUNT_CELL).setValue(rows.length-1);
}

function localDateString(dt){
  var config=getConfigDetails();
  return Utilities.formatDate(new Date(dt),config.whoop.timezone,'yyyy-MM-dd');
}

function whoop_get_incremental_history(){
  var config=getConfigDetails();
  var whoopSheet = SpreadsheetApp.getActive().getSheetByName(WHOOP_SHEET_NAME);
  var dateColumn=whoopSheet.getRange("A1:A").getValues().filter(String);
  if(dateColumn.length<2){
    console.log("Rebuilding history. No valid data in place");
    whoop_rebuild_history();
    return;
  }
 
  
  // Annlying Date Formatting requires us to stringify values
  var dateStrings=dateColumn.map(
    function(d,idx,arr){
      return (typeof d!='object')?d:localDateString(d);
                       });
  
  var startDate=whoopSheet.getRange("A"+dateColumn.length).getValue();
  startDate=new Date(startDate);
  var endDate=new Date(); // now
  console.log("Getting history from Start Date "+localDateString(startDate)+" to end date "+ localDateString(endDate));
  var history=whoop_get_history(startDate,endDate);
  console.log("Last Value is "+localDateString(startDate));
  var colCount=history[0].length; // even if its headers
  history.shift(); // remove headers
  console.log("History Size: "+history.length-1);
  if (history.length>0){
    var firstDate=history[0][0];
    console.log("First Date: "+localDateString(firstDate));
    console.log("Date Strings :"+JSON.stringify(dateStrings));
    var idx=dateStrings.indexOf(localDateString(firstDate));
    idx+=1 ; // correct index to be a 'row number'
    console.log("Found at row number:"+idx);
    console.log("History Length: "+history.length);
    whoopSheet.getRange(idx,1,history.length,colCount).setValues(history);
    var  confSheet = SpreadsheetApp.getActive().getSheetByName(CONFIG_SHEET_NAME);
    confSheet.getRange(LAST_UPDATED_CELL).setValue(new Date());
    confSheet.getRange(RECORD_COUNT_CELL).setValue(history.length);
  }
}


function whoop_get_heart_rate(start_date, end_date, interval){
  if(interval==null) interval=60;
  // You can go for every 60 seconds, every 600 seconds (10 min) or every 6 seconds.
  // Don't get too much data - this is a slow request - so keep the range very tight (1-2 days max!)
  
  if(start_date == null) {
    start_date=new Date();
    start_date.setDate(start_date.getDate()-1);
  }
  if(end_date==null) end_date=new Date();
    
  var config=getConfigDetails();
  var  whoopSheet = SpreadsheetApp.getActive().getSheetByName(HEART_RATE_SHEET_NAME);
  var data=whoop_get_time_series("metrics/heart_rate",start_date, end_date, {sort:'t', step:interval});
  // Data comes back as a values subarray for metrics calls;
  data=data.values;
  var rows=[];
  rows[0]=["Timestamp", "Date", "Time", "Rate"];
  console.log("Found "+data.length+"rows");
  data.forEach(row => {
               var d=new Date(row.time);
               var rowArr=[
                 row.time,
                 Utilities.formatDate(d, config.whoop.timezone, "yyyy-MM-dd"),
                 Utilities.formatDate(d, config.whoop.timezone, "HH:mm:ss"),
                 row.data
               ];
               rows[rows.length]=rowArr;
               });

  whoopSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  whoopSheet.autoResizeRows(1,rows.length);
  whoopSheet.autoResizeColumns(1,rows[0].length);
  return rows.length;
}


function whoop_get_time_series(series_name, start_date, end_date, params){
  whoop_refresh_token_if_needed();
  var config=getConfigDetails();
  var whoop=config.whoop;
  var timeZone = whoop.timezone;
  var start=Utilities.formatDate(start_date, timeZone, DATETIME_FORMAT_START);
  var end=Utilities.formatDate(end_date, timeZone, DATETIME_FORMAT_FULL);
  var url=whoop.http_base + "/users/"+ whoop.id +"/"+series_name+"?end="+end+"&start="+start;
  if(params) url+="&"+Object.keys(params).map(key => key + '=' + params[key]).join('&');
  console.log("URL:"+url);
  var json = UrlFetchApp.fetch(url,whoop.http_options).getContentText();
  var data = JSON.parse(json);
  return data;
}

// this is optimized for google data studio ordered lists of day names
function dayOfWeek(dateString){
  var config=getConfigDetails();
  var dt=new Date(dateString);
  var weekday=Utilities.formatDate(dt,config.whoop.timezone,'EEE');
  return dt.getDay()+ " "+weekday;
}


function whoop_get_history(start_date, end_date){
  var config=getConfigDetails();
  var data=whoop_get_time_series("cycles", start_date, end_date);
  var rows=[];
  rows[0]=["Date","Strain","Recovery","Sleep Score","Sleep Duration","Workouts","HRV","RHR","Average HR","Max HR", "KJ", "Comment", "Respiratory Rate", "HRV (ms)",	"Sleep (hr)",	"Recovery Type", "Day of Week","Calories","Sleep Start","Sleep End", "Bedtime", "Wakeup"];
  console.log(JSON.stringify(data[data.length-2]));
  data.forEach(row => {
  
               var rowArr=[
               row.days[0],
               (row.strain)?row.strain.score:null,
               (row.recovery)?row.recovery.score:null,
               (row.sleep)?row.sleep.score:null,
               (row.sleep)?row.sleep.qualityDuration:null,
               row.workouts?row.workouts.length:0,
               (row.recovery)?row.recovery.heartRateVariabilityRmssd:null,
               (row.recovery)?row.recovery.restingHeartRate:null,
               (row.strain)?row.strain.averageHeartRate:null,
               (row.strain)?row.strain.maxHeartRate:null,
               (row.strain)?row.strain.kilojoules:null,
               (row.during.upper == null)?"IN PROGRESS":null,
               (row.sleep && row.sleep.sleeps && row.sleep.sleeps.length>0 && row.sleep.sleeps[0].respiratoryRate)? row.sleep.sleeps[0].respiratoryRate:null,
               (row.recovery && row.recovery.heartRateVariabilityRmssd)?row.recovery.heartRateVariabilityRmssd * 1000:null,
               (row.sleep && row.sleep.qualityDuration)? row.sleep.qualityDuration / (1000*60*60) : null,
               (row.recovery && row.recovery.score )? (row.recovery.score>=67 ? "Green": (row.recovery.score>=34? "Yellow" : "Red")) :null,           
               '=WEEKDAY(INDIRECT("A"&ROW()),'+config.whoop.weekday_type+')& " " &TEXT(INDIRECT("A"&ROW()),"ddd")',
               (row.strain && row.strain.kilojoules)?row.strain.kilojoules/4.184:null,
               (row.sleep && row.sleep.sleeps && row.sleep.sleeps.length>0 && row.sleep.sleeps[0].during)?new Date(row.sleep.sleeps[0].during.lower).getTime():null,
               (row.sleep && row.sleep.sleeps && row.sleep.sleeps.length>0 && row.sleep.sleeps[0].during)?new Date(row.sleep.sleeps[0].during.upper).getTime():null,
               (row.sleep && row.sleep.sleeps && row.sleep.sleeps.length>0 && row.sleep.sleeps[0].during)?Utilities.formatDate( new Date(row.sleep.sleeps[0].during.lower), config.whoop.timezone, "HH:mm"):null,
               (row.sleep && row.sleep.sleeps && row.sleep.sleeps.length>0 && row.sleep.sleeps[0].during)?Utilities.formatDate( new Date(row.sleep.sleeps[0].during.upper), config.whoop.timezone, "HH:mm"):null,
                   ];
               
                    rows[rows.length]=rowArr;
               });

   console.log(rows);

   return rows;
}


function whoop_login(username, password){
  var auth={
    "username": username,
    "password": password,
    "grant_type": "password",
    "issueRefresh": true
  };
  var config=getConfigDetails();
  var whoop=config.whoop;
  var url=whoop.http_base+"/oauth/token";
  var response=UrlFetchApp.fetch(url,{'method':'POST','contentType': 'application/json', 'payload':JSON.stringify(auth)});                              
  var json = response.getContentText();
  var data = JSON.parse(json);
  console.log("response:" +json);
  return data;
}

function whoop_refresh_token(){
  var config=getConfigDetails();
  var whoop=config.whoop;
   var auth={
    "refresh_token": whoop.refresh_token,
    "grant_type": "refresh_token"
   };
  console.log("Refreshing Using Token: "+whoop.refresh_token );
  var url=whoop.http_base+"/oauth/token";
  var response=UrlFetchApp.fetch(url,{'method':'POST','contentType': 'application/json', 'payload':JSON.stringify(auth)});                              
  var json = response.getContentText();
  var data = JSON.parse(json);
  console.log("response:" +json);
  return data;
 }
