function whoop_setup(){
  var ui = SpreadsheetApp.getUi();
  var credentials=promptForCredentials("Enter Whoop E-mail","Enter Whoop Password (won't be stored)");
  if (credentials == null){
    console.log("Login Aborted");
    ui.alert("Whoop authentication cancelled.");
    return;
  }
  
  var now=new Date().getTime();
  var data=whoop_login(credentials.username,credentials.password);
  processWhoopToken(data);
  ui.alert("Your whoop tokens have been obtained and are stored on your CONFIG tab.");
}

function processWhoopToken(data){
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
      return (typeof d!='object')?d:Utilities.formatDate(new Date(d),config.whoop.timezone,'yyyy-MM-dd');
                       });
  
  var startDate=whoopSheet.getRange("A"+dateColumn.length).getValue();
  startDate=new Date(startDate);
  var endDate=new Date(); // now
  var history=whoop_get_history(startDate,endDate);
  console.log("Last Value is "+startDate);
  var colCount=history[0].length; // even if its headers
  history.shift(); // remove headers
  console.log("History Size: "+history.length-1);
  if (history.length>0){
    var firstDate=history[0][0];
    console.log("First Date: "+firstDate);
    var idx=dateStrings.indexOf(firstDate);
    console.log("Found at index:"+idx);
    whoopSheet.getRange(idx,1,history.length,colCount).setValues(history);
    var  confSheet = SpreadsheetApp.getActive().getSheetByName(CONFIG_SHEET_NAME);
    confSheet.getRange(LAST_UPDATED_CELL).setValue(new Date());
    confSheet.getRange(RECORD_COUNT_CELL).setValue(history.length);
  }
}



function whoop_get_history(start_date, end_date){
  whoop_refresh_token_if_needed();
  var  whoopSheet = SpreadsheetApp.getActive().getSheetByName(WHOOP_SHEET_NAME);
  var config=getConfigDetails();

  var whoop=config.whoop;
  var timeZone = whoop.timezone;
  var start=Utilities.formatDate(start_date, timeZone, "yyyy-MM-dd'T'00:00:00.SSS'Z'");
  var end=Utilities.formatDate(end_date, timeZone, "yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
  var url=whoop.http_base + "/users/"+ whoop.id +"/cycles?end="+end+"&start="+start;
  console.log("URL:"+url);
  var json = UrlFetchApp.fetch(url,whoop.http_options).getContentText();
   var data = JSON.parse(json);
   var rows=[];
  rows[0]=["Date","Strain","Recovery","Sleep Score","Sleep Duration","Workouts","HRV","RHR","Average HR","Max HR", "KJ", "Comment"];
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
                 (row.during.upper == null)?"IN PROGRESS":null
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
