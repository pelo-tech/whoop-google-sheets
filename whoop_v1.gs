var V1_COLUMN_HEADERS=["Date","Day of Week","Workouts","Cycle Start", "Scaled Strain", "Sleep Need", "Sleep Need Hrs","Timezone Offset", "Intensity Score", "Data State", "Day Strain", "KJ", "Calories", "Avg HR","Max HR", "Recovery Date", "Recovery", "Color", "RHR", "HRV", "Recov State", "Prob Covid", "HR Baseline", "SPO2","Skin Temp C", "Skin Temp F", "RHR Component", "Recovery Algo Version", "Recovery History Size", "Recovery Rate" , "Recov Normal", "Recov From SWS", "Total Sleeps", "Total Naps", "Sleep Cycles", "Disturbances", "Respiratory Rate", "Sleep Duration", "Sleep Debt Pre", "Sleep Debt Post", "Sleep Debt Delta Mins", "Sleep Score", "Sleep Latency", "Light Sleep", "SWS Sleep", "REM Sleep", "Awake", "No Sleep Data","Time in Bed", "Sleep Normal", "Calibrating" ];
var MINUTE_IN_MS=60*1000;
var HOUR_IN_MS=60*MINUTE_IN_MS;
var DAY_IN_MS=24*HOUR_IN_MS;

function whoop_get_v1_cycles_internal(params, whoop){
  // This is called by whoop_get_v1_cycles as many times as date range requires. 
  // Assumption here is that params are all prepped
  // startDate, endDate, apiVersion, limit have all been validated and populated
  // Assume whoop token has been refreshed if needed
  Logger.log("Get Cycles - full params ("+JSON.stringify(params));

  var url=whoop.http_base + "/activities-service/v1/cycles/aggregate/range/"+ whoop.id;
  if(params) url+="?"+Object.keys(params).map(key => key + '=' + params[key]).join('&');
  Logger.log("URL:"+url);
  var json = UrlFetchApp.fetch(url,whoop.http_options).getContentText();
  var data = JSON.parse(json);
  return data;
}

function v1_get_whoop_cycles(start_date, end_date, params){
  if(start_date.getTime()>end_date.getTime()) throw "Start date must be before End Date!";
  params = params || {};
  if(params['apiVersion']==null) params['apiVersion']=7;
  Logger.log("V1: Get Cycles ("+start_date+","+end_date+","+JSON.stringify(params));
  whoop_refresh_token_if_needed();
  var config=getConfigDetails();
  var whoop=config.whoop;
  var timeZone = whoop.timezone;
  var start=Utilities.formatDate(start_date, timeZone, DATETIME_FORMAT_START);
  var end=Utilities.formatDate(end_date, timeZone, DATETIME_FORMAT_FULL);
  params.endTime=end;
  params.startTime=start;
  params.limit=Math.min(params.limit || 50,50);
  if(typeof params["offset"]==undefined) params.offset=0;
  // All parameters now validated and formatted.
  var total_days=Math.ceil(Math.abs(end_date.getTime()-start_date.getTime()) / DAY_IN_MS);
  Logger.log("Total number of days to retrieve: "+total_days);
  var allResults=[];
  var currResults=null;
  var offset=0;
  while (offset <= total_days){
    Logger.log("Looking for up to "+params.limit+" results from offset " +offset)
    params.offset=offset;
    currResults=whoop_get_v1_cycles_internal(params, whoop );
    Logger.log("Total: "+currResults.total_count+" Offset:"+currResults.offset+" Arr Size: "+(currResults.records&&currResults.records.length)); 
    if(currResults.records ) {
      var records=currResults.records.map(a=>v1_result_to_row(a,config));
      allResults.push(... records);
    }
    offset+=params.limit;
  }
  Logger.log("Total results :" +allResults.length);
  return allResults;
} 



function get_cycle_day(cycle){
  return new Date(cycle.days.split('\'')[1]);
}
function v1_consolidate_sleeps(sleepArr){
  if(sleepArr.length==0) return {total_sleeps:0,total_naps:0};
  else {
    var sleeps=sleepArr.filter(s=>s.is_nap==false);
    var naps=sleepArr.filter(s=>s.is_nap==true);
    var sleep=sleeps.length?sleeps[0]:{};
    sleep.total_sleeps=sleeps.length;
    sleep.total_naps=naps.length;
    return sleep;
  }
}

function tzoffset_to_number(str){
  if(typeof str==Number) return str;
  else if((str||"").indexOf("+")>-1) str=parseInt(str.substring(1).replace(":",""));
  return str;
}
function v1_result_to_row(result, config){
  var sleep=v1_consolidate_sleeps(result.sleeps);
  var recovery=result.recovery || {};
  var cycle=result.cycle || {};
  var workouts=result.workouts || [];

               var rowArr=[
               get_cycle_day(cycle),
               '=WEEKDAY(INDIRECT("A"&ROW()),'+config.whoop.weekday_type+')& " " &TEXT(INDIRECT("A"&ROW()),"ddd")',

                 // workouts
                 workouts.length,
                 // cycle
                 cycle.created_at,
                 cycle.scaled_strain,
                 sleep.sleep_need,
                 (sleep.sleep_need || 0) / (HOUR_IN_MS),
                 tzoffset_to_number(cycle.timezone_offset),
                 cycle.intensity_score,
                 cycle.data_state,
                 cycle.day_strain,
                 cycle.day_kilojoules,
                 (cycle.day_kilojoules)?cycle.day_kilojoules/4.184:null,
                 cycle.day_avg_heart_rate,
                 cycle.day_max_heart_rate,
                 // recovery
                 recovery.date,
                 recovery.recovery_score,
                 (recovery.recovery_score )? (recovery.recovery_score>=67 ? "Green": (recovery.recovery_score>=34? "Yellow" : "Red")) :null,           
                 recovery.resting_heart_rate,
                 recovery.hrv_rmssd?  recovery.hrv_rmssd*1000 : null,
                 recovery.state,
                 recovery.prob_covid,
                 recovery.hr_baseline,
                 recovery.spo2,
                 recovery.skin_temp_celsius||0,
                 (recovery.skin_temp_celsius||0)* (9/5) + 32,
                 recovery.rhr_component,
                 recovery.algo_version,
                 recovery.history_size,
                 recovery.recovery_rate,
                 recovery.is_normal,
                 recovery.from_sws,
                // sleeps
                sleep.total_sleeps,
                sleep.total_naps,
                sleep.cycles_count||0,
                sleep.disturbances||0,
                sleep.respiratory_rate||0,
                sleep.quality_duration? sleep.quality_duration/ (HOUR_IN_MS):null,
                sleep.debt_pre||0,
                sleep.debt_post||0,
                Math.round((Math.max(sleep.debt_post||1,1)-Math.max(sleep.debt_pre||1,1))/(MINUTE_IN_MS)),
                sleep.score||0,
                sleep.latency||0,
                sleep.light_sleep_duration||0,
                sleep.slow_wave_sleep_duration||0,
                sleep.rem_sleep_duration||0,
                sleep.wake_duration||0,
                sleep.no_data_duration||0,
                sleep.time_in_bed||0,
                sleep["is_normal"],
                recovery.calibrating
                   ];
  return rowArr;
}
 

 function test_get_v1_cycles_last_5_days(){
  var total_days=50;
  var rows=[V1_COLUMN_HEADERS];
  var start=new Date("2020-01-01");//new Date("2020-01-01");.getTime()- total_days*24*60*60*1000);
  var end=new Date();//"2020-02-03");
  var cycles=whoop_get_v1_cycles(start,end, {offset:0, limit:12});
  // First column is date string, so lets ensure we are sorted
  cycles=cycles.sort((a,b)=>{  return  a[0]-b[0];});
  rows.push(...cycles);
  SpreadsheetApp.getActive().getSheetByName(V1_WHOOP_SHEET_NAME).getRange(1,1,rows.length,V1_COLUMN_HEADERS.length).setValues(rows);
}

function v1_rebuild_history(){
  var config=getConfigDetails();
  var rows=[V1_COLUMN_HEADERS];
  var endDate=new Date();
  Logger.log(config.whoop.history_size);
  var startDate=new Date(new Date().getTime()-(config.whoop.history_size * DAY_IN_MS));
  Logger.log("START :"+startDate);
  var cycles=v1_get_whoop_cycles(startDate,endDate,{offset:0, limit:50});
  Logger.log("V1: Getting history from Start Date (Configured as max "+config.whoop.history_size +" days back) "+localDateString(startDate)+" to end date "+ localDateString(endDate));
  var sheet=SpreadsheetApp.getActive().getSheetByName(V1_WHOOP_SHEET_NAME);
  sheet.clearContents();
  rows.push(...cycles.sort((a,b)=>{  return  a[0]-b[0];}));
  sheet.getRange(1,1,rows.length,V1_COLUMN_HEADERS.length).setValues(rows);
}

function v1_get_incremental_history(){
  var config=getConfigDetails();
  var whoopSheet = SpreadsheetApp.getActive().getSheetByName(V1_WHOOP_SHEET_NAME);
  var dateColumn=whoopSheet.getRange("A1:A").getValues().filter(String);
  if(dateColumn.length<2){
    Logger.log("V1: Rebuilding history. No valid data in place");
    v1_rebuild_history();
    return;
  } 
  
  // Annlying Date Formatting requires us to stringify values
  var dateStrings=dateColumn.map(
    function(d,idx,arr){
      return (typeof d!='object')?d:localDateString(d, config);
                       });
  
  // last valid date in the spreadsheet
  var startDate=whoopSheet.getRange("A"+dateColumn.length).getValue();
  Logger.log("Start Date "+startDate);
  startDate=new Date(startDate);
  var endDate=new Date(); // now

  Logger.log("V1: Getting history from Start Date "+localDateString(startDate, config)+" to end date "+ localDateString(endDate, config));
  var cycles=v1_get_whoop_cycles(startDate,endDate, {offset:0});
  var rows=cycles.sort((a,b)=>{  return  a[0]-b[0];});
  var dates=whoopSheet.getRange("A1:A").getValues();

  Logger.log("Last Value is "+localDateString(startDate, config));
  var colCount=cycles[0].length; // even if its headers
  Logger.log("Cycles Found: "+cycles.length);
  if (cycles.length>0){
    var firstDate=cycles[0][0];
    Logger.log("First Date: "+localDateString(firstDate,config));
    var idx=dateStrings.indexOf(localDateString(firstDate, config));
    idx+=1 ; // correct index to be a 'row number'
    Logger.log("Found at row number:"+idx);
    Logger.log("Cycles Found: "+cycles.length);
    whoopSheet.getRange(idx,1,cycles.length,colCount).setValues(cycles);
    var  confSheet = SpreadsheetApp.getActive().getSheetByName(CONFIG_SHEET_NAME);
    confSheet.getRange(LAST_UPDATED_CELL).setValue(new Date());
    confSheet.getRange(RECORD_COUNT_CELL).setValue(whoopSheet.getDataRange().getNumRows()-1);
  }
}
