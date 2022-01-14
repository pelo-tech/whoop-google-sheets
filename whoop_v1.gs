function whoop_get_v1_cycles(start_date, end_date, params){
  if(params==null) params={};
  if(params['apiVersion']==null) params['apiVersion']=7;
  Logger.log("Get Cycles");
  whoop_refresh_token_if_needed();
  var config=getConfigDetails();
  var whoop=config.whoop;
  var timeZone = whoop.timezone;
  var start=Utilities.formatDate(start_date, timeZone, DATETIME_FORMAT_START);
  var end=Utilities.formatDate(end_date, timeZone, DATETIME_FORMAT_FULL);
  params.endTime=end;
  params.startTime=start;
  params.limit=50;

  var url=whoop.http_base + "/activities-service/v1/cycles/aggregate/range/"+ whoop.id;
  if(params) url+="?"+Object.keys(params).map(key => key + '=' + params[key]).join('&');
  Logger.log("URL:"+url);
  var json = UrlFetchApp.fetch(url,whoop.http_options).getContentText();
  var data = JSON.parse(json);
  return data;
}

function test_get_v1_cycles_last_5_days(){
    var config=getConfigDetails();
  var rows=[V1_COLUMN_HEADERS];
  var start=new Date(new Date("2020-02-03").getTime()- 30*24*60*60*1000);
  var end=new Date("2020-02-03");
  var cycles=whoop_get_v1_cycles(start,end);
  cycles.records.forEach(cycle=>{
    var recovery=cycle.recovery;
    var row=v1_result_to_row(cycle,config);
    rows.push(row);
  });
  SpreadsheetApp.getActive().getSheetByName("WhoopV1").getRange(1,1,rows.length,V1_COLUMN_HEADERS.length).setValues(rows);

}

function get_cycle_day(cycle){
  return cycle.days.split('\'')[1];
}
function v1_consolidate_sleeps(sleepArr){
  if(sleepArr.length==0) return {total_sleeps:0,total_naps:0};
  else {
    var sleeps=sleepArr.filter(s=>s.is_nap==false);
    var naps=sleepArr.filter(s=>s.is_nap==true);
    var sleep= sleeps[0];
    sleep.total_sleeps=sleeps.length;
    sleep.total_naps=naps.length;
    return sleep;
  }
}
var V1_COLUMN_HEADERS=["Date","Day of Week","Workouts","Cycle Start", "Scaled Strain", "Sleep Need", "Sleep Need Hrs","Timezone Offset", "Intensity Score", "Data State", "Day Strain", "KJ", "Calories", "Avg HR","Max HR", "Recovery Date", "Recovery", "Color", "RHR", "HRV", "Recov State", "Prob Covid", "HR Baseline", "SPO2","Skin Temp C", "Skin Temp F", "RHR Component", "Recovery Algo Version", "Recovery History Size", "Recovery Rate" , "Recov Normal", "Recov From SWS", "Total Sleeps", "Total Naps", "Sleep Cycles", "Disturbances", "Respiratory Rate", "Sleep Duration", "Sleep Debt Pre", "Sleep Debt Post", "Sleep Debt Delta Mins", "Sleep Score", "Sleep Latency", "Light Sleep", "SWS Sleep", "REM Sleep", "Awake", "No Sleep Data","Time in Bed", "Sleep Normal", "Calibrating" ];

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
                 sleep.sleep_need || 0 / (60*1000*60),
                 cycle.timezone_offset,
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
                sleep.quality_duration? sleep.quality_duration/ (1000*60*60):null,
                sleep.debt_pre||0,
                sleep.debt_post||0,
                Math.round((Math.max(sleep.debt_post||1,1)-Math.max(sleep.debt_pre||1,1))/(60*1000)),
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

    // console.log(rowArr);
  return rowArr;
}
function test_get_v0_cycles_last_5_days(){
  var start=new Date(new Date().getTime()- 24*100*60*60*1000);
  var end=new Date();
  var cycles=whoop_get_history(start,end);
  Logger.log(JSON.stringify(cycles));
}