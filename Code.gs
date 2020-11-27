function getConfigDetails(){
  var cfg = SpreadsheetApp.getActive().getSheetByName(CONFIG_SHEET_NAME);
  var tz=cfg.getRange(TIME_ZONE_CELL).getValue();
 
  var whoop_access_token=cfg.getRange(WHOOP_ACCESS_TOKEN_CELL).getValue(); 
  var whoop_refresh_token=cfg.getRange(WHOOP_REFRESH_TOKEN_CELL).getValue(); 
  var whoop_id = cfg.getRange(WHOOP_ID_CELL).getValue();
  var whoop_username= cfg.getRange(WHOOP_USERNAME_CELL).getValue(); 
  var whoop_token_expiry = cfg.getRange(WHOOP_TOKEN_EXPIRY).getValue();
  var history_size = cfg.getRange(HISTORY_LOAD_SIZE_CELL).getValue();
  if(!history_size  || isNaN(history_size)) {
    history_size=60;
    cfg.getRange(HISTORY_LOAD_SIZE_CELL).setValue(history_size);
  }
  
  var weekday_type = cfg.getRange(WEEKDAY_TYPE_CELL).getValue();
  if(![1,2,3].includes(weekday_type)){
    // No weekday type specified. Let's go with 1.
    weekday_type=1;
  }  else {
    weekday_type=WEEKDAY_TYPE_CELL_REFERENCE;
  }

  return {
    "whoop":{
      "timezone" :tz,
      "http_base":WHOOP_API_BASE,
      "access_token": whoop_access_token,
      "refresh_token": whoop_refresh_token,
      "token_expiry": whoop_token_expiry,
      "username": whoop_username,
      "history_size": history_size,
      "weekday_type": weekday_type,
      "id": whoop_id,
      "http_options":
      {
        "headers":
        {
          "Authorization":"Bearer "+whoop_access_token
        }
      }
    }
  };
}
 
function promptForCredentials(username_prompt, password_prompt){
  var username=promptForText(username_prompt);
  if(username==null) return null;
  var password=promptForText(password_prompt);
  if (password==null) return null;
  return {
    username: username, 
    password: password
  };
}

function promptForText(msg) {
  var ui = SpreadsheetApp.getUi(); 
  var result = ui.prompt(
    msg+":",
      ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  var button = result.getSelectedButton();
  if(button == ui.Button.CANCEL) return null;
  var text = result.getResponseText();
  return text;
}