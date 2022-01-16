function getColor(recovery) {
  // 0-33
  var red_recovery=0xFF0D0D;
  // 34-65
  var yellow_recovery=0xFAB733;
  // 66-100
  var green_recovery=0x69B34C;

 /* if(recovery<34){
    return (recovery/34)
  }*/
  
  var lower_bound=red_recovery;
  var upper_bound=yellow_recovery;

  if(recovery >= 67) {
      upper_bound=green_recovery;
      lower_bound=yellow_recovery;
  }

  var percentage=((recovery % 34) / 33);
  Logger.log(percentage+" % progress ");
  var val=pickHex(rgb(upper_bound), rgb(lower_bound), percentage);
  var colorHex=(val.map(rgbToString).join(""));
  Logger.log("Recovery "+recovery +" is #"+colorHex);
  return colorHex;
}

function rgb(color){
  var RED=0xFF0000;
  var GREEN=0x00FF00;
  var BLUE=0x0000FF;
  return [
    (Math.abs((color & RED ) / GREEN)),
        (Math.abs((color & GREEN ) / BLUE)),
             (Math.abs(color & BLUE ))
  ];
}
function pickHex(color1, color2, weight) {
    var w1 = weight;
    var w2 = 1 - w1;
    var rgb = [Math.round(color1[0] * w1 + color2[0] * w2).toString(16),
        Math.round(color1[1] * w1 + color2[1] * w2).toString(16),
        Math.round(color1[2] * w1 + color2[2] * w2).toString(16)];
    return rgb;
}
function rgbToString( val ){
  var str=val.toString(16);
  if(str.length<2) return "0"+str;
  else return str;
}
function testInterpolate(){
  var recoveries=Array.from(Array(100).keys());
  Logger.log(recoveries.map(getColor).map(a=>{return "'#"+a+"'"}));
}