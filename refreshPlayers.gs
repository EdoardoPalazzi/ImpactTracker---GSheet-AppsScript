// author: Edoardo Palazzi

function refresh(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var masterT = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Tracker");
  var dict = {};
  // Total # of WSs
  var wsNum = 0;
  for (let i=5; i<300; i++){
    if (masterT.getRange(i, 1).getValue() != ""){
      wsNum += 1
    }
  }

  // For each WS get the needed info
  for (var row = 5; row < wsNum; row++){
    var intents = masterT.getRange(row, 7).getValue();
    var intentsList = intents.split(",");
    var locale = masterT.getRange(row, 5).getValue();
    var date = masterT.getRange(row, 3).getValue();
    if (date == ""){
      date = "In Progress";
    }
    // For each intent in a WS create a key and add it to dict
    for (let i=0; i< intentsList.length; i++){
      var key = locale.concat(" - ", intentsList[i].trim());
      if (key in dict){
        dict[key].push(date);
      }
      else{
        var dateList = [date];
        dict[key] = dateList;
      }
    }
  }

  console.log(dict);
  // Count # of locales are wanted to report data
  var localeNum = 0;
  for (let i=1; i<30; i++){
    if (sheet.getRange(1, i).getValue() != ""){
      localeNum += 1
    }
  }
  // For all combinations of intents and locales concatenate and search in dict
  for (var col = 3; col < localeNum; col++){
    var locale1 = sheet.getRange(1, col).getValue();
    for (var row = 3; row < 118; row++){
      var intent1 = sheet.getRange(row, 1).getValue();
      var key1 = locale1.concat(" - ", intent1);
      if (key1 in dict){
        var date1 = dict[key1];
        //var maxDate = new Date(Math.max(...date1.map(element => {return new Date(element.date);})))  
        var maxDate = date1.reduce(function (a, b) { return a > b ? a : b; });

        sheet.getRange(row, col).setValue(maxDate);
      }
    }
  }
  var todayDate = Date();
  sheet.getRange(1, 1).setNote("Last Refreshed: " + todayDate);
}


function refreshHorizontal(){
  var sheet = SpreadsheetApp.getActiveSheet();
  var masterT = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Tracker");
  var dict = {};
  //var wsNum = masterT.getRange(5, 1, 1000, 1).getValues().length;
  var wsNum = 0;
  for (let i=5; i<300; i++){
    if (masterT.getRange(i, 1).getValue() != ""){
      wsNum += 1
    }
  }

  for (var row = 5; row < wsNum; row++){
    var intents = masterT.getRange(row, 7).getValue();
    var intentsList = intents.split(",");
    var locale = masterT.getRange(row, 5).getValue();
    var date = masterT.getRange(row, 3).getValue();
    if (date == ""){
      date = "In Progress";
    }
    for (let i=0; i< intentsList.length; i++){
      var key = locale.concat(" - ", intentsList[i].trim());
      if (key in dict){
        dict[key].push(date);
      }
      else{
        var dateList = [date];
        dict[key] = dateList;
      }
    }
  }

  console.log(dict);
  var localeNum = 0;
  for (let i=1; i<30; i++){
    if (sheet.getRange(1, i).getValue() != ""){
      localeNum += 1
    }
  }
  localeNum;
  for (var col = 5; col < localeNum; col++){
    var locale1 = sheet.getRange(1, col).getValue();
    for (var row = 3; row < 70; row++){
      var intent1 = sheet.getRange(row, 1).getValue();
      var key1 = locale1.concat(" - ", intent1);
      if (key1 in dict){
        var date1 = dict[key1];
        //var maxDate = new Date(Math.max(...date1.map(element => {return new Date(element.date);})))  
        var maxDate = date1.reduce(function (a, b) { return a > b ? a : b; });

        sheet.getRange(row, col).setValue(maxDate);
      }
    }
  }
  var todayDate = Date();
  sheet.getRange(1, 1).setNote("Last Refreshed: " + todayDate);
  //var croppedCell = sheet.getRange(1, 1).getValue().toString().replace("", todayDate);
}




/*
function refresh1() {
  var sheet = SpreadsheetApp.getActiveSheet();

  for (var col = 3; col < 7; col++){
    var locale = sheet.getRange(1, col).getValue();
    for (var row = 3; row < 118; row++){
      var vertical = sheet.getRange(row, 2).getValue();
      var intent = sheet.getRange(row, 1).getValue();
      var masterT = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Master Tracker");
      for (var row1 = 5; row < 200; row1++){
        if (masterT.getRange(row1, 1).getValue() == vertical){
          if (masterT.getRange(row1, 5).getValue() == locale){
            var intents = masterT.getRange(row1, 7).getValue();
            var intentsList = intents.split(",");
            if (intentsList.includes(intent)){
              var newDate = masterT.getRange(row1, 3).getValue();
              var currentDate = sheet.getRange(row, col).getValue();
              if (newDate == ""){
                sheet.getRange(row, col).setValue("In Progress");
                var currentDate = sheet.getRange(row, col).getValue();
              }
              
              sheet.getRange(row, col).setValue(newDate);

              
            }
          }
        }
      }
    }
  }
}
*/
