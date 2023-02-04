// author: Edoardo Palazzi

function updateLaunches(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var lastRow = ss.getLastRow();

  var columnJ = ss.getSheetByName("Launches (data)").getRange("J2:J").getValues();
  var lastRowJ = columnJ.filter(String).length;
  console.log(lastRowJ);

  var allTaskIds = ss.getSheetByName("Launches (data)").getRange(2, 11, lastRowJ, 1).getValues();
  var taskIds = [].concat(...allTaskIds);
  //console.log(taskIds);

  var columnA = ss.getSheetByName("Launches").getRange("A2:A").getValues();
  var lastRowA = columnA.filter(String).length;
  if (lastRowA != 1){
    lastRowA = lastRowA-1;
  }
  console.log(lastRowA);

  var allAddedTaskIds = ss.getSheetByName("Launches").getRange(3, 11, lastRowA, 1).getValues();
  var addedTaskIds = [].concat(...allAddedTaskIds);
  console.log(addedTaskIds);
  //console.log(taskIds.length);

  for (let i=0; i<taskIds.length; i++){
    //console.log(i);
    //if (addedTaskIds.includes(taskIds[i]) == FALSE)
    if(addedTaskIds.indexOf(taskIds[i], 0) == -1){
      var newEntry = ss.getSheetByName("Launches (data)").getRange(i+2, 1, 1, 34).getValues();
      //var lastRow = ss.getLastRow();

      var columnA = ss.getSheetByName("Launches").getRange("A1:A").getValues();
      var lastRow = columnA.filter(String).length;
      console.log(lastRow);

      var rangeToPaste = ss.getSheetByName("Launches").getRange(lastRow+1, 1, 1, 34);
      rangeToPaste.setValues(newEntry);
    }
  }


  // Update the status of the tasks that have already been added (either Blocked or Completed)
  var masterTracker = ss.getSheetByName("Master Tracker");
  var columnAs = masterTracker.getRange("A5:A").getValues();
  var lastRowAs = columnAs.filter(String).length;
  var allWSs = masterTracker.getRange(5, 11, lastRowAs, 1).getValues();
  var allWSsIDs = [].concat(...allWSs);
  //console.log(allWSsIDs);

  var columnAl = ss.getSheetByName("Launches").getRange("A2:A").getValues();
  var lastRowAl = columnAl.filter(String).length;
  //console.log(lastRowAl);
  if (lastRowAl != 1){
    lastRowAl = lastRowAl-1;
  }

  var allAddedTaskIds = ss.getSheetByName("Launches").getRange(3, 11, lastRowAl, 1).getValues();
  var addedTaskIds = [].concat(...allAddedTaskIds);
  //console.log(addedTaskIds);

  for (let i=0; i<addedTaskIds.length; i++){
    if(allWSsIDs.indexOf(addedTaskIds[i], 0) != -1){
      var rowNum = allWSsIDs.indexOf(addedTaskIds[i], 0)+5;
      var newStatus = masterTracker.getRange(rowNum, 25).getValue();
      var keImpact = masterTracker.getRange(rowNum, 16).getValue();
      var sxsImpact = masterTracker.getRange(rowNum, 15).getValue();
      var newStartDate = masterTracker.getRange(rowNum, 2).getValue();

      var cellToPaste = ss.getSheetByName("Launches").getRange(i+3, 25);
      cellToPaste.setValue(newStatus);
      var cellToPasteKE = ss.getSheetByName("Launches").getRange(i+3, 16);
      cellToPasteKE.setValue(keImpact);
      var cellToPasteSXS = ss.getSheetByName("Launches").getRange(i+3, 15);
      cellToPasteSXS.setValue(sxsImpact);
      var cellToPasteStartDate = ss.getSheetByName("Launches").getRange(i+3, 2);
      cellToPasteStartDate.setValue(newStartDate);
    }
  }

  var todayDate = Date();
  ss.getSheetByName("Launches").getRange(1, 1).setNote("Last Refreshed: " + todayDate);


}
