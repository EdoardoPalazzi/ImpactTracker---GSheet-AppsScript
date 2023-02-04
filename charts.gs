// author: Edoardo Palazzi

function onEdit(e) {
  if(e.range.getA1Notation() == "E1"){
      updateTitles();
  }
  if(e.range.getA1Notation() == "B1"){
      updateTitles();
  }

}


function updateTitles() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var verticalTitle = sheet.getRange(1, 5).getValue();
  var quarterTitle = sheet.getRange(1, 2).getValue();
  if (verticalTitle == "All"){
    verticalTitle = "All Verticals"
  }
  if (quarterTitle == "All"){
    quarterTitle = "All Quarters"
  }
  var newtitle = verticalTitle.concat(" ").concat(quarterTitle);

  var charts = sheet.getCharts()[0];
  var chart = charts.modify()
    .setOption('subtitle', newtitle)
    .build()
  sheet.updateChart(chart)

  var scnCharts = sheet.getCharts()[1];
  var chart2 = scnCharts.modify()
    .setOption('subtitle', newtitle)
    .setOption('applyGroupingData', 0)
    .build()
  sheet.updateChart(chart2)

  var trdCharts = sheet.getCharts()[2];
  var chart3 = trdCharts.modify()
    .setOption('subtitle', newtitle)
    .build()
  sheet.updateChart(chart3)

  var fthCharts = sheet.getCharts()[3];
  var chart4 = fthCharts.modify()
    .setOption('subtitle', newtitle)
    .build()
  sheet.updateChart(chart4)

  var fithCharts = sheet.getCharts()[4];
  var chart5 = fithCharts.modify()
    .setOption('subtitle', newtitle)
    .build()
  sheet.updateChart(chart5)

  var sthCharts = sheet.getCharts()[5];
  var chart6 = sthCharts.modify()
    .setOption('subtitle', newtitle)
    .build()
  sheet.updateChart(chart6)

  var sethCharts = sheet.getCharts()[6];
  var chart7 = sethCharts.modify()
    .setOption('subtitle', newtitle)
    .build()
  sheet.updateChart(chart7)

  var ethCharts = sheet.getCharts()[7];
  var chart8 = ethCharts.modify()
    .setOption('subtitle', newtitle)
    .setOption('applyAggregateData',0)
    .build()
  sheet.updateChart(chart8)

  var nthCharts = sheet.getCharts()[8];
  var chart9 = nthCharts.modify()
    .setOption('subtitle', newtitle)
    .build()
  sheet.updateChart(chart9)

  var tthCharts = sheet.getCharts()[9];
  var chart10 = tthCharts.modify()
    .setOption('subtitle', newtitle)
    .setOption('applyAggregateData',0)
    .build()
  sheet.updateChart(chart10)

  var elthCharts = sheet.getCharts()[10];
  var chart11 = elthCharts.modify()
    .setOption('subtitle', newtitle)
    .build()
  sheet.updateChart(chart11)

  var twthCharts = sheet.getCharts()[11];
  var chart12 = twthCharts.modify()
    .setOption('subtitle', newtitle)
    .build()
  sheet.updateChart(chart12)

  var thCharts = sheet.getCharts()[12];
  var chart13 = thCharts.modify()
    .setOption('subtitle', newtitle)
    .build()
  sheet.updateChart(chart13)

  var fttthCharts = sheet.getCharts()[14];
  var chart15 = fttthCharts.modify()
    .setOption('subtitle', newtitle)
    .build()
  sheet.updateChart(chart15)

  var tenFiveCharts = sheet.getCharts()[15];
  var chart16 = tenFiveCharts.modify()
    .setOption('subtitle', newtitle)
    .build()
  sheet.updateChart(chart16)

}
