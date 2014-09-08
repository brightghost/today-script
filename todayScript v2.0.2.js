/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
//function readRows() {
//  var sheet = SpreadsheetApp.getActiveSheet();
//  var rows = sheet.getDataRange();
//  var numRows = rows.getNumRows();
//  var values = rows.getValues();
//
//  for (var i = 0; i <= numRows - 1; i++) {
//    var row = values[i];
//    Logger.log(row);
//  }
//};

var ss = SpreadsheetApp.getActiveSpreadsheet();
var activeSheet = SpreadsheetApp.getActiveSheet();
var activeRange = activeSheet.getActiveRange();

var dateFormat = "DDDD MM/dd";
var skipWknds = true;

function dupeSelection(source, dest) {
  source.copyTo(dest);

}

function writeDate(destRange, dateObj, formatString) {
  destCell = destRange.getCell(1, 1);
  destCell.setNumberFormat(dateFormat);
  destCell.setValue(dateObj);
}

function buildCal(startDate, endDate) {
  var proto = activeRange;
  var protoHeight= activeRange.getHeight();
  var destSheet = activeSheet;
  var destRange = destSheet.getRange(1, 1, 1, 1);

  var oneDay = (1000 * 60 * 60 * 24);
  startDate.setHours(0, 0, 0, 0);
  endDate.setHours(0, 0, 0, 0);
  var startMillis = startDate.getTime();
  var endMillis = endDate.getTime();
  Logger.log("startDate.getTime:"); Logger.log(startDate.getTime());
  Logger.log('starting loop');
  Logger.log('startMillis: '); Logger.log(startMillis);
  while (startDate.getTime() <= endDate.getTime())
  {
    if ((startDate.getDay() === 0 || startDate.getDay() === 6) && skipWknds === true) { //skip weekends
      
      startDate.setTime(startDate.getTime() + oneDay);
      continue;
      
    }
    else {
      var monthName = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "MMMM");
   

      if ( monthName != activeSheet.getName() ) { //check if we need to switch to a new sheet
        if (ss.getSheetByName(monthName) !== null) {
          var nextSheet = (ss.getSheetByName(monthName));
          ss.setActiveSheet(nextSheet);
          destRange = nextSheet.getRange(1, 1, 1, 1);
          ss.setActiveRange(destRange);
        }
        else { // make a new sheet for the new month
          var newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(monthName);
          destRange = newSheet.getRange(1, 1, 1, 1);
          ss.setActiveRange(destRange);
          Logger.log("couldnt find a sheet so we're making one.");
          Logger.log(e);
        }
      }


      // if ( monthName != activeSheet.getName() ) { //make a new sheet if we're in a new month
      //   try {
      //     destRange = ss.getSheetByName(monthName).getRange(1, 1, 1, 1);
      //     ss.setActiveRange(destRange);
      //   }
      //   catch (e) {
      //     destRange = SpreadsheetApp.getActiveSpreadsheet().insertSheet(monthName).getRange(1, 1, 1, 1);
      //     ss.setActiveRange(destRange);
      //     Logger.log("couldnt find a sheet so we're making one.");
      //     Logger.log(e);
      //   }
      // }



      dupeSelection(proto, destRange);
      writeDate(destRange, startDate, dateFormat);
       
      //dest = dest + protoHeight;
      var newRow = (destRange.getRow() + protoHeight);
      var newCol = destRange.getColumn();
      Logger.log(' newRow: '); Logger.log(newRow); Logger.log('newCol: '); Logger.log(newCol);
      destRange = SpreadsheetApp.getActiveSheet().getRange(newRow, newCol);
      // startMillis += oneDay;
      startDate.setTime(startDate.getTime() + oneDay);
      Logger.log('executed loop, startDate is now '); Logger.log(startDate);
      Logger.log('destRange: '); Logger.log(destRange);
      
    }
  }
    
}

function test(){
  var startDate = new Date(2014, 8, 2);
  var endDate = new Date(2014, 10, 6);
  Logger.log(startDate); Logger.log(endDate);
  buildCal(startDate, endDate);
}

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Test",
    functionName : "test"
  }];
  spreadsheet.addMenu("Script Center Menu", entries);
}
