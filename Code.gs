var sheetToSort = "Webcast Tracker (Submit Form)"; // replace this with "Form Responses" or whichever sheet you want to sort automatically
var columnToSortBy = 7; // column A = 1, B = 2, etc.
var rangeToSort = "A2:AA";
var sheetDESPName = "DESP (Auto)";

//Variables for finding the number of rows
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetToSort);
var numOfRows = 0;
var tempRowCounter = sheet.getRange(1,2);//
var tempCounters = 1;


//variables for 
var columnToFill = 13; //column I
var columnToFillWith = 7; // column F
var columnToFill2 = 14; //column J
var columnToFill3 = 5; //column D

var isNotEmpty = true;

var emailAddress = "jsimbol@brighttalk.com";


function onEdit() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var editedCell = sheet.getActiveCell();
  if (editedCell.getColumn() == columnToSortBy && sheet.getName() == sheetToSort) {
    sortFormResponsesSheet();
  }
}

//-----------------------------------------------------------------------------

//function for calculating the number of rows in the webcast tracker
function countRowsWebcastTracker() {
  while (true) {
    if(tempRowCounter.isBlank()) break;
    numOfRows++;
    tempCounters++;
    tempRowCounter = sheet.getRange(tempCounters,2);
  }
  //Logger.log(numOfRows);
  return numOfRows;
}

//function for automatically filling out the webinar id in column A
function fillColumnA() {
  var webinarIDCol = 5;
 countRowsWebcastTracker();
  for ( var x = 2; x < numOfRows; x++ ) {
     var cell = sheet.getRange(x,1)
     var copyFrom = sheet.getRange(x,webinarIDCol);
    if (cell.isBlank()){
      cell.setValue(copyFrom.getValue());
    }
  }
}

//function for adding days to a Date Object
function addDays(date, days) {
    var result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
}


//function for filling in send dates for all webcasts
function fillSendDates() {
  
  //update numOfRows
  countRowsWebcastTracker();
  Logger.log(numOfRows);
  
  //Webcast Tracker Submit Form Reference
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetToSort);
  
  //webcast date column
  var dateColumn = 7;
  
  //create row iterator range
  var currentRow;
  
  //variable for blankDateCell reference
  var blankDateCell;
  var blankDateCell2;
  var blankDateCell3;
  
  //variable to track how far current date is from webcast date
  var switchVar;
  
  //variables to check if sends 1-3 are present
  var blankCheck;
  var blankCheck2;
  var blankCheck3;
  
  //variable to check if case 2 ( 2&3 blank, 3 blank, 2 blank )
  var case2 = false;
  
  //create Date object for row iterator
  var dateIterator;
  
  //Send Column References
  var firstSendColumn = 12;
  var secondSendColumn = 13;
  var thirdSendColumn = 14;
  var fourthSendColumn = 15;
  var fifthSendColumn = 16;
  var sixthSendColumn = 17;
  var seventhSendColumn = 18;
  var eighthSendColumn = 19;
  var ninthSendColumn = 20;
  var tenthSendColumn = 21;
  
  
  //get current date and create vars for 1 week and 2 weeks out
  var currentDate = new Date();
  var oneWeekOut = new Date();
  var twoWeeksOut = new Date();
  
  //start counter at row 2 to avoid header row
  var rowCounter = 2;

  //iterate through all known rows #A1
  for ( ; rowCounter <= numOfRows; rowCounter++ ) {
    Logger.log(rowCounter + " = RowCounter");
   
    //if blank
    blankCheck = sheet.getRange(rowCounter,firstSendColumn);
    blankCheck2 = sheet.getRange(rowCounter, secondSendColumn);
    blankCheck3 = sheet.getRange(rowCounter, thirdSendColumn);
       
    if (blankCheck.isBlank() && blankCheck2.isBlank() && blankCheck3.isBlank())  {
      //get webcast date
      currentRow = sheet.getRange(rowCounter,dateColumn);
      dateIterator = new Date(currentRow.getValue());
      
      //set cell to be filled in
      blankDateCell = sheet.getRange(rowCounter, firstSendColumn);
      blankDateCell2 = sheet.getRange(rowCounter, secondSendColumn);
      blankDateCell3 = sheet.getRange(rowCounter, thirdSendColumn);
      
      //calculate 1 and 2 weeks before markers from webcast date
      var oneWeekOut = addDays(dateIterator,-7);
      var twoWeeksOut = addDays(dateIterator,-14);
      
      //check if current date is after webcast date
      if ( currentDate >= dateIterator ) switchVar = "ONDEMAND";
      //check if current date is within one week of webcast date
      else if ( currentDate > oneWeekOut ) switchVar = "ONEWEEK";
      //check if current date is within two weeks of webcast date
      else if ( currentDate > twoWeeksOut ) switchVar = "TWOWEEKS";
      //check if current date is beyond two weeks of webcast date
      else if ( currentDate < twoWeeksOut)  switchVar = "STANDARD";
      
      //switch statement handling on demand, one week, two weeks, and standard send dates
      switch(switchVar) {
          
        //On Demand Webcast Promotions
        case "ONDEMAND":
          //ondemand code block
          //set two days from now as first send 
          var twoDaysFromNow = addDays(currentDate, 2);
          //set nine days from now as second send
          var nineDaysFromNow = addDays(currentDate, 9);
          //set sixteen days from now as third send
          var thirteenDaysFromNow = addDays(currentDate, 13);
          //fill in blank cells with proper send dates
          if ( blankCheck.isBlank() ) blankDateCell.setValue(twoDaysFromNow);
          if ( blankCheck2.isBlank() ) blankDateCell2.setValue(nineDaysFromNow);
          if ( blankCheck3.isBlank() ) blankDateCell3.setValue(thirteenDaysFromNow);
          break;
          
        //Webcast date is between today and one week from today  
        case "ONEWEEK":
          //one week code block
          //set two days from now as first send 
          var twoDaysFromNow = addDays(currentDate, 2);
          var liveDay = new Date(sheet.getRange(rowCounter, dateColumn).getValue());
          
          
          // if liveday is before two days from now
          if (liveDay < twoDaysFromNow) {
           //set first send to live day
            if ( blankCheck.isBlank() ) blankDateCell.setValue(liveDay);
           //set second send to two days after live day
            if ( blankCheck2.isBlank() ) blankDateCell2.setValue(addDays(liveDay,2));
           //set third send to 9 days after live day
            if ( blankCheck3.isBlank() ) blankDateCell3.setValue(addDays(liveDay,9));
          }
          else {
           //set first send day to two days from now
           if ( blankCheck.isBlank() ) blankDateCell.setValue(addDays(currentDate, 2));
           //set second send day as live day
           if ( blankCheck2.isBlank() ) blankDateCell2.setValue(liveDay);
           //set third send as seven days after live day
           if ( blankCheck3.isBlank() ) blankDateCell3.setValue(addDays(liveDay,7));
          }
          break;
        
        //if liveday is greater than one week days but less than 2 weeks out
        case "TWOWEEKS":
          //two weeks code block
          //set live day variable
          var liveDay = new Date(sheet.getRange(rowCounter, dateColumn).getValue());
          //set first send as 1 week before
          if ( blankCheck.isBlank() ) blankDateCell.setValue(addDays(liveDay, -7));
          //set second send as live day
          if ( blankCheck2.isBlank() ) blankDateCell2.setValue(liveDay);
          //set third send as seven days after live day
          if ( blankCheck3.isBlank() ) blankDateCell3.setValue(addDays(liveDay, 7));
          break;
          
        //Standard handling of send dates >2 weeks out  
        case "STANDARD":
          //standard code block
          //set live day variable
          var liveDay = new Date(sheet.getRange(rowCounter, dateColumn).getValue());
          //set up 2 weeks out as first send
          if ( blankCheck.isBlank() ) blankDateCell.setValue(addDays(liveDay, -14));
          //set up 1 week out as second send
          if ( blankCheck2.isBlank() ) blankDateCell2.setValue(addDays(liveDay, -7));
          //set up live day as third send
          if ( blankCheck3.isBlank() ) blankDateCell3.setValue(liveDay);
          break;
          
      }//end of switch statement
      blankCheck.setBackground("green");
      blankCheck2.setBackground("green");
      blankCheck3.setBackground("green");
    }//end of if blank code block
    
  }//end of iterating through all known rows #A1

}


function sendEmails() {
  var currentTime = new Date();
  var sheet = SpreadsheetApp.getActiveSheet();
  var editedCell = sheet.getActiveCell();
  MailApp.sendEmail("jsimbol@brighttalk.com", "subject", currentTime);
}

function dailyEmailSchedulePlanner() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetToSort);
  var sheetDESP = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetDESPName);
  var sheetRange = sheet.getRange(2,9);
  var sheetDESPRange = sheetDESP.getRange(1,2);
  var dateCounter = 0;
  var dateComparitor = sheetDESPRange.getValue();
  var dateComparitorObject = new Date(dateComparitor);
  var copySpotRow = 4;
  var copySpot = sheetDESP.getRange(copySpotRow,2);
  var tempRowCounter = sheet.getRange(1,2);//
  var tempCounters = 1;
  var rowCounter = 2;

  //count how many rows there are
  countRowsWebcastTracker();
  


  //iterate through the 14 dates starting from the start date
  for (dateCounter = 0; dateCounter < 14; dateCounter++) {
    //interating through the ten columns responsible for the ten sends
    for (var sendCounter = 1; sendCounter <= 10; sendCounter++) {

      //convert SendCounter to columnID
      var sendColumn = sendCounter + 11;
      Logger.log(sendColumn + " and also " + dateCounter);
      //iterate through rows until number of rows is reached   
      for (rowCounter = 2; rowCounter <= numOfRows; rowCounter++) {
        //get value inside temporaryCell
        var tempRange = sheet.getRange(rowCounter, sendColumn);
        //create date object for date from temporary cell
        var tempDate = new Date(tempRange.getValue());
        
        //compare date, month, and year of tempDate to comparitorDate
        var equalsDate = tempDate.getDate() == dateComparitorObject.getDate();
        var equalsMonth = tempDate.getMonth() == dateComparitorObject.getMonth();
        var equalsYear = tempDate.getYear() == dateComparitorObject.getYear();
        //compare tempDate to comparitorDate
        if (equalsDate && equalsMonth && equalsYear) {
          //get whole tempRow
          var tempRow = sheet.getRange(rowCounter, 1, 1, 26);
          Logger.log(tempRow.getValue());
          tempRow.copyTo(copySpot);
          copySpotRow++;
          copySpot = sheetDESP.getRange(copySpotRow, 2);
        
        //end of comparing tempDate to comparitorDate
        }
      //end of iterating through rows
      }
    }//end iterating through the four sends
    dateComparitorObject.setDate(dateComparitorObject.getDate()+1);
    sheetDESP.insertRowBefore(copySpotRow);
    copySpotRow++;
    sheetDESP.getRange(copySpotRow, 1).setValue(dateComparitorObject);
    copySpotRow++;
    copySpot = sheetDESP.getRange(copySpotRow, 2);

  }//end of iterating through the 14 dates
}

  
//function for copying 


function sortFormResponsesSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetToSort);
  var range = sheet.getRange(rangeToSort);
  range.sort( { column : columnToSortBy, ascending: false } );
}

