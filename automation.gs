//_________________________________________________________________________//
// SET THE NAME OF THE HEADERS HERE. THEY MUST BE THE SAME IN ALL THE TABS //
//_________________________________________________________________________//
  let headerDate = "Date";                                                 //
  let headerReleaseBugID = "Release Bug ID";                               //
  let headerQAedBy = "QA'ed by";                                           //
  let headerDevice = "Device";                                             //
  let headerExpectedResult = "Expected Result"                             // 
  let headerTimeOfExecution = "Time of execution"                          //
//_________________________________________________________________________//
  
  let ui = SpreadsheetApp.getUi();
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let ldap = Session.getActiveUser().getEmail().split("@")[0];
  let date = Utilities.formatDate(new Date(), "GMT+1", 'MM/dd/yyyy');
  let tests = ss.getSheetByName('Test Suite').getRange('D3:D').getValues();
  let tabs = ss.getSheets();
  let sheetNames = ['Test Suite'];

function newTest(shift) {
  Logger.log(ldap + " added a test for " + shift + " - " + date)
  //Create the alert to double check if the operation is intended.
  var response = ui.alert(`⚠️ Do you want to create a test for ${date} - ${shift}? ⚠️
  (This operation might take a few minutes!)`, ui.ButtonSet.OK_CANCEL);
  //If ok, resume operation.
  if ( response == ui.Button.OK){
    //Create the loop to perform verifications
    for(i=0;i<tabs.length;i++){
      //Get the names for each existing sheet.
      var name = tabs[i].getSheetName();
    
      //For each sheet, verify if its name is equal to one of the test names on the Test Suite
      for(b=0;b<tests.length;b++){
        if(tests[b]==name){
          //Verify if the sheet is hidden or not, meaning the test is performed or paused.
          if(!ss.getSheetByName(name).isSheetHidden()){
          
            //Creates the list where the new columns are going to be added.        
            sheetNames.push(name);
          }
        }
      }
    }

    //Verify if there as different ranges on the tests
    var maxColumns = []
    for(a=0;a<sheetNames.length;a++){
      maxColumns.push(ss.getSheetByName(sheetNames[a]).getMaxColumns())
    }
    //Gets the minimum range
    var shortestRange = Math.min(...maxColumns)
    //Check if there is any of the tests done on all tabs that have more than the minimum number
    for(a=0;a<sheetNames.length;a++){
      //take the sheet
      var verifySheet = ss.getSheetByName(sheetNames[a])
      //take its range
      var verifyRange = verifySheet.getMaxColumns()
      //verify if it's longer than the shortest
      if(verifyRange > shortestRange){
        //verify if there is a test on it
        var currentResult = verifySheet.getRange(3, verifyRange).getDisplayValue();
        if(currentResult == "Passed" || currentResult == "Failed" || currentResult == "Alpha not deployed" || currentResult == "Blocked"){
          verifySheet.getRange(1, verifyRange).setValue("Extra");
          verifySheet.getRange(2, verifyRange).setValue("Entry");
          //if one test is done, create the others tests as Skipped.
          //loop though all the sheets again
          for(z=0;z<sheetNames.length;z++){
          //if they have less columns then the currentResult, create the columns (verifyRange - rangeCurrentSheet(create))
            var rangeCurrentSheet = ss.getSheetByName(sheetNames[z]).getMaxColumns();
            if(rangeCurrentSheet < verifyRange){
              //mark the new columns as Skipped       
              var diffBetweenColumns = verifyRange - rangeCurrentSheet
              ss.getSheetByName(sheetNames[z]).insertColumnsAfter(rangeCurrentSheet, diffBetweenColumns)
              ss.getSheetByName(sheetNames[z]).getRange(1, rangeCurrentSheet+1,1,diffBetweenColumns).setValue("Extra")
              ss.getSheetByName(sheetNames[z]).getRange(2, rangeCurrentSheet+1,1,diffBetweenColumns).setValue("Entry")
              ss.getSheetByName(sheetNames[z]).getRange(3, rangeCurrentSheet+1,1,diffBetweenColumns).setValue("Skipped")
            }   
          }
          //Pass the verification formulas to the Test Suite
          var sh = ss.getSheetByName('Test Suite');
          var lastColumnNumber = sh.getLastColumn();
            var newColumn = sh.getRange(1, lastColumnNumber).getA1Notation().match(/[^0-9]/g).join("");
          //Loop though the lines and put the formulas.
          for(t=0;t<sh.getLastRow()-3;t++){
            var formulaSuiteVerification = `=if(${newColumn}1='${tests[t]}'!${newColumn}1, if(${newColumn}2='${tests[t]}'!${newColumn}2,'${tests[t]}'!${newColumn}3, "NoMatch"),"")`
            sh.getRange(t+3,lastColumnNumber).setValue(formulaSuiteVerification)
          }
          break
        } else {
          //if not, delete unused tabs and make everything the same.
          var tabsToDelete = verifyRange - shortestRange
          verifySheet.deleteColumns(shortestRange+1, tabsToDelete)
        }
      }
    }
    //Start the process of adding the current tests for all the tests.
    for(a=0;a<sheetNames.length;a++){
      var sheet = ss.getSheetByName(sheetNames[a]);
      var lastColumnNumber = sheet.getLastColumn();
      var width = sheet.getRange(1, lastColumnNumber).getWidth();
      sheet.insertColumnsAfter(lastColumnNumber, 1).setColumnWidths(1,1,width);
    
      //Put date
      sheet.getRange(1, lastColumnNumber+1).setValue(date);
      //Put the shift
      sheet.getRange(2,lastColumnNumber+1).setValue(shift);
      //Copy formulas and paste
      if(sheetNames[a] !== 'Test Suite'){
        //Grab new length info
        var newLastColumn = sheet.getLastColumn();
      
        //Translate do A1 notation and drop the info of rows.
        var newColumn = sheet.getRange(1, newLastColumn).getA1Notation().match(/[^0-9]/g).join("");
        //Gets the last Row for Test Suite (To know from which row to take the link of the request from in case any new test is added)
        var testSuiteLastRow = ss.getSheetByName('Test Suite').getLastRow()

        //Discover the rows
        var firstRowInfos = sheet.getRange(1, 1, 1, newLastColumn).getValues();
        var infoColNumber = firstRowInfos[0].indexOf(headerDate) +1;
        var infoColData = sheet.getRange(1, infoColNumber, sheet.getMaxRows(), 1).getValues();
        for(p=0;p<infoColData.length;p++){
          if(infoColData[p] == headerQAedBy){
            var rowQA = p + 1;
          } else 
          if(infoColData[p] == headerDevice){
            var rowDevice = p + 1;
          } else 
          if(infoColData[p] == headerExpectedResult){
            var rowExpectedResult = p + 1;
          } else
          if(infoColData[p] == headerReleaseBugID){
            var rowBugID = p + 1;
          }
        }

        //Construct the formulas with grabbed info
        var formulaLinkFromTestSuite = `=if(${newColumn}1='Test Suite'!${newColumn}1, if(${newColumn}2='Test Suite'!${newColumn}2,'Test Suite'!${newColumn}${testSuiteLastRow}, "NoMatch"),"")`
        var formulaDevice = `=if(${newColumn}2="Nightly",if(WEEKDAY(${newColumn}1)=2,"Pixel 4",if(weekday(${newColumn}1)=3,"Pixel 4",vlookup(${newColumn}${rowQA},IMPORTRANGE("1dfkQB5iKJKB4iCiOQd-M_aGgcoF38LWNVI3PxgKpV1Y","Callphone Type!$A1:$B10"),2,0))),vlookup(${newColumn}${rowQA},IMPORTRANGE("1dfkQB5iKJKB4iCiOQd-M_aGgcoF38LWNVI3PxgKpV1Y","Callphone Type!$A1:$B10"),2,0))`

        //Input the formulas
        sheet.getRange(rowBugID,newLastColumn).setValue(formulaLinkFromTestSuite); //release bug ID
        sheet.getRange(rowQA, newLastColumn).setValue(ldap) //QA by
        sheet.getRange(rowExpectedResult-2, newLastColumn, 2,1).setValue("N/A");
        sheet.getRange(rowExpectedResult, newLastColumn).setValue('Current Results');
    
        //Create verifications to put the Device exceptions
        var devices = ss.getSheetByName("Test Suite").getRange("B3:D").getValues()
        var deviceIsNos = []
        var deviceIsLaptop = []
        for(r=0;r<devices.length;r++){
          if(devices[r][0]=="NOS- Debug"){
            deviceIsNos.push(devices[r][2])
          }
          if(devices[r][0]=="Laptop (Corp)"){
            deviceIsLaptop.push(devices[r][2])
          }
        }

        //Input device exceptions
        if(deviceIsNos.includes(sheetNames[a])){
          sheet.getRange(rowDevice, newLastColumn).setValue("NOS- Debug")
        }
    
        else if(deviceIsLaptop.includes(sheetNames[a])){
          sheet.getRange(rowDevice, newLastColumn).setValue("Laptop (Corp)")
        }

        //Input the personal devices
        else {
          sheet.getRange(rowDevice, newLastColumn).setValue(formulaDevice)
        }
      }
    }
  //Prepare info for formulas in the test suite
    var sheet = ss.getSheetByName('Test Suite');
    var lastColumnNumber = sheet.getLastColumn();

    //Loop though the lines and put the formulas.
    for(t=0;t<sheet.getLastRow()-3;t++){
      var formulaSuiteVerification = `=if(${newColumn}1='${tests[t]}'!${newColumn}1, if(${newColumn}2='${tests[t]}'!${newColumn}2,'${tests[t]}'!${newColumn}3, "NoMatch"),"")`
    
      sheet.getRange(t+3,newLastColumn).setValue(formulaSuiteVerification)
    }

    //Alert confirming everything was done.
    ui.alert('Done! ✅')

    //Send verification email
    var message = `<p><strong>${ldap}</strong> just added a test for ${date} - ${shift} and the operation was finished.</p>
  <p><strong>Time of operation:</strong> ${new Date()}.
  <p><strong>Address of new test:</strong> ${sheet.getRange(1,lastColumnNumber,2,1).getA1Notation()}.
  <p><strong>Link to the test in the Test Suite:</strong> https://docs.google.com/spreadsheets/d/1ib91qjeFkuzLo3qMm1zh5o2OiSZ9X0NuJgoaH9yGdQo/edit#gid=438535577&range=${sheet.getRange(1,lastColumnNumber,2,1).getA1Notation()}</p>`
   email(message)
    } else {
    //If the user clicked cancel on the alert, prompt this.
      ui.alert('Operation canceled. ❌')
      var message = `<p><strong>${ldap}</strong> just clicked to add a test for ${date} - ${shift} but canceled the operation.</p>
      <p><strong>Time of operation:</strong> ${new Date()}.`
      email(message)
    }
}

function email(message){
  var emails = "zanollachaves@google.com";
  var subject = "Automated trix - model"
  MailApp.sendEmail({
    to: emails,
    subject: subject,
    htmlBody: message
  })
}

function onOpen() {
ui.createMenu('Automations')
.addSubMenu(ui.createMenu('Start new test')
.addItem('Nightly test', 'nightNew')
.addItem('Afternoon test', 'afternoonNew')
.addItem('Extra test', 'extraNew'))
.addItem('Hide/show columns', 'hideColumnsAllSheets')
.addSeparator()
.addItem('Add new Use Case', 'addNewUseCase')
.addItem('Help section', 'helpSection')
.addToUi();
}

function nightNew(){
  let shift = "Nightly";
  newTest(shift)
}

function afternoonNew(){
  let shift = "Afternoon";
  newTest(shift)
}

function extraNew(){
  let shift = "Extra Release";
  newTest(shift)
}

function helpSection(){
  //Generate the form.
  var html = HtmlService.createHtmlOutputFromFile('Help section').setWidth(1920).setHeight(1080)
  ui.showModalDialog(html, "Help section")
}


function timeStamp(e){ 
 //Take info from the sheet
 var row = e.range.getRow();
 var col = e.range.getColumn();
 var newValue = e.source.getActiveSheet().getRange(row, col).getValue();
 var sheet = e.source.getActiveSheet().getName(); 
 var firstRowElements = e.source.getActiveSheet().getRange(1, 1,1,e.source.getActiveSheet().getMaxColumns()).getValues();
 var infoColumnNumber = firstRowElements[0].indexOf(headerDate) +1;
 var infoColumnData = e.source.getActiveSheet().getRange(1, infoColumnNumber, e.source.getActiveSheet().getMaxRows(),1).getDisplayValues();

  for(a=0;a<infoColumnData.length;a++){
    if(infoColumnData[a] == headerTimeOfExecution){
      var rowTimeOfExecution = a + 1;
      break
    }
  }

 //Transform the list in a readable version
 var sheetNames = []
  for(b=0;b<tests.length;b++){
   sheetNames.push(tests[b][0]);
  }
  //Verify if the tab edited is a test
 var verification = sheetNames.includes(sheet)
 //Create the Time value
 var time = Utilities.formatDate(new Date(), "GMT", "hh:mm:ss a")
 //Verify if it's a test, if the new value is different from "" and if the row is the overall result
 if(verification && newValue !== "" && row == 3){
   //Put the value in Row 11.
 ss.getSheetByName(sheet).getRange(rowTimeOfExecution, col).setValue(time)
 }
}

function onEdit(e) {
 timeStamp(e)
}


function hideColumnsAllSheets(){
  //Create the alert to double check if the operation is intended.
  var maxNumber = ss.getSheetByName("Test Suite").getMaxColumns()-6

 var response = ui.prompt(`⚠️Insert how many tests you'd like to keep visible for all tabs.⚠️
  (It must be between 1 and ${maxNumber})`).getResponseText()


 if(isNaN(response) || response == 0 || response > maxNumber){
   ui.alert(`Response needs to be a number between 1 and ${maxNumber}`)
   Logger.log(ldap + " tried to hide columns but inserted - " + response + " - " + date)
 } else {
   Logger.log(response)
   //Create the loop to perform verifications
   for(i=0;i<tabs.length;i++){
     //Get the names for each existing sheet.
     var name = tabs[i].getSheetName();
    
     //For each sheet, verify if its name is equal to one of the test names on the Test Suite
     for(b=0;b<tests.length;b++){
       if(tests[b]==name){
         //Verify if the sheet is hidden or not, meaning the test is performed or paused.
         if(!ss.getSheetByName(name).isSheetHidden()){
          
           //Creates the list where the new columns are going to be added.         
           sheetNames.push(name);
         }
       }
     }
   }

    //Cycle through the tests and hide all the requested tabs
   for(a=0;a<sheetNames.length;a++){
     var sheet = ss.getSheetByName(sheetNames[a])
     var lastColumn = sheet.getLastColumn()
     sheet.showColumns(1,lastColumn)
     var lastColumn = sheet.getLastColumn();
     var range = lastColumn-response-5
     sheet.hideColumns(6,range)
   }


 }
}


function addNewUseCase() {
  //Generate the form.
  var html = HtmlService.createHtmlOutputFromFile('New Use Case').setWidth(400).setHeight(250)
  ui.showModalDialog(html, "Create new Use Case")
}

function submitForm(useCaseName, useCaseType) {
  // gets the answers of the form
  var testSuite = ss.getSheetByName("Test Suite");
  var testTemplate = ss.getSheetByName("Test template QA");
  var newTest = testTemplate.copyTo(ss).setName(useCaseName);

  newTest.showSheet();
  newTest.activate();
  
  var testSuiteRange = testSuite.getMaxColumns();
  var differenceOfTabs = testSuiteRange - 6;

  //In the new tab, put the name and the case device in B3:B4  
  newTest.getRange(2,2).setValue(useCaseName);
  newTest.getRange(3,2).setValue(useCaseType);
  
  //Insert as many columns as the test suite and hide all the previous phantom columns
  newTest.insertColumnsAfter(6,differenceOfTabs); 
  newTest.hideColumns(6, differenceOfTabs);

  //Get the tests Lenght in the test suite and add the name to the list;
  var testSuiteRows = testSuite.getMaxRows();
  testSuite.insertRowsAfter(testSuiteRows-1, 1);
  testSuite.getRange(testSuiteRows, 2).setValue(useCaseType);
  testSuite.getRange(testSuiteRows, 4).setValue(useCaseName);
  testSuite.getRange(testSuiteRows-1,1,1,testSuiteRange).copyTo(testSuite.getRange(testSuiteRows,1,1,testSuiteRange), SpreadsheetApp.CopyPasteType.PASTE_CONDITIONAL_FORMATTING, false); 

  //verify and delete duplicates.
  for(a=0;a<tabs.length;a++){
    var currentTabName = tabs[a].getSheetName();
    if(currentTabName == "Copy of Test template QA" || currentTabName == "Copy of Copy of Test template QA"){
      ss.deleteSheet(tabs[a])
    }
  }

  ui.alert(`You've successfully added the "${useCaseName}" test. All the automations are already working on this tab and on the Test Suite.
  
  📌 Next step:
     👉🏻 Write the test steps on the dedicated space and provide the screenshots.
     
  📌 For the automations to work we have to have:
    👉🏻 One header "Date" in the first row;
    👉🏻 In the same column as "Date", we need the titles "Release Bug ID", "Time of execution", "Qa'ed by", "Device" and "Expected Result".`)
}
