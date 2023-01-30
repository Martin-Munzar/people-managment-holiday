var book = SpreadsheetApp.getActiveSpreadsheet();
var sheetPeople = book.getSheetByName("xxx"); //CHANGE THIS TO THE NAME OF YOUR SHEET
var sheetDetailPeople = book.getSheetByName("xxx"); //CHANGE THIS TO THE NAME OF YOUR DETAIL SHEET

// EDIT THIS CONFIRURATION TO FIT YOUR TABLE
var startColumn; //IF THERE ARE SOME HIDDEN COLUMNS 
var startRowDetail = 6;
var startRowPeople = 3;
var weakCheck = 4; // -1

// initializing variables
var row = 0;
var collomn = 0;
var collomnPeople = 0;
var holidayTakenCounter = 0;

function main() {

  for(let k = 0; k < GetPeopleCount(); k++) { // repete for how much ppl
    collomn = 0;
    collomnPeople = 0;
     
    for(let i = 1; i < weakCheck + 1; i++) { // how many weaks to cheack
      var weak = [6]; 
      holidayTakenCounter = 0;

      for(let j = 0; j < 7; j++) { // cheack each day in a weak
        try {
          weak[j] = sheetDetailPeople.getRange(startRowDetail + row, collomn + FirstShownColumn("DetailPeople")).getValues(); 
          //Logger.log(weak[j]);  
          if(weak[j] == "Holiday") { holidayTakenCounter++; }  // if the day contaions "holyday" count it.
      
          collomn++;
        } catch (e) {
          throw new Error("Error: Could not retrieve data. Check the configuration.");
        }
      }

      if(holidayTakenCounter > 3) {  // if a person has taken more then 3 days off write it
        // Logger.log("**Writing**");
        try {
          sheetPeople.getRange(startRowPeople + row, collomnPeople + FirstShownColumn("People")).setValue("Holiday - most of the weak"); 
         } catch (e) {
            throw new Error("Error: Could not write data. Check the configuration.");
        }
        collomnPeople++;  
      }

      // Logger.log("----END OF WEAK " + i + "----");    
    }

  row++;
    
  // Logger.log("----END OF PERSON " + row + "----");   
  }
}

function GetPeopleCount() {
  var peopleCount = 60; // estimated people count to make the while faster

  while (sheetDetailPeople.getRange(peopleCount, 1).getValue() !== "") {
    peopleCount++;
  }

  return peopleCount - startRowDetail;
}

function FirstShownColumn(sheetName) {
  var estimatedStartColumn; // if there are some hidden columns 

  switch (sheetName) {
    case "People":
      estimatedStartColumn = 225;
      
      while (sheetPeople.isColumnHiddenByUser(estimatedStartColumn) == true) {
        estimatedStartColumn++;
      }
      
      startColumn = estimatedStartColumn; //Logger.log(startColumn);
      return startColumn;
    break;  
    
    case "DetailPeople":
      estimatedStartColumn = 1105;
      
      while (sheetDetailPeople.isColumnHiddenByUser(estimatedStartColumn) == true) {
        estimatedStartColumn++;
      }

      startColumn = estimatedStartColumn; //Logger.log(startColumn);
      return startColumn; 
    break;
  }
}
