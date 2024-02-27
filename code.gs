var incomingEntries = SpreadsheetApp.openByUrl(DOCUMENT URL);
var inventoryData = SpreadsheetApp.openByUrl(DOCUMENT URL); 

var index = incomingEntries.getLastRow(); 


// function letter index turns a column letter into an index 
function letterIndex(column) {
  const alphabetMap = { 'A': 1,'B': 2,'C': 3,  'D': 4,'E': 5,'F': 6,'G': 7,'H': 8,'I': 9};
  return alphabetMap[column];
}

// function column index takes in a column index and returns its letter 
function columnIndexToLetter(column) {
  let letter = String.fromCharCode(64 + column);
  return letter;
}     

// takes in a row and column (A1) and returns the range around it (A1:B2)
function surroundingRange(row, column){
  // row is number 
  // column comes in as the index (beacuse of how it is called in returnItemInCell)

  var originRow = row; 
  var origincolToChange = columnIndexToLetter(column); 

  // add one to final column and final row. This creates the range we will need to search later  
  var finalColumn= column + 1; 
  var finalRow = row + 1;

  // convert finalColumn back into a letter 
  var finalcolToChange = columnIndexToLetter(finalColumn); 

  return (origincolToChange + originRow + ":" + finalColumn + finalRow ); 
}

function returnItemInCell(row, column, spreadsheet){
  // row will come in as a number 
  // column will come in as a letter 
  return (spreadsheet.getRange(surroundingRange(row,letterIndex(column))).getValue()); 
} 

//action is stored in the column I 
function getAction(row){
  if (returnItemInCell(row, "I", incomingEntries) == "Adding item to inventory"){
    return "ADD"; 
  } else if (returnItemInCell(row, "I", incomingEntries) == "Removing item from inventory"){
    return "REMOVE"; 
  }
}

// item name is in column C 
function getItemName (row){
  return (returnItemInCell(row, "C", incomingEntries)); 
}  


function returnSizeScheme(row) {
  // Get values in relevant columns directly
  const values = [returnItemInCell(row,"D", incomingEntries),returnItemInCell(row,"E", incomingEntries),returnItemInCell(row,"F", incomingEntries)];

  for (let i = 0; i < values.length; i++) {
    if (values[i] !== ""){
      var colNumber=i; 
    }
  }

  // Map the columnNumber to the corresponding size scheme column
  switch (colNumber) {
    case 0:
      return "Standard";
    case 1:
      return "Form-Fit";
    case 2:
      return "Kids";
    default:
      return "none"; // No size data in any column
  }
}


// get the size value (XS, S, M, L, XL) which is in the column of the size scale  
function getSizeValue(row){
  sizeScheme = returnSizeScheme(row); 
  var colToLook; 
  switch (sizeScheme) {
      case "Standard":
        colToLook= "D";
        return (returnItemInCell(row, colToLook, incomingEntries));
      case "Form-Fit":
        colToLook= "E";
        return (returnItemInCell(row, colToLook, incomingEntries));
      case "Kids":
        colToLook= "F";
        return (returnItemInCell(row, colToLook, incomingEntries));
    }
}


// get the inventory delta which is column G
function getInventoryDelta(row){
  if (getAction(row) == "REMOVE"){
    return(returnItemInCell(row,"G",incomingEntries)); 
  } else if (getAction(row) == "ADD"){
    return(returnItemInCell(row,"H",incomingEntries)); 
  }
  
}

// changes column H to "TRUE", this will trigger when the row has been processed
function logDataProcessed(row){
  incomingEntries.getRange("J" + row).setValue("TRUE");
}

// uses a row from the changelog and locates the corresponding row in the CURRENT INVENTORY tab needs to be changed 
function findRowToChange(inputRow){
 var range = inventoryData.getRange('A1:AA28');

  // Creates  a text finder for the range.
  var textFinder = range.createTextFinder(getItemName(inputRow));

  // Returns the first occurrence of 'dog'.
  var firstOccurrence = textFinder.findNext();
  var rowNumber = firstOccurrence.getRow();
  return(rowNumber);
}


// uses a row from the changelog and locates the appropriate column to change
function findColumnToChange(logRow){
  // determine the size scheme of the item in the input row 
  var scheme = returnSizeScheme(logRow); 

  // determine the size of the item 
  var size = getSizeValue(logRow); 

  switch (scheme) {
    case "Form-Fit": switch(size){
      case "XS": 
        return "J"; 
      case "S":
        return "K"; 
      case "M":
        return "L"; 
      case "L": 
        return  "M"; 
      case "XL": 
        return "N"; 
      case "XXL": 
        return "O"; 
    } 
    case "Standard": switch(size){
      case "XS": 
        return  "P"; 
      case "S":
        return  "Q"; 
      case "M":
        return "R"; 
      case "L": 
        return  "S"; 
      case "XL": 
        return  "T"; 
      case "XXL": 
        return "U"; 
    }
    case "Kids": switch(size){
      case "XS": 
        return  "V"; 
      case "S":
        return   "W"; 
      case "M":
        return "X"; 
      case "L": 
        return  "Y"; 
      case "XL": 
        return   "Z"; 
      case "XXL": 
        return   "AA"; 
    }
  }
}



function tendInventory(){ 
  try {
      let debug = false; 

      if (debug){
    console.log("item name " + getItemName(index));
    console.log("action: " + getAction(index)); 
    console.log("latest incoming entry size scale: " + returnSizeScheme(index)); 
    console.log ("get the size value for last row: " + getSizeValue(index));
    console.log("get the inventory delta for the last item: " + getInventoryDelta(index));
    
    // log the row and column of the target cell in the inventory spreadsheet 
    console.log("row: " + findRowToChange(index)); 
    console.log("column: " + findColumnToChange(index)); 
  } 

  // check that the data hasn't already been processed 
  if (logDataProcessed(index) == "TRUE"){
    console.log("data has already been processed - doing nothing instead "); 
  } else {
    // get the current inventory value of that item
    var currentValueOfCell = inventoryData.getRange(findColumnToChange(index)+findRowToChange(index)).getValues()
    if (debug){console.log(currentValueOfCell[0][0])}  

    // find how many are added/subtracted from the current value 
    var delta = getInventoryDelta(index); 

    // iff adding to the cell, set the new value to currentValue + delta 
    if (getAction(index) == "ADD"){
      //var output = currentValueOfCell.valueOf() + delta.valueOf(); 
      const numberValue = Number(currentValueOfCell)
      var output = numberValue + delta;
      console.log("number output " + output); 
      inventoryData.getRange(findColumnToChange(index) + findRowToChange(index)).setValue(output);
      if (debug){console.log("Added " + delta + "to " + currentValueOfCell)}
    } else if (getAction(index) == "REMOVE") {
      // otherwise, subtract from that cell
      inventoryData.getRange(findColumnToChange(index) + findRowToChange(index)).setValue(currentValueOfCell-delta);
      if(debug){console.log("Subtracted " + delta + "from " + currentValueOfCell)}
    }

    // report the current cell value 
    if(debug){
      var postOperationCellValue = inventoryData.getRange(findColumnToChange(index) + findRowToChange(index)).getValue(); 
      console.log("post operation cell value! " + postOperationCellValue);    
    }

    logDataProcessed(index);
  } // end of if data hasn't already been processed loop 

    } catch (error) {
      // send an email to Ashley saying something failed
      GmailApp.sendEmail("EMAIL@ORGANIZATION.org, EMAIL@ORGANIZATION.org", "Subject: Failed Inventory Form Connection", 
      "Something went wrong, you should check the inventory document: DOCID or the inventory tending form log: DOCID or the submittal form DOCID")
    }
} // end of tend inventory 

