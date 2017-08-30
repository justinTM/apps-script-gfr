// Global Variables. Not best practice, but faster than Script Properties. Especially thrifty in onEdit(), which needs to be as lightweight as possible
// !!! -- Change these values when columns / headings are changed -- !!!
var firstBarcodeRow = 4;
var firstBarcodeColumn = 5;
var firstLocationColumn = 7;
var concatStringColumn = 2;
var masterLastRow = 400;

var quickDescriptionCell = "E3";
var quickBarcodeCell = "F3";
var quickDescriptionColumn = "E";
var quickBarcodeColumn = "F";

var quickBarcodeFormula = "=IFERROR(VLOOKUP("+quickDescriptionCell+",$"+quickDescriptionColumn+"$"+firstBarcodeRow+":"+quickBarcodeColumn+",2,FALSE))";
var quickDescriptionFormula = "=IFERROR(INDEX("+quickDescriptionColumn+firstBarcodeRow+":"+quickDescriptionColumn+",MATCH("+quickBarcodeCell+","+quickBarcodeColumn+firstBarcodeRow+":"+quickBarcodeColumn+",0),1))";

// !!! -------------------------- !!!

var firstLocationIndex = firstLocationColumn - firstBarcodeColumn;





// This function is run every time a cell is edited by the user, try to keep the code lightweight (eg. only update if cells are of-interest, ignore others; use minimal SpreadsheetApp function calls)
function onEdit(e) {

  var rawCell = e.range;
  var cellValue = e.value;
  var cell = rawCell.getA1Notation(); // set the cell to a cell notation (eg. C5, A1, etc.)
  var sheet = e.source.getActiveSheet();
  
  // Ignore the protected sheet "Master Inventory" (a read-only sheet)
  if (sheet.getName() !== "Master Inventory") return;
  
  
  // -- BEGIN quick barcode updating --
  // if the cell edited by user is "Item Description" in quick-add row, auto-fill the item's "Barcode"
  if (cell == quickDescriptionCell) { 
    var barcodeCellFormula = sheet.getRange(quickBarcodeCell).getFormula();
    
    // A difference in formulas indicates the user 
    if (barcodeCellFormula != quickBarcodeFormula) {
      sheet.getRange(quickBarcodeCell).setFormula(quickBarcodeFormula) // if data was ENTERED, look up the item's barcode
    }
      
  // otherwise, if the cell edited by user is item's "Barcode" in quick-add row, auto-fill the item's "Barcode"
  } else if (cell == quickBarcodeCell) {
    var descriptCellFormula = sheet.getRange(quickDescriptionCell).getFormula();

    if (descriptCellFormula != quickDescriptionFormula) {
      sheet.getRange(quickDescriptionCell).setFormula(quickDescriptionFormula) // if data was ENTERED, look up the item's description
    }
  }
  // -- END quick barcode updating --
  
  
  // -- BEGIN location concatenation updating -- (eg. Column B "Location" cells look like: "10 in Right Side B, 5 in Right Side A")
  // Only update locations when a user edits data in a location-related cell. Previously, we UPDATE the cells if the cells were IN range, now we EXIT if the cells are OUT of range
  var locationsRange = {
    top : firstBarcodeRow,
    bottom : 399,
    left : firstLocationColumn,
    right : 49
  };

  // Exit if we're out of this locationsRange
  var thisRow = e.range.getRow();
  if (thisRow < locationsRange.top || thisRow > locationsRange.bottom) return;

  var thisCol = e.range.getColumn();
  if (thisCol < locationsRange.left || thisCol > locationsRange.right) return;

  // Update the Locations column (ie. column B) to represent the current location quantities (eg. cells might look like: "10 in Right Side B, 5 in Right Side A")
  collectLocations();  
  // -- END location concatenation updating --
}




// onOpen() runs whenever the Spreadsheet is opened (or refreshed). This code adds custom menu items to the toolbar
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{name : "Check Duplicates",functionName : "checkDuplicates"}];
  sheet.addMenu("Scripts", entries);
  
  ui.createMenu('Manage Inventory')      
  .addItem('Remove Items from Inventory (dialog)', 'showDialogRemove') 
  .addItem('Add Items to Inventory (dialog)', 'showDialogAdd')
  .addItem('Create Part List', 'showDialogPartListCreation')
  .addToUi();
  
  ui.createMenu('Manage Barcodes')
  .addItem('Barcode Generator (dialog)', 'showDialogBarcodeCreation')
  .addItem('Barcode Generator (bulk)', 'showDialogBulkBarcodeCreation')
  .addToUi();
  
  ui.createMenu('Manage Sheet')      
  .addItem('Generate "Items Needed" List', 'generateList') 
  .addToUi();
  
  ui.createMenu('About')
  .addItem('Changelog', 'showChangelog')
  .addToUi();
}



// --- BEGIN dialog boxes ---

// Shows the "Add a Part" to inventory dialog (see file '<templateName>.html')
function showDialogAdd() {
  var title = 'Add Parts to Locations'; 
  var templateName = 'addUi'; 
  var width = 600; 
  
  // custom function to insert part data into the dialog box. For code modularity and simplicity
  createDialogWithPartData(title,templateName,width);
}


// Shows the dialog box for the creation of a "Parts-Needed" list (see file '<templateName>.html')
function showDialogRemove() {
  var title = 'Remove Parts to Inventory'; 
  var templateName = 'removeUi'; 
  var width = 600; 
  
  // custom function to insert part data into the dialog box. For code modularity and simplicity
  createDialogWithPartData(title,templateName,width);
}


// Shows the dialog box for the creation of a "Parts-Needed" list (see file '<templateName>.html')
function showDialogPartListCreation() {
  var title = 'Create Part List'; 
  var templateName = 'createPartList'; 
  var width = 600; 
  
  // Protect read-only sheets, like "Master Inventory"
  if (isProtectedSheet() == true) {
    SpreadsheetApp.getUi().alert("You cannot overwrite the master inventory, please create a new sheet and try again");
    return;
  }
  
  createDialogWithPartData(title,templateName,width);
}


// Shows the dialog box for the creation of an item's barcode (see file '<templateName>.html')
function showDialogBarcodeCreation() {
  var title = 'Generate and Export Barcodes'; 
  var templateName = 'barcodes'; 
  var width = 800; 
  
  createDialog(title,templateName,width);
}


// Shows the dialog box for the creation of many barcodes, from the "" (see file '<templateName>.html')
function showDialogBulkBarcodeCreation() {
  var title = 'Generate Barcodes from Sheet'; 
  var templateName = 'bulkBarcodes'; 
  var width = 800; 
  
  createDialogWithAllData(title,templateName,width);
}


// Shows the dialog box for the changelog for this app. Very outdated.
function showChangelog() {
  var title = 'Changes and To-Do List'; 
  var templateName = 'changelog'; 
  var width = 900; 
  
  createDialog(title,templateName,width);
}




// This function creates a modal (moveable) dialog box, given various parameters. 
function createDialog(title, templateName, width) { 
  var ui = SpreadsheetApp.getUi();
  var templateContent = HtmlService.createTemplateFromFile(templateName).evaluate().getContent();
  var html = HtmlService.createTemplate(templateContent)
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(width);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ui.showModalDialog(html,title);
}




// Creates a modal dialog box, and appends only the part data (using "getInvData()")
// inside a <script> tag at the end of the html file
function createDialogWithPartData(title,templateName,width) { 
  var ui = SpreadsheetApp.getUi();
  var createUi = HtmlService.createTemplateFromFile(templateName).evaluate().getContent();
  var html = HtmlService.createTemplate(createUi+
                                        "<script>\n" +
                                           "var data = "+getInvData()+
                                        "</script>")
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(width);
  
    var ss = SpreadsheetApp.getActiveSpreadsheet();
  ui.showModalDialog(html,title);
}




// Creates a modal dialog box, and appends the entire sheet's cell values at the end of the html file
function createDialogWithAllData(title,templateName,width) { 
  var ui = SpreadsheetApp.getUi();
  var createUi = HtmlService.createTemplateFromFile(templateName).evaluate().getContent();
  var html = HtmlService.createTemplate(createUi+
                                        "<script>\n" +
                                           "var data = "+getAllData()+
                                        "</script>")
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(width);
  
    var ss = SpreadsheetApp.getActiveSpreadsheet();
  ui.showModalDialog(html,title);
}



// Returns an array of column header strings, given any sheet's range
function getHeaders(sheet,range,columnHeadersRowIndex) {
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  
  // The 1D range of sheet headers (column titles) for a given range
  var headersRange = sheet.getRange(columnHeadersRowIndex, firstBarcodeColumn, 1, numColumns);

  return headersRange.getValues()[0];
}



function synchronousSort(sheetToSort, columnToSort) {
  sheetToSort.sort(columnToSort);
  return true;
}



// This function breaks apart merged cells in column A (part categories), sorts column A,
  //re-merges cells with matching values, then applies a border between each group of merged cells
function unmergeThenSortByCategoryAndFixBorders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Master Inventory');
  var range = sheet.getRange(firstBarcodeRow, 1, sheet.getLastRow());
  
  unmergeAndDuplicateValues();
  var result = synchronousSort(sheet, 1);
  
  if (result) {
    applyBottomBorderToGoupedValues();
    mergeCellsContainingDuplicatesValues();
  }
}




// This function breaks apart any merged cells in column A (part categories), 
  // then copies the merged-cell-value to each respective individual cell
function unmergeAndDuplicateValues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Master Inventory');
  var range = sheet.getRange(firstBarcodeRow, 1, sheet.getLastRow());
  range.breakApart();
  
  var columnToSearchZeroIndexed = 0;
  var rowRangesArray = getRowRangesOfGroupedValues(columnToSearchZeroIndexed);
  var numGroups = rowRangesArray.length;
  
  for (var i=0; i < numGroups; i++) {
    var firstTempRow = rowRangesArray[i][0];
    var lastTempRow = rowRangesArray[i][1];
    var valueToDuplicate = rowRangesArray[i][2]
    
    var tempRange = sheet.getRange("A"+firstTempRow+':'+"A"+lastTempRow);
    tempRange.setValue(valueToDuplicate);
  }

}



// This function places a border between different groups of values in column A (part categories)
function applyBottomBorderToGoupedValues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Master Inventory');
  var numColumns = sheet.getLastColumn();
    
  var columnToSearchZeroIndexed = 0;
  var rowRangesArray = getRowRangesOfGroupedValues(columnToSearchZeroIndexed);
  var numGroups = rowRangesArray.length;
  
  for (var i=0; i < numGroups; i++) {
    var firstTempRow = rowRangesArray[i][0];
    var lastTempRow = rowRangesArray[i][1];
    var numTempRows = lastTempRow - firstTempRow;
    
    var tempRange = sheet.getRange(firstTempRow, columnToSearchZeroIndexed+1, numTempRows+1, numColumns);
    tempRange.setBorder(false, false, true, false, false, false);
  }
}



// This function finds cells in column A (part category) with matching values and merges them
function mergeCellsContainingDuplicatesValues() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Master Inventory');
  
  var columnToSearchZeroIndexed = 0;
  var rowRangesArray = getRowRangesOfGroupedValues(columnToSearchZeroIndexed);
  var numGroups = rowRangesArray.length;
  
  for (var i=0; i < numGroups; i++) {
    var firstTempRow = rowRangesArray[i][0];
    var lastTempRow = rowRangesArray[i][1];
    
    var tempRange = sheet.getRange("A"+firstTempRow+':'+"A"+lastTempRow);
    tempRange.mergeVertically();
  }
}




// This function gathers duplicate-valued cells in column A (part category column)
  // then returns a 2D array of row number ranges (eg. [[1,7], [8,13]]) for these groups
// If rows 4 through 8 had the value "Alternator", then the returned array would contain "[4,8]"
function getRowRangesOfGroupedValues(columnNumberZeroIndexed) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Master Inventory');
  var range = sheet.getRange(firstBarcodeRow,columnNumberZeroIndexed+1,sheet.getLastRow(),1);
  var data = range.getValues(); // [][]
    
  var rowRangesArray = [];
  
  // Track both the previous cell's value (lastValue) and the most recent marged-cell text (lastText)
  var lastValue = "";
  var lastText = "";
  var rowBegin = firstBarcodeRow+1;
  var rowEnd = firstBarcodeRow+1;
    
  // Loop through each row, copying the merged valued for all cells in each now-unmerged range
  for (var row in data) {
    var value = data[row][columnNumberZeroIndexed];
    var cellNumber = +row + +firstBarcodeRow;
    
    // Continually update the merged range's end, in order to update each cell in bulk whenever we reach the end of the merged cells
    if (value == lastValue || value == "") {
      rowEnd = cellNumber;
    }
    
    // If we encounter a NEW merged cell range, go ahead and apply the merged value to all cells of the previous range.
      // Skips the very first iteration (row = 0) by ignoring if the default value = ""
    else if (value != lastValue && value != ""){
      // Store this range of rows, plus the cell value, into an array. Append this array to rowRangesArray
      var tempRowsArray = [rowBegin, rowEnd, lastText];
      rowRangesArray.push(tempRowsArray);
      
      // Reset the range, for the next iteration of merged cell values
      rowBegin = cellNumber;
      rowEnd = cellNumber;
    }
    
    if (value != "")
      lastText = value;
    
    lastValue = value;
  }
  
  Logger.log(JSON.stringify(rowRangesArray));
  return rowRangesArray;
}




function merge() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Master Inventory');
  
  var range = sheet.getRange(firstBarcodeRow,1,sheet.getLastRow(),1);
  var data = range.getValues(); // [][]
  
  var lastValue = "";
  var lastText = "";
  var rangeBegin = "A1"+firstBarcodeRow;
  var rangeEnd = "A1"+firstBarcodeRow;
  
  range.breakApart();
  
  for (var row in data) {
    var value = data[row][0];
    var cellNumber = +row + +firstBarcodeRow;
    
    if (value == lastValue || value == "") 
      rangeEnd = "A" + cellNumber;
    else if (value != lastValue && value != ""){
      if (rangeBegin != rangeEnd) {
        var tempRange = sheet.getRange(rangeBegin+':'+rangeEnd)
        tempRange.mergeVertically()
      }
      rangeBegin = "A" + cellNumber;
      rangeEnd = "A" + cellNumber;
    }
    if (value != "")
      lastText = value;
    lastValue = value;
  }
} 




function borders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Master Inventory');
  
  var range = sheet.getRange(firstBarcodeRow,1,sheet.getLastRow(),1);
  var data = range.getValues(); // [][]
  
  var lastColumn = sheet.getLastColumn();
  
  var lastValue = "";
  var lastText = "";
  var rangeBegin = "A1"+firstBarcodeRow;
  var rangeEnd = lastColumn+firstBarcodeRow;
    
  for (var row in data) {
    var value = data[row][0];
    var cellNumber = +row + +firstBarcodeRow;
    
    if (value == lastValue || value == "") {
      rangeBegin = "A" + cellNumber;
      rangeEnd = lastColumn + cellNumber;
    }
    else if (value != lastValue && value != ""){
      if (rangeBegin != rangeEnd) {
        var tempRange = sheet.getRange(rangeBegin+':'+rangeEnd)
        tempRange.setBorder(false, null, true, null, null, null)
      }
      rangeBegin = "A" + cellNumber;
      rangeEnd = lastColumn + cellNumber;
    }
    if (value != "")
      lastText = value;
    lastValue = value;
  }
} 




function mergeDuplicates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Master Inventory');
  
  var range = sheet.getRange(firstBarcodeRow,1,sheet.getLastRow(),1);
  var data = range.getValues(); // [][]
  
  var lastValue = "";
  var lastText = "";
  var rangeBegin = "A1"+firstBarcodeRow;
  var rangeEnd = "A1"+firstBarcodeRow;
  
  sheet.getRange(firstBarcodeRow,1,sheet.getLastRow(),sheet.getLastColumn()).sort(1);
  
    for (var row in data) {
    var value = data[row][0];
    var cellNumber = +row + +firstBarcodeRow;
    
    if (value == lastValue || value == "") 
      rangeEnd = "A" + cellNumber;
    else if (value != lastValue && value != ""){
      if (rangeBegin != rangeEnd) {
        var tempRange = sheet.getRange(rangeBegin+':'+rangeEnd)
        tempRange.mergeVertically();
      }
      rangeBegin = "A" + cellNumber;
      rangeEnd = "A" + cellNumber;
    }
    if (value != "")
      lastText = value;
    lastValue = value;
  }
}




// Gets all row data in the inventory sheet, stores cell data as Object (eg. items[U-603]['description'] = "Uni Foam Air Filter")
// returns string (JSON.stringify(Object))
function getInvData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Master Inventory');
  var lastColumn = sheet.getLastColumn();
  
  var range = sheet.getRange(firstBarcodeRow,firstBarcodeColumn,masterLastRow,lastColumn);
  var invData = range.getValues(); 
  var headers = getHeaders(sheet,range,1);
  Logger.log("headers = " + headers)
  var items = {};
    
  for (var m in invData) {
    // We want to rename the object properties for columns not likely to change. 
    // We want to dynamically add all location column headers. 
    // Because these new column headers aren't constants, we'll need to add them separately using a for loop
    var item = {};

    item['description'] = invData[m][0];
    item['partNumber'] = invData[m][1];
    
    items['locationsList'] = {};
    // Location columns will change over time, this adds ALL column headers after column 5 (headers[4], "Location") to the new item Object
    for (var i = firstLocationIndex; i < headers.length; ++i) {
      item[headers[i]] = invData[m][i];
      //Create location properties from the headers. Used in the html scripts to check if the barcode entered is a location or a part
      items['locationsList'][headers[i]] = "";
    }    
    items[item.partNumber] = item; // assign this info to an object key = Part Number (eg. items['MX2030']['description'] = "12.5:1 Piston Kit")
  }
  return JSON.stringify(items)
}




// Concatenates each part's quantity with each location, and pastes it into the "Location" column (global variable 'concatStringColumn')
// Looks like "4 in Right Side T", with a new line for each location the part is in
function collectLocations() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Master Inventory');
  var lastColumn = sheet.getLastColumn();
  
  var range = sheet.getRange(firstBarcodeRow,firstBarcodeColumn,masterLastRow,lastColumn);
  var sheetValues = range.getValues(); 
  var locationNames = getHeaders(sheet,range,1);
  
  var paste = [];
  
  // Loop through all rows (ie. parts)
  for (var i in sheetValues) {
    var item = {concatString: ""}; 
    
    // Loop through columns (ie. location names)
    for (var j = firstLocationIndex; j < locationNames.length; ++j) {
      var locationQuantity = sheetValues[i][j];
      
      // Ignore empty locations
      if (locationQuantity > 0) {
        
        // Add a line break if there is a pre-existing string in the cell
        if (item['concatString']) {
          item['concatString'] += "\n";     
        }
        
        // Append the new locaiton string to the cell (eg. "4 in Right Side T")
        item['concatString'] += locationQuantity+" in "+locationNames[j];
      }
    }
    
    // Sheet values must be a 2D array (array[row][column])
    paste[i] = [item['concatString']]; // assign this info to an object key = Part Number (eg. items['MX2030']['description'] = "12.5:1 Piston Kit")
  }
  
  // Get the range representing the "Location" column, for each part row
  var range = sheet.getRange(firstBarcodeRow,concatStringColumn,masterLastRow);
  
  // Paste the entire 2D array of values we created previously
  range.setValues(paste)
  
  return
}




// Function to subtract quantities from each part, to update inventory
function removeItems(items) {
  Logger.log('running removeItems function!')

  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var itemsLength = Object.keys(items).length;  
  var sheet = ss.getSheetByName('Master Inventory');
  var range = sheet.getRange(firstBarcodeRow,firstBarcodeColumn,masterLastRow,sheet.getLastColumn());
  
  var data = range.getValues();
  var headers = getHeaders(sheet,range,1);
  var locations = headers.slice(firstLocationIndex); // Removes the headers which aren't locations (eg. headers[0] might equal 'Part Number' so we don't want it to become a location)
  
  var used = [];       
  for (var i = 0; i < data.length; i++) { // for each row in the sheet range values

      for (var partName in items) { // for each key in the items object (eg. "MX2030")
          if (data[i][1] == partName) { // if the part name in the sheet (row x, column 2 of the range) = part changed in the dialog        
            var item = items[partName];
              for (var k = 0; k < locations.length; k++) {                
                  var location = locations[k];
                  if (typeof item[location] !== "undefined") {
                    var before = data[i][k];
                    var cell = k + firstLocationIndex;
                    
                    if (data[i][cell] != "" || data[i][cell] > 0)
                      data[i][cell] = data[i][cell] - +item[location];
                    Logger.log("data["+i+"]["+cell+"] changed from: ' "+before+"' to: '"+data[i][cell]+"'")
                  }
              }
          }
      }
  }

  range.setValues(data);  
  collectLocations();
  
  return itemsLength;
}

// Warns the user to 
function isProtectedSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sheetName = sheet.getName();
  
  if (sheetName == "Master Inventory" || sheetName == "Qty Wanted + Prices") {
    SpreadsheetApp.getUi().alert("You cannot overwrite this inventory sheet, please create a new sheet and try again");
    return true;
  }
  else
    return false;
}
 // Puts a simple list of item info into a sheet
function pasteList(items) {
  Logger.log("items = " + JSON.stringify(items))
  var itemsLength = Object.keys(items).length; 
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  var range = sheet.getRange(1,1,itemsLength+1,2);
  var data = range.getValues();
  
  for (var row in data) {
    if (row == 0) {
      data[row][0] = "Quantity";
      data[row][1] = "Barcode";
    }
    if (items['row'+row]) {
      var item = items['row'+row];
      var adjustedRow = Number(row) + 1;
      
      data[adjustedRow][0] = item.qty;
      data[adjustedRow][1] = item.name;
    }
  }

  range.setValues(data);
}

// Creates an html table and saves it as a .pdf to your Drive
function listBarcodes(items) {
  Logger.log('running listBarcodes function!')
  Logger.log('var items = ' + JSON.stringify(items) + "\n")
  
  var itemsLength = Object.keys(items).length;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var sheet = ss.getSheetByName('Master Inventory');
  var range = sheet.getRange(firstBarcodeRow,firstBarcodeColumn,masterLastRow,sheet.getLastColumn());
  
  var data = range.getValues();
  var headers = getHeaders(sheet,range,1);
  var locations = headers.slice(firstLocationIndex); // Removes the headers which aren't locations (eg. headers[0] might equal 'Part Number' so we don't want it to become a location)
        
  for (var i = 0; i < data.length; i++) { // for each row in the sheet range values

      for (var partName in items) { // for each key in the items object (eg. "MX2030")
        
        var table = document.getElementById("myTable");
        var row = table.insertRow(1);
        var cell1 = row.insertCell(0);
        var cell2 = row.insertCell(1);
        
        cell1.innerHTML = partName;
              for (var k = 0; k < locations.length; k++) {
                  var location = locations[k];
                 
                  if (typeof item[location] !== "undefined") {
                     cell2.innerHTML = items[partName][location];
                  }
              }
      }
  }
  return itemsLength;
}

function addItems(items) {
  Logger.log('running addItems function!')  
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var itemsLength = Object.keys(items).length;  
  var sheet = ss.getSheetByName('Master Inventory');
  var range = sheet.getRange(firstBarcodeRow,firstBarcodeColumn,masterLastRow,sheet.getLastColumn());
  
  var data = range.getValues();
  var headers = getHeaders(sheet,range,1);
  var locations = headers.slice(firstLocationIndex); // Removes the headers which aren't locations (eg. headers[0] might equal 'Part Number' so we don't want it to become a location)
  
  var used = [];
          
  for (var i = 0; i < data.length; i++) { // for each row in the sheet range values
      for (var partName in items) { // for each key in the items object (eg. "MX2030")
          if (data[i][1] == partName) { // if the part name in the sheet (row x, column 2 of the range) = part changed in the dialog            
            var item = items[partName];
              for (var k = 0; k < locations.length; k++) {                
                  var location = locations[k];
                  if (typeof item[location] !== "undefined") {
                    var before = data[i][k];
                    var cell = k + firstLocationIndex;
                    data[i][cell] = data[i][cell] + +item[location];
                    Logger.log("data["+i+"]["+cell+"] changed from: ' "+before+"' to: '"+data[i][cell]+"'")
                  }
              }
          }
      }
  }

  range.setValues(data);
  collectLocations();
  
  return itemsLength;
}

function insertImage(barcodesObject) {
  Logger.log(JSON.stringify(barcodesObject))
  for(var key in barcodesObject) {
    // Create a document.
    var doc = DocumentApp.openById("1R_rvpcp-lF72_7SiWxK0Iac1YRC8xHbDbSkOQ3DPwjU");
    
    // Get the URL string within each Object property (property name is the barcode text, eg. "Testing" barcode is barcodeObjects.Testing)
    var URL = barcodesObject[key];
    
    // Retrieve an image from the web.
    var resp = UrlFetchApp.fetch(URL);
    
    // Append the image to the first paragraph.
    doc.getChild(0).asParagraph().appendInlineImage(resp.getBlob());
  }
}

function getEmail() {
  var email = Session.getActiveUser().getEmail();
  return(email)
}
