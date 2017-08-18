// Runs whenever the PES spreadhseet is opened
// Appends menu item(s) to the toolbar
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Admin Tools')
  .addItem('Email part responsibles', 'ShowPasswordPrompt')
  .addToUi();
}




// Show the modal dialog box
function ShowEmailDialog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var html = HtmlService.createHtmlOutputFromFile('EmailDialog');
  html.setWidth(800)
      .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Send emails');
}


// Emails each recipient, using arg 'userData' array
//  (eg. emails[0]['contents'] = "blah<br>blah", emails[0]['recipients'] = "first.last@ba-racing-team.de")
function SendEmails(userData) {
  var senderName = 'PES Admins';
  
  // For each user
  for (var i=0; i < userData.length; i++) {
    var sheetName = userData[i]['data'][0]['sheetName'];
    var subject = 'Missing Data in Your PES Part(s) from '+sheetName;
    var recipient = userData[i]['user'];
    var firstName = recipient.substring(0, recipient.indexOf("."));
    firstName = firstName.replace(/^./, function(str){
      return str.toUpperCase(); 
    });
    
    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      htmlBody: userData[i]['emailBody'],
      name: senderName
    });
  }
}



// Adds an html email to, then returns, userData array.
// 'userData' is array of objects, 'emailBody' is custom user-made string from textbox in form dialog
function MakeEmails(userData, emailBody) {

  // For each user, construct an html email and add to userData array
  for (var i=0; i < userData.length; i++) {
    // Extract first name from 'first.last@email.com'
    var recipient = userData[i]['user'];
    var firstName = recipient.substring(0, recipient.indexOf("."));
    firstName = firstName.replace(/^./, function(str){
      return str.toUpperCase(); 
    });
    
    // Write email and add table of data to the body
    // note: 'emailBody' is a string, from field from the dialog's textbox
    var htmlTable = userData[i]['table'];
    var body = 'Hello '+firstName+', <br><br>';
    body += emailBody;
    body += '<br><br>'+htmlTable+'<br><br>';
    body += 'Thank you,<br>PES Admins';
    
    // Make table and table header
    var style = 'td {text-align: left; padding: 8px; border: 1px solid #ddd;}';
    style += 'th {text-align: center; padding: 8px; border: 1px solid #ddd;}';
    
    // Wrap contents with html tags, push to userData array
    var htmlEmail = '<html><head><style type="text/css">' +style+ '</style></head><body>' +body+ '</body></html>';
    userData[i]['emailBody'] = htmlEmail;
  }
  
  Logger.log(JSON.stringify(userData));
  return userData;
}




// Sorts, ascendingly, a sheet (object) according to a columnIndex (int)
function SortSheet(sheet, columnIndex) {
  sheet.sort(columnIndex);
  return true;
}




// Get missing data from each part in each sheet, return an array of objects, containing html table of missing data
// eg. htmlTablesPerSheet[0]['sheet'] = 'Chassis'
// eg. htmlTablesPerSheet[0]['table'] = '<table><tr><td>user.name@ba-racing-team.de</td></tr></table>'

// Called inside EmailDialog.html (gathers user-selected sheets and columns)
function GetMissingPartData(sheetNames, requiredColumns) {
  var htmlTablesPerSheet = [];
  var missingDataAllSheets = [];
  var allUsersData = [];
  var previousUser = '';
  var previousUserIndex = 0;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  for (var i=0; i < sheetNames.length; i++) { // for each sheet
    var sheetValues = [["Failed to get this sheet's values"]];
    var sheetName = sheetNames[i];
    var sheet = ss.getSheetByName(sheetName);
    var columns = GFR.GetColumnHeadings(sheet);
    var columnIndexToSort = columns['Creator'];
    if (columnIndexToSort) 
      var sorted = SortSheet(sheet, columnIndexToSort);
    
    if (sorted)  // Forces code to run synchronously, otherwise it might grab cell values before sheet is sorted
      var sheetValues = sheet.getDataRange().getValues();
    sheet.sort(1);
    
    var lastRow = sheetValues.length;
    var lastColumn = sheetValues[0].length;
    
    for (var j=0; j < lastRow; j++) {  // for each part (sheet row)
      if (!sheetValues[j][0])  // skip sheet rows without part numbers in them (ie. if [row n, column 1] is blank)
        break;
      
      var missingData = [];
      var anythingMissing = false;
      
      for (var k=0; k < lastColumn; k++) {  // for each column (part data, like 'Cost')
        
        var columnName = sheetValues[0][k];
        var cellContents = sheetValues[j][k];
        
        for (var m=0; m < requiredColumns.length; m++) {  // for each required column (selected by user in checkbox)
          
          if (!cellContents) {  // if cell is blank (missing data)
            if (columnName == requiredColumns[m]) {  // if cell is a required column
              anythingMissing = true; 
              missingData.push(sheetValues[0][k]);
            }
          }
        }
      }
      
      if (anythingMissing) { // If any data is missing from row
        var user = sheetValues[j][columns['Creator']-1];
        var thisPartData = {
          sheetName: sheetName,
          partNumber: sheetValues[j][columns['Part Number']-1],
          partName: sheetValues[j][columns['Part Name']-1],
          missingData: missingData
        };
        
        if (user == previousUser) { // if this is an existing user in the array, append the data to their list
          allUsersData[previousUserIndex]['data'].push(thisPartData);
        }
        else { // otherwise, if this user not already in the array, add them
          allUsersData.push({
            user: user,
            data: [thisPartData]
          });
        }
        
        previousUser = user;
        previousUserIndex = allUsersData.length - 1;
      }
    }
    
    // make array of headers
    var tableHeaders = [
      'sheetName',
      'partNumber',
      'partName',
      'missingData'
    ];

    previousUser = user;
    previousUserIndex = allUsersData.length - 1;
  }
  
  return allUsersData;
}




// Make array of missing part data, sorted by user email
// NOT CURRENTLY USED
function SortMissingData() {
  var newArray = [];
  var missingData = GetMissingPartData(['Chassis', 'Aero'], ['Cost ($)']);
  
  // For each sheet in missingData
  for (var i=0; i < missingData.length; i++) {
    
    var previousUser = '';
    var previousUserIndex = 0;
    // For each part (row) in missingData
    for (var j=0; j < missingData[i].length; j++) {
      
      var user = missingData[i][j]['user'];
      Logger.log('User: %s', user);
      
      if (user == previousUser) { // if this part is from same user as previous part, add this data object to the user
        newArray[previousUserIndex]['data']
        .push(missingData[i][j]);
      }
      else { // if this is a new user, add this user to the new array
        newArray.push({
          user: user,
          data: [missingData[i][j]]
        });
      }
      
      previousUser = user;
      previousUserIndex = newArray.length - 1;
    }
  }
  
  Logger.log(JSON.stringify(newArray));
}




// Generate an HTML table from header and rows array of objects (each header element must match an object key in rows array)
function MakeHTMLTable(headers, rows) {
  Logger.log('\n\nHeaders = %s\n\n', headers);
  Logger.log('\ninputted data array = %s\n', JSON.stringify(rows));
  var table = '';
  
  // Format array of data into HTML cells
  for (var i=0; i < rows.length; i++) { // each row (user)
    var tableRow = '';
    for (var j=0; j < headers.length; j++) { // each column
      var header = headers[j];
      Logger.log('Header = %s', header);
      var cell = rows[i][header];
      
      // If the cell is a list of item, make it bulleted
      if (cell.constructor === Array) {
        var copy = '';
        for (var k=0; k < cell.length; k++)
          copy += '<li>'+cell[k]+'</li>';
        cell = '<ul>'+copy+'</ul>';
      }
        
      tableRow += '<td style="text-align: left; padding: 8px; border: 1px solid #ddd;">'+cell+'</td>';
    }
    
    // Make it an html table row
    table += '<tr>'+tableRow+'</tr>'; 
  }
  
  // Make header row
  var titleCaseHeaders = ToTitleCase(headers);
  var headerRow = '';
  for (var i=0; i < titleCaseHeaders.length; i++)
    headerRow += '<th style="text-align: left; padding: 8px; border: 1px solid #ddd;"><b>'+titleCaseHeaders[i]+'<b></th>';
  headerRow = '<tr>'+headerRow+'</tr>';
  
  // Add headers to table
  table = headerRow + table;
  
  // Make it a table
  table = '<table style="border: 1px solid black; border-collapse: collapse;">'+table+'</table>';
  
  return table;
}



// This function makes an HTML table for each user. Runs MakeHTMLTable() for each user in array
function MakeHTMLTables(allUsersArray) {
  Logger.log("array looks like = %s", JSON.stringify(allUsersArray));
  Logger.log('\n %s \n', allUsersArray.length);
  
  for (var i=0; i < allUsersArray.length; i++) {
    var headers = [
      'sheetName',
      'partNumber',
      'partName',
      'missingData'];
    
    allUsersArray[i]['table'] = MakeHTMLTable(headers, allUsersArray[i]['data']);
    Logger.log(allUsersArray[i]['table']);
  }
  
  Logger.log("array looks like = %s", JSON.stringify(allUsersArray));
  
  return allUsersArray;
}




// Get all valid sheet names (array of strings)
function GetSheetNames() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  // Get all PES sheet names
  var sheetNames = [];
  for (var i=0; i < sheets.length; i++)
    sheetNames.push(sheets[i].getSheetName());
  
  // Define sheets to exclude (in email dropdown)
  var invalid = [
    'Drop Down Options',
    'SVS Sponsors List',
    'Visualization'
  ];
  
  // Remove all invalid sheet Names from the sheetNames array
  sheetNames = sheetNames.filter(function(x) {
    return invalid.indexOf(x) < 0 
  });
  
  return sheetNames;
}




// Converts a camelCaseString to a Title Case String
function ToTitleCase(string) {
  
  // if the input is an array, run function on each element in array and return array  
  if (string.constructor === Array) {
    for (var i=0; i < string.length; i++) {
      var element = string[i].toString();
      
      string[i] = element
      // insert a space before all caps
      .replace(/([A-Z])/g, ' $1')
      // uppercase the first character
      .replace(/^./, function(str){ return str.toUpperCase(); })
    }
  }
  
  // if input is a single string, run the function
  else {
    string
    // insert a space before all caps
    .replace(/([A-Z])/g, ' $1')
    // uppercase the first character
    .replace(/^./, function(str){ return str.toUpperCase(); })
  }
  
  return string;
}




function GetColumns(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var values = sheet.getSheetValues(1, 1, 1, sheet.getLastColumn())
  var columns = [];
  
  // Convert to 1D array
  for (var i=0; i < values[0].length; i++)
    columns.push(values[0][i]);
  
  // Define sheets to exclude (in email dropdown)
  var invalid = [
    'Last Modified',
    'Edited by',
  ];
  
  // Remove all invalid values, according to the invalid array
  columns = columns.filter(function(x) {
    return invalid.indexOf(x) < 0 
  });
  
  return columns;
}
    
    
    
    
