function onOpen() {
  ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('CS 192 Menu')
      .addItem('Fetch data from OJ', 'fetchData')
      .addSeparator()
      .addItem('Group by verdict', 'groupByVerdict')
      .addSeparator()
      .addItem('Clear results', 'clearResults')
      .addSeparator()
      .addItem('Remove temporary sheets', 'removeTemporarySheets')
      .addSeparator()
      .addToUi();
}

function deleteExtraSheets() {
    sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for (i = 0; i < sheets.length ; i++ ) {
      sheet = sheets[i];

      if (sheet.getSheetName() != "Main Sheet" && sheet.getSheetName()  != "Results") {
        SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
      }
    }
}

/**
 * Wrapper for Spreadsheet.insertSheet() method to support customization.
 * All parameters are optional & positional.
 *
 * @param {String}  sheetName     Name of new sheet (defaults to "Sheet #")
 * @param {Number}  sheetIndex    Position for new sheet (default 0 means "end")
 * @param {Number}  rows          Vertical dimension of new sheet (default 0 means "system default", 1000)
 * @param {Number}  columns       Horizontal dimension of new sheet (default 0 means "system default", 26)
 * @param {String}  template      Name of existing sheet to copy (default "" means none)
 *
 * @returns {Sheet}               Sheet object for chaining.
 */
function insertSheet( sheetName, sheetIndex, rows, columns, template ) {
  
  // Check parameters, set defaults
  ss = SpreadsheetApp.getActive();
  numSheets = ss.getSheets().length;
  sheetIndex = sheetIndex || (numSheets + 1);
  sheetName = sheetName || "Sheet " + sheetIndex;
  options = template ? { 'template' : ss.getSheetByName(template) } : {};
  
  // Will throw an exception if sheetName already exists
  newSheet = ss.insertSheet(sheetName, sheetIndex, options);
 
  if (rows !== 0) {
    // Adjust dimension: rows
    newSheetRows = newSheet.getMaxRows();
    
    if (rows < newSheetRows) {
      // trim rows
      newSheet.deleteRows(rows+1, newSheetRows-rows);
    }
    else if (rows > newSheetRows) {
      // add rows
      newSheet.insertRowsAfter(newSheetRows, rows-newSheetRows);
    }
  }
  
  if (columns !== 0) {
    // Adjust dimension: columns
    newSheetColumns = newSheet.getMaxColumns();
    
    if (columns < newSheetColumns) {
      // trim rows
      newSheet.deleteColumns(columns+1, newSheetColumns-columns);
    }
    else if (columns > newSheetColumns) {
      // add rows
      newSheet.insertColumnsAfter(newSheetColumns,columns-newSheetColumns);
    }
  }
  
  // Return new Sheet object
  return newSheet;
}

function clearResultsNoPrompt(){
  resultsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results");
  resultsSheet.activate();

  numRows = resultsSheet.getLastRow();
  Logger.log("Number of Rows:" + numRows);
  numCols = resultsSheet.getLastColumn();
  Logger.log("Number of Cols:" + numCols);

  if (numRows >= 2 && resultsSheet.getRange("A1:F2").getDisplayValues() != [['Date','Problem','User','Language','Results','Points'],['','','','','','']]) {
    resultsSheetDataRange = resultsSheet.getRange(3,1,numRows - 2, numCols);
    resultsSheet.deleteRows(2,numRows - 2);

    resultsSheetDataRange = resultsSheet.getRange("A2:F2");
    resultsSheetDataRange.setValues([['','','','','','']]);
  }
}

function clearResults() {
  // Display a dialog box with a title, message, and "Yes" and "No" buttons. The user can also
  // close the dialog by clicking the close button in its title bar.
  ui = SpreadsheetApp.getUi();
  response = ui.alert('Confirm', 'Are you sure you want to continue?', ui.ButtonSet.YES_NO);

// Process the user's response.
  if (response == ui.Button.YES) {
    Logger.log('The user clicked "Yes."');

    clearResultsNoPrompt();

  } else if (response == ui.Button.NO) {
    Logger.log('The user clicked "No"');
  } else {
    Logger.log('The user clicked the close button in the dialog\'s title bar.');
  }
}

function removeTemporarySheets() {
  ui = SpreadsheetApp.getUi();
  response = ui.alert('Remove temporary sheets', 'Are you sure you want to remove all temporary sheets?', ui.ButtonSet.YES_NO);

// Process the user's response.
  if (response == ui.Button.YES) {
    Logger.log('The user clicked "Yes."');
    deleteExtraSheets();
  } else if (response == ui.Button.NO) {
    Logger.log('The user clicked "No"');
  } else {
    Logger.log('The user clicked the close button in the dialog\'s title bar.');
  }

}

function fetchData() {
  clearResultsNoPrompt();

  mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Main Sheet");
  resultsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results");
  mainSheet.activate();

  // get values for API Token, Username and language all at once

  values = SpreadsheetApp.getActiveSheet().getRange(1, 2, 3, 1).getDisplayValues();
  Logger.log("APIToken: " + values[0][0]);

  url = 'https://oj.dcs.upd.edu.ph/api/v2/submissions';

  if (values[1][0] != '' && values[2][0] != '') {
    url = url + '?user=' + values[1][0] + '&language=' + values[2][0];
  }
  else if (values[1][0] != '') {
    url = url + '?user=' + values[1][0];
  }
  else {
    url = url + '?language=' + values[2][0];
  }
  
  headers = {
    "Authorization" : "Bearer "  + values[0][0]
  };

  params = {
    "method":"GET",
    "headers":headers
  };

  if (values[0][0] != '') {
    response = UrlFetchApp.fetch(url, params);
    Logger.log("Response: " + response);
    Logger.log("URL: " + url);
    Logger.log("Authorization given");
  }
  else {
    response = UrlFetchApp.fetch(url);
    Logger.log("Response: " + response);
    Logger.log("URL: " + url);
  }

  json = response.getContentText();
  results = JSON.parse(json);
  Logger.log("Results: " + results);

  resultCount = results.data.total_objects;
  Logger.log("result count: " + resultCount);

  if (resultCount > 0) {
    resultsSheet.activate();
  }

    // Show a 3-second popup with the title "Status" and the message "Task started".
  SpreadsheetApp.getActiveSpreadsheet().toast('Results found: ' + resultCount , 'Results', 3);

  resultsSheet = SpreadsheetApp.getActiveSpreadsheet().insertRowsAfter(2,resultCount-1);
  
  newResultsSheet = [];

  for(row = 0; row < resultCount; row++){
    newResultsSheet.push([results.data.objects[row].date, results.data.objects[row].problem, results.data.objects[row].user, results.data.objects[row].language, results.data.objects[row].result, results.data.objects[row].points ]);
  }
  
  Logger.log("Rows: " + newResultsSheet.length);
  Logger.log("Columns: " + newResultsSheet[0].length);

  resultsSheetDataRange = resultsSheet.getRange(2,1,resultCount,6);
  resultsSheetDataRange.setValues(newResultsSheet);
}

function groupByVerdict() {
  groupCount = 0;

  resultsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Results");

  numRows = resultsSheet.getLastRow();
  Logger.log("Number of Rows:" + numRows);
  numCols = resultsSheet.getLastColumn();
  Logger.log("Number of Cols:" + numCols);

  if (numRows >= 2 && resultsSheet.getRange("A1:F2").getDisplayValues() != [['Date','Problem','User','Language','Results','Points'],['','','','','','']]) {
    resultsSheet.activate();

    deleteExtraSheets();

    numRows = resultsSheet.getLastRow();
    verdictTypes = [];
    groupByVerdictValues = [];

    sourceRange = resultsSheet.getRange(2, 1, numRows-1, 6);

    resultsSheetValues = sourceRange.getValues();
      Logger.log("results: " + resultsSheetValues);

    for (currRow = 0; currRow < numRows-1; currRow++) {
      verdict = resultsSheetValues[currRow][4];
      Logger.log(verdict);
      if (verdictTypes.indexOf(verdict) == -1) {
        groupCount++;
        Logger.log(verdictTypes.indexOf(verdict));
        Logger.log(verdictTypes);
        verdictTypes.push(verdict);
        groupByVerdictValues.push(verdict);

        groupByVerdictValues.push([resultsSheetValues[currRow]]);

      }
      else {
        groupByVerdictValues[groupByVerdictValues.indexOf(verdict) + 1].push(resultsSheetValues[currRow]);
      }
    }

    for (i = 0; i < groupCount ; i++ ) {
      newSheet = insertSheet(verdictTypes[i], 3, groupByVerdictValues[groupByVerdictValues.indexOf(verdictTypes[i]) + 1].length + 1, 6, resultsSheet);
      newSheet.setColumnWidth(1,resultsSheet.getColumnWidth(1));
      newSheet.setColumnWidth(2,resultsSheet.getColumnWidth(2));
      newSheet.setColumnWidth(3,resultsSheet.getColumnWidth(3));
      newSheet.setColumnWidth(4,resultsSheet.getColumnWidth(4));
      newSheet.setColumnWidth(5,resultsSheet.getColumnWidth(5));
      newSheet.setColumnWidth(6,resultsSheet.getColumnWidth(6));

      sourceHeaderRange = resultsSheet.getRange("A1:F2");
      targetHeaderRange = newSheet.getRange("A1:F2");

      targetHeaderRange.setValues([['Date','Problem','User','Language','Results','Points'],['','','','','','']]);
      
      sourceRange = resultsSheet.getRange("A1:F"+(groupByVerdictValues[groupByVerdictValues.indexOf(verdictTypes[i]) + 1].length + 1));
      targetRange = newSheet.getRange("A1:F"+(groupByVerdictValues[groupByVerdictValues.indexOf(verdictTypes[i]) + 1].length + 1));

      sourceRange.copyTo(targetRange, {formatOnly:true})
      
      Logger.log("verdict: " + verdictTypes[i] + " length: " + groupByVerdictValues[groupByVerdictValues.indexOf(verdictTypes[i]) + 1].length);

      targetRange = newSheet.getRange("A2:F"+(groupByVerdictValues[groupByVerdictValues.indexOf(verdictTypes[i]) + 1].length + 1));
      Logger.log(groupByVerdictValues[groupByVerdictValues.indexOf(verdictTypes[i]) + 1]);
      targetRange.setValues(groupByVerdictValues[groupByVerdictValues.indexOf(verdictTypes[i]) + 1]);
    }
    SpreadsheetApp.getActiveSpreadsheet().toast('Groups created: ' + groupCount , 'Grouping', 3);
  }

}
