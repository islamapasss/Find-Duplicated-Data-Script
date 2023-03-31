//Globals
var SHEET_URL = 'https://docs.google.com/spreadsheets/d/1laImmon0hqu1ph__IWDMIEowcIY4qry3eDFMFW2EUXQ/edit#gid=0';

var spreadSheet = SpreadsheetApp.openByUrl(SHEET_URL);
var sheet = spreadSheet.getSheets()[0];
 
// Creates an array with data from a chosen column
function readColumnData(column) {
  var lastRow = sheet.getLastRow();
  var columnRange = sheet.getRange(1, column, lastRow);
  var rangeArray = columnRange.getValues();
  // Convert to one dimensional array
  rangeArray = [].concat.apply([], rangeArray);
  return rangeArray;
}

// Creates an array with data from a chosen row
function readRowData(row) {
  var lastColumn = sheet.getLastColumn();
  var rowRange = sheet.getRange(row, 1, 1, lastColumn);
  var rangeArray = rowRange.getValues();
  // Convert to one dimensional array
  rangeArray = [].concat.apply([], rangeArray);
  Logger.log(rangeArray);
  return rangeArray;
}

// Sort data and find duplicates
function findDuplicates(data) {
  var sortedData = data.slice().sort();
  var duplicates = [];
  for (var i = 0; i < sortedData.length - 1; i++) {
    if (sortedData[i + 1] == sortedData[i]) {
      duplicates.push(sortedData[i]);
    }
  }
  return duplicates;
}

// Find locations of all duplicates
function getIndexes(data, duplicates) {
  var column = 2;
  var indexes = [];
  i = -1;
  // Loop through duplicates to find their indexes
  for (var n = 0; n < duplicates.length; n++) {
    while ((i = data.indexOf(duplicates[n], i + 1)) != -1) {
      indexes.push(i);
    }
  }
  return indexes;
}

// Highlight all instances of duplicate values in a column
function highlightColumnDuplicates(column, indexes) {
  for (n = 0; n < indexes.length; n++) {
    sheet.getRange(indexes[n] + 1, column).setBackground("gray");
  }
}

// Highlight all instances of duplicate values in a row
function highlightRowDuplicates(row, indexes) {
  for (n = 0; n < indexes.length; n++) {
    sheet.getRange(row, indexes[n] + 1).setBackground("gray");
  }
}

//----------- Main -------------
function columnMain(column) {
  var data = readColumnData(column);
  var duplicates = findDuplicates(data);
  var indexes = getIndexes(data, duplicates);
  highlightColumnDuplicates(column, indexes);
}

function rowMain(row) {
  var data = readRowData(row);
  var duplicates = findDuplicates(data);
  var indexes = getIndexes(data, duplicates);
  highlightRowDuplicates(row, indexes);
}

// ---------- Menu ----------
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Find Duplicates')
    .addItem('Select a Row', 'showRowPrompt')
    .addItem('Select a Column', 'showColumnPrompt')
    .addSeparator()
    .addSubMenu((ui.createMenu('Row or Column'))
      .addItem('Column', 'showColumnPrompt')
      .addItem('Row', 'showRowPrompt'))
    .addToUi();
}
// ---------- Prompt ----------
function showColumnPrompt() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Find Duplicates',
    'Enter letter of column to search:',
    ui.ButtonSet.OK_CANCEL);
  // Get user response, run main
  var button = response.getSelectedButton();
  var text = response.getResponseText();
  if (button == ui.Button.OK) {
    text = sheet.getRange(text + "1");
    text = text.getColumn();
    columnMain(text);
  }
}

function showRowPrompt() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Find Duplicates',
    'Enter number of row to search:',
    ui.ButtonSet.OK_CANCEL);
  // Get user response, run main
  var button = response.getSelectedButton();
  var text = response.getResponseText();
  if (button == ui.Button.OK) {
    rowMain(text);
  }
}



