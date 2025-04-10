/*
 * Version 0.02 made by yippym - 2025-04-10 01:05
 * https://github.com/Yippy/signal-tally-sheet
 */

/**
* Add Formula Exclusive Channel Signal History
*/
function addFormulaExclusiveChannelSignalHistory() {
  addFormulaBysignalHistoryName(SIGNAL_TALLY_EXCLUSIVE_CHANNEL_SIGNAL_SHEET_NAME);
}
/**
* Add Formula Stable Channel Signal History
*/
function addFormulaStableChannelSignalHistory() {
  addFormulaBysignalHistoryName(SIGNAL_TALLY_STABLE_SIGNAL_SHEET_NAME);
}
/**
* Add Formula W-Engine Channel Signal History
*/
function addFormulaWEngineChannelSignalHistory() {
  addFormulaBysignalHistoryName(SIGNAL_TALLY_W_ENGINES_SIGNAL_SHEET_NAME);
}
/**
* Add Formula Bangboo Channel Signal History
*/
function addFormulaBangbooChannelSignalHistory() {
  addFormulaBysignalHistoryName(SIGNAL_TALLY_BANGBOO_SIGNAL_SHEET_NAME);
}

/**
* Add Formula for selected Signal History sheet
*/
function addFormulaSignalHistory() {
  var sheetActive = SpreadsheetApp.getActiveSpreadsheet();
  var signalHistoryName = sheetActive.getSheetName();
  if (SIGNAL_TALLY_NAME_OF_SIGNAL_HISTORY.indexOf(signalHistoryName) != -1) {
    addFormulaBysignalHistoryName(signalHistoryName);
  } else {
    var message = 'Sheet must be called "' + SIGNAL_TALLY_EXCLUSIVE_CHANNEL_SIGNAL_SHEET_NAME + '" or "' + SIGNAL_TALLY_STABLE_SIGNAL_SHEET_NAME + '" or "' + SIGNAL_TALLY_W_ENGINES_SIGNAL_SHEET_NAME + '" or "' + SIGNAL_TALLY_BANGBOO_SIGNAL_SHEET_NAME + '"';
    var title = 'Invalid Sheet Name';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function addFormulaBysignalHistoryName(name, settingsSheet = null) {
  var sheetSource = getSourceDocument();
  if (sheetSource) {
    // Add Language
    var signalHistorySource;
    if (!settingsSheet) {
      settingsSheet = getSettingsSheet();
    }
    if (settingsSheet) {
      var languageFound = settingsSheet.getRange(2, 2).getValue();
      signalHistorySource = sheetSource.getSheetByName(SIGNAL_TALLY_SIGNAL_HISTORY_SHEET_NAME+"-"+languageFound);
    }
    if (signalHistorySource) {
      // Found language
    } else {
      // Default
      signalHistorySource = sheetSource.getSheetByName(SIGNAL_TALLY_SIGNAL_HISTORY_SHEET_NAME);
    }
    var sheet = findSignalHistoryByName(name,sheetSource);
    sheet.getRange("A1").setFontColor("white").setValue(name);
    var signalHistorySourceNumberOfColumn = signalHistorySource.getLastColumn();
    // Reduce two column due to paste and override
    var signalHistorySourceNumberOfColumnWithFormulas = signalHistorySourceNumberOfColumn - 2;

    var lastRowWithoutTitle = sheet.getMaxRows() - 1;

    var currentOverrideTitleCell = sheet.getRange(1, 2).getValue();
    var sourceOverrideTitleCell = signalHistorySource.getRange(1, 2).getValue();
    if (currentOverrideTitleCell != sourceOverrideTitleCell) {
      // If override column don't exist, populate from source
      var overrideCells = signalHistorySource.getRange(2, 2).getFormula();
      sheet.getRange(2, 2, lastRowWithoutTitle, 1).setValue(overrideCells);
      sheet.getRange(1, 2).setValue(sourceOverrideTitleCell);
      sheet.setColumnWidth(2, signalHistorySource.getColumnWidth(2));
    }
    
    // Get second row formula columns and set current sheet
    var formulaCells = signalHistorySource.getRange(2, 3, 1, signalHistorySourceNumberOfColumnWithFormulas).getFormulas();
    sheet.getRange(2, 3, lastRowWithoutTitle, signalHistorySourceNumberOfColumnWithFormulas).setValue(formulaCells);

    // Get title columns and set current sheet
    var titleCells = signalHistorySource.getRange(1, 3, 1, signalHistorySourceNumberOfColumnWithFormulas).getFormulas();
    sheet.getRange(1, 3, 1, signalHistorySourceNumberOfColumnWithFormulas).setValues(titleCells);

    for (var i = 3; i <= signalHistorySourceNumberOfColumn; i++) {
      // Apply formatting for cells
      var numberFormatCell = signalHistorySource.getRange(2, i).getNumberFormat();
      sheet.getRange(2, i, lastRowWithoutTitle, 1).setNumberFormat(numberFormatCell);
      // Set column width from source
      sheet.setColumnWidth(i, signalHistorySource.getColumnWidth(i));
    }

    // Ensure new row is not the same height as first, if row 2 did not exist
    sheet.autoResizeRows(2, 1);
  } else {
    var message = 'Unable to connect to source';
    var title = 'Error';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

/**
* Check is sheet exist in active spreadsheet, otherwise pull sheet from source
*/
function findSignalHistoryByName(name, sheetSource) {
  var signalHistorySheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (signalHistorySheet == null) {
    if (sheetSource == null) {
      sheetSource = getSourceDocument();
    }
    if (sheetSource) {
      var sheetCopySource = sheetSource.getSheetByName(SIGNAL_TALLY_SIGNAL_HISTORY_SHEET_NAME);
      sheetCopySource.copyTo(SpreadsheetApp.getActiveSpreadsheet()).setName(name);
      signalHistorySheet = SpreadsheetApp.getActive().getSheetByName(name);
      signalHistorySheet.showSheet();
    }
  }
  return signalHistorySheet;
}

/**
* Add sort for selected Signal History sheet
*/
function sortSignalHistory() {
  var sheetActive = SpreadsheetApp.getActiveSpreadsheet();
  var signalHistoryName = sheetActive.getSheetName();
  if (SIGNAL_TALLY_NAME_OF_SIGNAL_HISTORY.indexOf(signalHistoryName) != -1) {
    sortSignalHistoryByName(signalHistoryName);
  } else {
    var message = 'Sheet must be called "' + SIGNAL_TALLY_EXCLUSIVE_CHANNEL_SIGNAL_SHEET_NAME + '" or "' + SIGNAL_TALLY_STABLE_SIGNAL_SHEET_NAME + '" or "' + SIGNAL_TALLY_W_ENGINES_SIGNAL_SHEET_NAME + '" or "' + SIGNAL_TALLY_BANGBOO_SIGNAL_SHEET_NAME + '"';
    var title = 'Invalid Sheet Name';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

/**
* Sort Exclusive Channel Signal History
*/
function sortExclusiveChannelSignalHistory() {
  sortSignalHistoryByName(SIGNAL_TALLY_EXCLUSIVE_CHANNEL_SIGNAL_SHEET_NAME);
}

/**
* Sort Stable Channel Signal History
*/
function sortStableChannelSignalHistory() {
  sortSignalHistoryByName(SIGNAL_TALLY_STABLE_SIGNAL_SHEET_NAME);
}

/**
* Sort W-Engine Channel Signal History
*/
function sortWEngineChannelSignalHistory() {
  sortSignalHistoryByName(SIGNAL_TALLY_W_ENGINES_SIGNAL_SHEET_NAME);
}

/**
* Sort Bangboo Channel Signal History
*/
function sortBangbooChannelSignalHistory() {
  sortSignalHistoryByName(SIGNAL_TALLY_BANGBOO_SIGNAL_SHEET_NAME);
}

function sortSignalHistoryByName(sheetName) {
  var sheet = findSignalHistoryByName(sheetName, null);
  if (sheet) {
    if (sheet.getLastColumn() > 6) {
      var range = sheet.getRange(2, 1, sheet.getMaxRows()-1, sheet.getLastColumn());
      range.sort([{column: 5, ascending: true}, {column: 2, ascending: true}, {column: 7, ascending: true}]);
    } else {
      var message = 'Invalid number of columns to sort, run "Refresh Formula" or "Update Items"';
      var title = 'Error';
      SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
    }
  } else {
    var message = 'Unable to connect to source';
    var title = 'Error';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}