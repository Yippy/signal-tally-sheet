/*
 * Version 0.02 made by yippym - 2025-04-10 01:05
 * https://github.com/Yippy/signal-tally-sheet
 */
function getSourceDocument() {
  // Due to the nature of the document, when new contents is being added to the source. It would be disabled from access, which this function will try and load a message or backup document for the user.
  var sheetRedirectSource = SpreadsheetApp.openById(SIGNAL_TALLY_SHEET_SOURCE_REDIRECT_ID);
  var isSourceAvailable = sheetRedirectSource.getRange("B6").getValue();
  var sheetSource = null;
  if (isSourceAvailable == 'YES') {
    var sheetSourceId = sheetRedirectSource.getRange("B8").getValue();
    sheetSource = SpreadsheetApp.openById(sheetSourceId);
  } else {
    var isBackupAvailable = sheetRedirectSource.getRange("F6").getValue();
    if (isBackupAvailable == 'YES') {
      var sheetBackupId = sheetRedirectSource.getRange("F8").getValue();
      sheetSource = SpreadsheetApp.openById(sheetBackupId);
      var showBackupMessage = sheetRedirectSource.getRange("F10").getValue();
      if (showBackupMessage == 'YES') {
        displayBackup();
      }
    } else {
      displayMaintenance();
    }
  }
  return sheetSource;
}