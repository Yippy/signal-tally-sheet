/*
 * Version 0.02 made by yippym - 2025-04-10 01:05
 * https://github.com/Yippy/signal-tally-sheet
 */
function importButtonScript() {
  var settingsSheet = getSettingsSheet();
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName(SIGNAL_TALLY_DASHBOARD_SHEET_NAME);
  if (!dashboardSheet || !settingsSheet) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Unable to find '" + SIGNAL_TALLY_DASHBOARD_SHEET_NAME + "' or '" + SIGNAL_TALLY_SETTINGS_SHEET_NAME + "'", "Missing Sheets");
    return;
  }

  var userImportSelection = dashboardSheet.getRange(dashboardEditRange[4]).getValue();
  var autoImportSelection = dashboardSheet.getRange(dashboardEditRange[5]).getValue();
  var importSelectionText = dashboardSheet.getRange(dashboardEditRange[6]).getValue();
  var importSelectionTextSubtitle = dashboardSheet.getRange(dashboardEditRange[7]).getValue();
  var urlInput = null;
  if (importSelectionText === autoImportSelection) {
    urlInput = getCachedAuthKeyInput();
    var isInfoRetrieved = false;
    var sheetRedirectSource = SpreadsheetApp.openById(SIGNAL_TALLY_SHEET_SOURCE_REDIRECT_ID);
    // Check redirect source is available
    if (sheetRedirectSource) {
      // attempt to load latest message for auto import, as Genshin Impact can sometimes change method.
      var sheetAutoImportSource = sheetRedirectSource.getSheetByName(SIGNAL_TALLY_REDIRECT_SOURCE_AUTO_IMPORT_SHEET_NAME);
      if (sheetAutoImportSource) {
        importSelectionTextSubtitle = sheetAutoImportSource.getRange("B5").getValue();
        isInfoRetrieved = true;
      }
    }
    if (!isInfoRetrieved) {
      // Resort to default when blurp has not been retrieved
      importSelectionTextSubtitle = "Please note Feedback URL no longer works for Auto Import\n\nPress [YES] to continue\nPress [NO] to visit tutorial";
    }
  }

  if (urlInput === null) {
    const result = displayUserPrompt(importSelectionText, importSelectionTextSubtitle, SpreadsheetApp.getUi().ButtonSet.YES_NO_CANCEL);
    var button = result.getSelectedButton();
    if (button !== SpreadsheetApp.getUi().Button.YES) {
      if (button == SpreadsheetApp.getUi().Button.NO) {
        displayAutoImport();
      }
      return;
    }
    urlInput = result.getResponseText();

    if (importSelectionText === autoImportSelection) {
      setCachedAuthKeyInput(urlInput);
    }
  }

  if (userImportSelection === importSelectionText) {
    settingsSheet.getRange("D6").setValue(urlInput);
    importDataManagement();
  } else {
    importFromAPI(urlInput);
  }
}

function importDataManagement() {
  var settingsSheet = getSettingsSheet();
  var userImportInput = settingsSheet.getRange("D6").getValue();
  var userImportStatus = settingsSheet.getRange("E7").getValue();
  var message = "";
  var title = "";
  var statusMessage = "";
  var rowOfStatusSignalHistory = 9;
  if (userImportStatus == IMPORT_STATUS_COMPLETE) {
      title = "Error";
      message = "Already done, you only need to run once";
  } else {
    if (userImportInput) {
      // Attempt to load as URL
      var importSource;
      try {
        importSource = SpreadsheetApp.openByUrl(userImportInput);
      } catch(e) {
      }
      if (importSource) {
      } else {
        // Attempt to load as ID instead
        try {
          importSource = SpreadsheetApp.openById(userImportInput);
        } catch(e) {
        }
      }
      if (importSource) {
        // Go through the available sheet list
        for (var i = 0; i < SIGNAL_TALLY_NAME_OF_SIGNAL_HISTORY.length; i++) {
          var bannerImportSheet = importSource.getSheetByName(SIGNAL_TALLY_NAME_OF_SIGNAL_HISTORY[i]);

          if (bannerImportSheet) {
            var numberOfRows = bannerImportSheet.getMaxRows()-1;
            var range = bannerImportSheet.getRange(2, 1, numberOfRows, 2);

            if (numberOfRows > 0) {
              var bannerSheet = SpreadsheetApp.getActive().getSheetByName(SIGNAL_TALLY_NAME_OF_SIGNAL_HISTORY[i]);

              if (bannerSheet) {
                bannerSheet.getRange(2, 1, numberOfRows, 2).setValues(range.getValues());
                settingsSheet.getRange(rowOfStatusSignalHistory+i, 5).setValue(IMPORT_STATUS_WISH_HISTORY_COMPLETE);
              } else {
                settingsSheet.getRange(rowOfStatusSignalHistory+i, 5).setValue(IMPORT_STATUS_WISH_HISTORY_NOT_FOUND);
              }
            } else {
              settingsSheet.getRange(rowOfStatusSignalHistory+i, 5).setValue(IMPORT_STATUS_WISH_HISTORY_EMPTY);
            }
          } else {
            settingsSheet.getRange(rowOfStatusSignalHistory+i, 5).setValue(IMPORT_STATUS_WISH_HISTORY_NOT_FOUND);
          }
        }
        var sourceSettingsSheet = importSource.getSheetByName(SIGNAL_TALLY_SETTINGS_SHEET_NAME);
        if (sourceSettingsSheet) {
          var sourcePityCheckerSheet = importSource.getSheetByName(SIGNAL_TALLY_PITY_CHECKER_SHEET_NAME);
          if (sourcePityCheckerSheet) {
            savePityCheckerSettings(sourcePityCheckerSheet, settingsSheet);
          }
          if (sourceSettingsSheet.getMaxColumns() >= 8) {
            var version = sourceSettingsSheet.getRange("H1").getValue();
            if (version == "2.7") {
              var pityCheckerIsShow4Star = sourceSettingsSheet.getRange("B18").getValue();
              settingsSheet.getRange("B18").setValue(pityCheckerIsShow4Star == true);
              var pityCheckerIsShow5Star = sourceSettingsSheet.getRange("B19").getValue();
              settingsSheet.getRange("B19").setValue(pityCheckerIsShow5Star == true);
            }
          }
          var pityCheckerSheet = SpreadsheetApp.getActive().getSheetByName(SIGNAL_TALLY_PITY_CHECKER_SHEET_NAME);
          if (pityCheckerSheet) {
            restorePityCheckerSettings(pityCheckerSheet, settingsSheet);
          }
          var offset = sourceSettingsSheet.getRange("B10").getValue();
          if (offset >= -11 && offset <= 12) {
             settingsSheet.getRange("B10").setValue(offset);
          }
          var language = sourceSettingsSheet.getRange("B2").getValue();
          if (language) {
             settingsSheet.getRange("B2").setValue(language);
          }
          var server = sourceSettingsSheet.getRange("B3").getValue();
          if (server) {
             settingsSheet.getRange("B3").setValue(server);
          }
        }
        //Restore Events
        var sourceEventsSheet = importSource.getSheetByName(SIGNAL_TALLY_EVENTS_SHEET_NAME);
        if (sourceEventsSheet) {
          saveEventsSettings(sourceEventsSheet,settingsSheet);
          var eventsSheet = SpreadsheetApp.getActive().getSheetByName(SIGNAL_TALLY_EVENTS_SHEET_NAME);
          if (eventsSheet) {
            restoreEventsSettings(eventsSheet, settingsSheet)
          }
        }
        
        //Restore Results
        var sourceResultsSheet = importSource.getSheetByName(SIGNAL_TALLY_RESULTS_SHEET_NAME);
        if (sourceResultsSheet) {
          saveResultsSettings(sourceResultsSheet, settingsSheet);
          var resultsSheet = SpreadsheetApp.getActive().getSheetByName(SIGNAL_TALLY_RESULTS_SHEET_NAME);
          if (resultsSheet) {
            restoreResultsSettings(resultsSheet, settingsSheet)
          }
        }
        
        //Restore Agents
        var sourceConstellationSheet = importSource.getSheetByName(SIGNAL_TALLY_AGENTS_SHEET_NAME);
        if (sourceConstellationSheet) {
          saveCollectionSettings(sourceConstellationSheet, settingsSheet,"G7","H7");
          var constellationSheet = SpreadsheetApp.getActive().getSheetByName(SIGNAL_TALLY_AGENTS_SHEET_NAME);
          if (constellationSheet) {
            restoreCollectionSettings(constellationSheet, settingsSheet,"G7","H7");
          }
        }

        //Restore W-Engines
        var sourceWEnginesSheet = importSource.getSheetByName(SIGNAL_TALLY_W_ENGINES_SHEET_NAME);
        if (sourceWEnginesSheet) {
          saveCollectionSettings(sourceWEnginesSheet, settingsSheet,"G8","H8");
          var WEnginesSheet = SpreadsheetApp.getActive().getSheetByName(SIGNAL_TALLY_W_ENGINES_SHEET_NAME);
          if (WEnginesSheet) {
            restoreCollectionSettings(WEnginesSheet, settingsSheet,"G8","H8");
          }
        }

        //Restore Bangboos
        var sourceBangboosSheet = importSource.getSheetByName(SIGNAL_TALLY_BANGBOOS_SHEET_NAME);
        if (sourceBangboosSheet) {
          saveCollectionSettings(sourceBangboosSheet, settingsSheet,"G11","H11");
          var BangboosSheet = SpreadsheetApp.getActive().getSheetByName(SIGNAL_TALLY_W_ENGINES_SHEET_NAME);
          if (BangboosSheet) {
            restoreCollectionSettings(BangboosSheet, settingsSheet,"G11","H11");
          }
        }

        title = "Complete";
        message = "Imported all rows in column Paste Value and Override";
        statusMessage = IMPORT_STATUS_COMPLETE;
      } else {
        title = "Error";
        message = "Import From URL or Spreadsheet ID is invalid";
        statusMessage = IMPORT_STATUS_FAILED;
      }
    } else {
      title = "Error";
      message = "Import From URL or Spreadsheet ID is empty";
      statusMessage = IMPORT_STATUS_FAILED;
    }

    settingsSheet.getRange("E7").setValue(statusMessage);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
}