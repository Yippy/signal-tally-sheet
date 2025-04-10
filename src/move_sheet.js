/*
 * Version 0.02 made by yippym - 2025-04-10 01:05
 * https://github.com/Yippy/signal-tally-sheet
 */
function moveToSettingsSheet() {
  moveToSheetByName(SIGNAL_TALLY_SETTINGS_SHEET_NAME);
}

function moveToDashboardSheet() {
  moveToSheetByName(SIGNAL_TALLY_DASHBOARD_SHEET_NAME);
}

function moveToExclusiveChannelSignalHistorySheet() {
  moveToSheetByName(SIGNAL_TALLY_EXCLUSIVE_CHANNEL_SIGNAL_SHEET_NAME);
}

function moveToStableChannelSignalHistorySheet() {
  moveToSheetByName(SIGNAL_TALLY_STABLE_SIGNAL_SHEET_NAME);
}

function moveToWEngineChannelSignalHistorySheet() {
  moveToSheetByName(SIGNAL_TALLY_W_ENGINES_SIGNAL_SHEET_NAME);
}

function moveToBangbooChannelSignalHistorySheet() {
  moveToSheetByName(SIGNAL_TALLY_BANGBOO_SIGNAL_SHEET_NAME);
}

function moveToChangelogSheet() {
  moveToSheetByName(SIGNAL_TALLY_CHANGELOG_SHEET_NAME);
}

function moveToPityCheckerSheet() {
  moveToSheetByName(SIGNAL_TALLY_PITY_CHECKER_SHEET_NAME);
}

function moveToEventsSheet() {
  moveToSheetByName(SIGNAL_TALLY_EVENTS_SHEET_NAME);
}

function moveToAgentsSheet() {
  moveToSheetByName(SIGNAL_TALLY_AGENTS_SHEET_NAME);
}

function moveToBangboosSheet() {
  moveToSheetByName(SIGNAL_TALLY_BANGBOOS_SHEET_NAME);
}

function moveToWEnginesSheet() {
  moveToSheetByName(SIGNAL_TALLY_W_ENGINES_SHEET_NAME);
}

function moveToResultsSheet() {
  moveToSheetByName(SIGNAL_TALLY_RESULTS_SHEET_NAME);
}

function moveToReadmeSheet() {
  moveToSheetByName(SIGNAL_TALLY_README_SHEET_NAME);
}

function moveToMonochromeCalculatorSheet() {
  moveToSheetByName(SIGNAL_TALLY_MONOCHROME_CALCULATOR_SHEET_NAME);
}

function moveToSheetByName(nameOfSheet) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(nameOfSheet);
  if (sheet) {
    sheet.activate();
  } else {
    var settingsForOptionalSheet = SETTINGS_FOR_OPTIONAL_SHEET[nameOfSheet];
    if (settingsForOptionalSheet) {
      var settingsSheet = SpreadsheetApp.getActive().getSheetByName(SIGNAL_TALLY_SETTINGS_SHEET_NAME);
      if (settingsSheet) {
        var settingOption = settingsForOptionalSheet["setting_option"];
        if (!settingsSheet.getRange(settingOption).getValue()) {
          displayUserAlert("Optional Sheet", nameOfSheet+" has been disabled within Settings, enable this sheet at cell '"+settingOption+"', and run 'Update Items'",  SpreadsheetApp.getUi().ButtonSet.OK)
        }
      }
    }
    title = "Error";
    message = "Unable to find sheet named '"+nameOfSheet+"'.";
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}