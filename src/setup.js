/*
 * Version 0.01 made by yippym - 2024-12-18 01:05
 * https://github.com/Yippy/signal-tally-sheet
 */
function onInstall(e) {
  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    generateInitialiseToolbar();
  } else {
    onOpen(e);
  }
}

function onOpen(e) {
  if (e && e.authMode == ScriptApp.AuthMode.NONE) {
    generateInitialiseToolbar();
  } else {
    var settingsSheet = SpreadsheetApp.getActive().getSheetByName(SIGNAL_TALLY_SETTINGS_SHEET_NAME);
    if (!settingsSheet) {
      generateInitialiseToolbar();
    } else {
      getDefaultMenu();
    }
    checkLocaleIsSetCorrectly();
  }
}

function generateInitialiseToolbar() {
  var menu = SIGNAL_TALLY_SHEET_SCRIPT_IS_ADD_ON ? SpreadsheetApp.getUi().createAddonMenu() : SpreadsheetApp.getUi().createMenu(SIGNAL_TALLY_SHEET_TOOLBAR_NAME);
  menu
    .addItem('Initialise', 'updateItemsList')
    .addToUi();
}

function displayUserPrompt(titlePrompt, messagePrompt, buttonSet) {
  const ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    titlePrompt,
    messagePrompt,
    buttonSet);
  return result;
}

function displayUserAlert(titleAlert, messageAlert, buttonSet) {
  const ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    titleAlert,
    messageAlert,
    buttonSet);
  return result;
}

/* Ensure Sheets is set to the supported locale due to source document formula */
function checkLocaleIsSetCorrectly() {
  var currentLocale = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale();
  if (currentLocale != SIGNAL_TALLY_SHEET_SUPPORTED_LOCALE) {
    SpreadsheetApp.getActiveSpreadsheet().setSpreadsheetLocale(SIGNAL_TALLY_SHEET_SUPPORTED_LOCALE);
    var message = 'To ensure compatibility with formula from source document, your locale "'+currentLocale+'" has been changed to "'+SIGNAL_TALLY_SHEET_SUPPORTED_LOCALE+'"';
    var title = 'Sheets Locale Changed';
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

function getDefaultMenu() {
  var menu = SIGNAL_TALLY_SHEET_SCRIPT_IS_ADD_ON ? SpreadsheetApp.getUi().createAddonMenu() : SpreadsheetApp.getUi().createMenu(SIGNAL_TALLY_SHEET_TOOLBAR_NAME);
  var ui = SpreadsheetApp.getUi();
  menu
  .addSeparator()
  .addSubMenu(ui.createMenu('Exclusive Channel Signal History')
            .addItem('Sort Range', 'sortExclusiveChannelSignalHistory')
            .addItem('Refresh Formula', 'addFormulaExclusiveChannelSignalHistory'))
  .addSubMenu(ui.createMenu('Stable Channel Signal History')
            .addItem('Sort Range', 'sortStableChannelSignalHistory')
            .addItem('Refresh Formula', 'addFormulaStableChannelSignalHistory'))
  .addSubMenu(ui.createMenu('W-Engine Channel Signal History')
            .addItem('Sort Range', 'sortWEngineChannelSignalHistory')
            .addItem('Refresh Formula', 'addFormulaWEngineChannelSignalHistory'))
  .addSubMenu(ui.createMenu('Bangboo Channel Signal History')
            .addItem('Sort Range', 'sortBangbooChannelSignalHistory')
            .addItem('Refresh Formula', 'addFormulaBangbooChannelSignalHistory'))
  .addSeparator()
  .addSubMenu(ui.createMenu('Data Management')
            .addItem('Import', 'importDataManagement')
            .addSeparator()
            .addItem('Set Schedule', 'setTriggerDataManagement')
            .addItem('Remove All Schedule', 'removeTriggerDataManagement')
            .addSeparator()
            .addItem('Auto Import', 'importFromAPI')
            .addItem('Tutorial', 'displayAutoImport')
            )
  .addSeparator()
  .addItem('Quick Update', 'quickUpdate')
  .addItem('Update Items', 'updateItemsList')
  .addItem('Get Latest README', 'displayReadme')
  .addItem('About', 'displayAbout')
  .addToUi();
}

function getSettingsSheet() {
  var settingsSheet = SpreadsheetApp.getActive().getSheetByName(SIGNAL_TALLY_SETTINGS_SHEET_NAME);
  var sheetSource;
  if (!settingsSheet) {
    sheetSource = getSourceDocument();
    var sheetSettingSource = sheetSource.getSheetByName(SIGNAL_TALLY_SETTINGS_SHEET_NAME);
    if (sheetSettingSource) {
      settingsSheet = sheetSettingSource.copyTo(SpreadsheetApp.getActiveSpreadsheet());
      settingsSheet.setName(SIGNAL_TALLY_SETTINGS_SHEET_NAME);
      getDefaultMenu();
    }
  } else {
    settingsSheet.getRange("H1").setValue(SIGNAL_TALLY_SHEET_SCRIPT_VERSION);
  }
  var dashboardSheet = SpreadsheetApp.getActive().getSheetByName(SIGNAL_TALLY_DASHBOARD_SHEET_NAME);
  if (!dashboardSheet) {
    if (!sheetSource) {
      sheetSource = getSourceDocument();
    }
    var sheetDashboardSource = sheetSource.getSheetByName(SIGNAL_TALLY_DASHBOARD_SHEET_NAME);
    if (sheetDashboardSource) {
      dashboardSheet = sheetDashboardSource.copyTo(SpreadsheetApp.getActiveSpreadsheet());
      dashboardSheet.setName(SIGNAL_TALLY_DASHBOARD_SHEET_NAME);
    }
  }
  if (dashboardSheet) {
    if (SIGNAL_TALLY_SHEET_SCRIPT_IS_ADD_ON) {
      dashboardSheet.getRange(dashboardEditRange[8]).setFontColor("green").setFontWeight("bold").setHorizontalAlignment("left").setValue("Add-On Enabled");
    } else {
      dashboardSheet.getRange(dashboardEditRange[8]).setFontColor("white").setFontWeight("bold").setHorizontalAlignment("left").setValue("Embedded Script");
    }
  }
  return settingsSheet;
}