/*
 * Version 0.01 made by yippym - 2024-12-18 01:05
 * https://github.com/Yippy/signal-tally-sheet
 */
function extractAuthKeyFromInput(userInput) {
  urlForAPI = userInput.toString().split("&");
  var foundAuth = "";
  for (var i = 0; i < urlForAPI.length; i++) {
    var queryString = urlForAPI[i].toString().split("=");
    if (queryString.length == 2) {
      if (queryString[0] == "authkey") {
        foundAuth = queryString[1];
        break;
      }
    }
  }
  return foundAuth;
}

function testAuthKeyInputValidity(userInput) {
  var authKey = extractAuthKeyFromInput(userInput);
  if (authKey == "") {
    return false;
  }

  const USING_BANNER = "Stable Channel Signal History";

  var settingsSheet = getSettingsSheet();
  var queryBannerCode = AUTO_IMPORT_BANNER_SETTINGS_FOR_IMPORT[USING_BANNER].gacha_type;
  var selectedServer = settingsSheet.getRange("B3").getValue();
  var languageSettings = AUTO_IMPORT_LANGUAGE_SETTINGS_FOR_IMPORT[settingsSheet.getRange("B2").getValue()];
  if (languageSettings == null) {
    // Get default language
    languageSettings = AUTO_IMPORT_LANGUAGE_SETTINGS_FOR_IMPORT["English"];
  }
  var url = getSignalHistoryUrl(selectedServer, queryBannerCode, languageSettings, 1, authKey);
  responseJson = JSON.parse(UrlFetchApp.fetch(url).getContentText());
  if (responseJson.retcode === 0) {
    return true;
  }
  return false;
}

function getCachedAuthKeyProperty() {
  return  "cachedAuthKey_" + SpreadsheetApp.getActiveSpreadsheet().getId();
}

function invalidateCachedAuthKey() {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty(getCachedAuthKeyProperty());
}

function setCachedAuthKeyInput(userInput) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(getCachedAuthKeyProperty(), JSON.stringify({ userInput, timeOfInput: new Date() }));
}

function getCachedAuthKeyInput() {
  const userProperties = PropertiesService.getUserProperties();
  const cachedAuthKey = JSON.parse(userProperties.getProperty(getCachedAuthKeyProperty()));

  if (cachedAuthKey == null) {
    return null;
  }

  const timeOfInput = new Date(cachedAuthKey.timeOfInput);
  const timeDiff = new Date().getTime() - timeOfInput.getTime();
  if (timeDiff > CACHED_AUTHKEY_TIMEOUT) {
    invalidateCachedAuthKey();
    return null;
  }

  if (!testAuthKeyInputValidity(cachedAuthKey.userInput)) {
    invalidateCachedAuthKey();
    return null;
  }

  return cachedAuthKey.userInput;
}


function importFromAPI(urlForAPI) {
  var settingsSheet = getSettingsSheet();
  settingsSheet.getRange("E42").setValue(new Date());
  settingsSheet.getRange("E43").setValue("");

  if (AUTO_IMPORT_URL_FOR_API_BYPASS != "") {
    urlForAPI = AUTO_IMPORT_URL_FOR_API_BYPASS;
  }
  var authKey = extractAuthKeyFromInput(urlForAPI);
  var bannerName;
  var bannerSheet;
  var bannerSettings;
  if (authKey == "") {
    // Display auth key not available
    for (var i = 0; i < SIGNAL_TALLY_NAME_OF_SIGNAL_HISTORY.length; i++) {
      bannerName = SIGNAL_TALLY_NAME_OF_SIGNAL_HISTORY[i];
      bannerSettings = AUTO_IMPORT_BANNER_SETTINGS_FOR_IMPORT[bannerName];
      bannerSettings.setStatusText("No auth key", settingsSheet);
    }
  } else {
    var selectedLanguageCode = settingsSheet.getRange("B2").getValue();
    var selectedServer = settingsSheet.getRange("B3").getValue();
    var languageSettings = AUTO_IMPORT_LANGUAGE_SETTINGS_FOR_IMPORT[selectedLanguageCode];
    if (languageSettings == null) {
      // Get default language
      languageSettings = AUTO_IMPORT_LANGUAGE_SETTINGS_FOR_IMPORT["English"];
    }
    // Clear status
    for (var i = 0; i < SIGNAL_TALLY_NAME_OF_SIGNAL_HISTORY.length; i++) {
      bannerName = SIGNAL_TALLY_NAME_OF_SIGNAL_HISTORY[i];
      bannerSettings = AUTO_IMPORT_BANNER_SETTINGS_FOR_IMPORT[bannerName];
      bannerSettings.setStatusText("", settingsSheet);
    }
    for (var i = 0; i < SIGNAL_TALLY_NAME_OF_SIGNAL_HISTORY.length; i++) {
      bannerName = SIGNAL_TALLY_NAME_OF_SIGNAL_HISTORY[i];
      bannerSettings = AUTO_IMPORT_BANNER_SETTINGS_FOR_IMPORT[bannerName];
      if (bannerSettings.isEnabled(settingsSheet)) {
        bannerSheet = SpreadsheetApp.getActive().getSheetByName(bannerName);
        if (bannerSheet) {
          var success = checkPages(
            bannerSheet,
            bannerName,
            bannerSettings,
            languageSettings,
            selectedServer,
            settingsSheet,
            authKey);
          if (!success) {
            bannerSettings.setStatusText(
              "Stopped Due to Error:\n" + bannerSettings.getStatusText(settingsSheet),
              settingsSheet);
            break;
          }
        } else {
          bannerSettings.setStatusText("Missing sheet", settingsSheet);
        }
      } else {
        bannerSettings.setStatusText("Skipped", settingsSheet);
      }
    }
  }
  settingsSheet.getRange("E43").setValue(new Date());
}

function checkPages(bannerSheet, bannerName, bannerSettings, languageSettings, selectedServer, settingsSheet, authKey) {
  bannerSettings.setStatusText("Starting", settingsSheet);
  /* Get latest signal from banner */
  var iLastRow = bannerSheet.getRange(2, 5, bannerSheet.getLastRow(), 1).getValues().filter(String).length;
  var signalTextString;
  var lastSignalDateAndTimeString;
  var lastSignalDateAndTime;
  if (iLastRow && iLastRow != 0 ) {
    iLastRow++;
    lastSignalDateAndTimeString = bannerSheet.getRange("E" + iLastRow).getValue();
    signalTextString = bannerSheet.getRange("A" + iLastRow).getValue();
    if (lastSignalDateAndTimeString) {
      bannerSettings.setStatusText(`Last signal: ${lastSignalDateAndTimeString}`, settingsSheet);
      lastSignalDateAndTimeString = lastSignalDateAndTimeString.split(" ").join("T");
      lastSignalDateAndTime = new Date(lastSignalDateAndTimeString+".000Z");
    } else {
      iLastRow = 1;
      bannerSettings.setStatusText("No previous signals", settingsSheet);
    }
    iLastRow++; // Move last row to new row
  } else {
    iLastRow = 2; // Move last row to new row
    bannerSettings.setStatusText("", settingsSheet);
  }
  
  var extractSignals = [];
  var page = 1;
  var queryBannerCode = bannerSettings.gacha_type;
  var numberOfSignalPerPage = 6;
  var urlForBanner = getSignalHistoryUrl(
    selectedServer,
    queryBannerCode,
    languageSettings,
    numberOfSignalPerPage,
    authKey);
  var failed = 0;
  var is_done = false;
  var end_id = 0;
  
  var checkPreviousDateAndTimeString = "";
  var checkPreviousDateAndTime;
  var checkOneSecondOffDateAndTime;
  var overrideIndex = 0;
  var textSignal;
  var oldTextSignal;
  while (!is_done) {
    bannerSettings.setStatusText("Loading page: "+page, settingsSheet);
    var response = UrlFetchApp.fetch(urlForBanner+"&page="+page+"&end_id="+end_id);
    var jsonResponse = response.getContentText();
    var jsonDict = JSON.parse(jsonResponse);
    var jsonDictData = jsonDict["data"];
    if (jsonDictData) {
      var listOfSignals = jsonDictData["list"];
      var listOfSignalsLength = listOfSignals.length;
      var signal;
      if (listOfSignalsLength > 0) {
        for (var i = 0; i < listOfSignalsLength; i++) {
          signal = listOfSignals[i];
          var dateAndTimeString = signal['time'];
          textSignal = signal['item_type']+signal['name'];
          /* Mimic the website in showing specific language wording */
          if (signal['rank_type'] == 2) {
            textSignal += languageSettings["4_star"];
          } else if (signal['rank_type'] == 3) {
            textSignal += languageSettings["5_star"];
          }
          oldTextSignal = textSignal+dateAndTimeString;
          var gachaString = "gacha_type_"+signal['gacha_type'];
          var bannerName = "Error New Banner";
          if (gachaString in languageSettings) {
            bannerName = languageSettings[gachaString];
          }
          textSignal += bannerName+dateAndTimeString;

          var dateAndTimeStringModified = dateAndTimeString.split(" ").join("T");
          var signalDateAndTime = new Date(dateAndTimeStringModified+".000Z");

          if (overrideIndex == 0 && checkPreviousDateAndTime) {
            /* Check one second difference from previous single signal */
            checkOneSecondOffDateAndTime = new Date(checkPreviousDateAndTime.valueOf());
            checkOneSecondOffDateAndTime.setSeconds(checkOneSecondOffDateAndTime.getSeconds()-1);
            if (checkOneSecondOffDateAndTime.valueOf() == signalDateAndTime.valueOf()) {
              var nextSignalIndex = i+1;
              if (nextSignalIndex < listOfSignalsLength) {
                var nextSignal = listOfSignals[nextSignalIndex];
                var nextDateAndTimeString = nextSignal['time'];
                var nextDateAndTimeStringModified = nextDateAndTimeString.split(" ").join("T");
                var nextSignalDateAndTime = new Date(nextDateAndTimeStringModified+".000Z");
                if (checkOneSecondOffDateAndTime.valueOf() == nextSignalDateAndTime.valueOf()) {
                  // Due to signal date and time is only second difference, it's therefore a multi. Override previous signal to match.
                  checkPreviousDateAndTimeString = dateAndTimeString;
                  checkPreviousDateAndTime = new Date(signalDateAndTime.valueOf());
                }
              }
            }
          }
          if (checkPreviousDateAndTimeString === dateAndTimeString) {
            // Found matching date and time to previous signal
            if (overrideIndex == 0) {
              // Start multi 10 index
              var previousSignalIndex = extractSignals.length - 1;
              var previousSignal = extractSignals[previousSignalIndex];
              overrideIndex = 10;
              previousSignal[1] = overrideIndex;
              extractSignals[previousSignalIndex] = previousSignal;
            }
            if (overrideIndex == 1) {
              bannerSettings.setStatusText(
                `Error: Multi signal contains 11 within same date and time: ${dateAndTimeString}, found so far: ${extractSignals.length}`,
                settingsSheet);
              return false;
            } else {
              overrideIndex--;
            }
          } else {
            if (overrideIndex > 1) {
              // Resume counting down when override is set more than 1, add a second to checkPreviousDateAndTime
              checkPreviousDateAndTime.setSeconds(checkPreviousDateAndTime.getSeconds()-1);
              if (checkPreviousDateAndTime.valueOf() == signalDateAndTime.valueOf()) {
                // Within 1 second range resuming multi count
                overrideIndex--;
              } else {
                bannerSettings.setStatusText(
                  `Error: Multi signal is incomplete with override ${overrideIndex}@${dateAndTimeString}, found so far: ${extractSignals.length}`,
                  settingsSheet);
                return false;
              }
            } else {
              // Default value for single signals
              overrideIndex = 0;
            }
            checkPreviousDateAndTimeString = dateAndTimeString;
            checkPreviousDateAndTime = new Date(signalDateAndTime.valueOf());
          }
          if (lastSignalDateAndTime >= signalDateAndTime) {
            // Banner already got this signal
            is_done = true;
            break;
          } else {
            extractSignals.push([textSignal, (overrideIndex > 0 ? overrideIndex:null)]);
          }
        }
        if (!is_done && numberOfSignalPerPage == listOfSignalsLength) {
          end_id = signal['id'];
          page++;
        } else {
          // If list isn't the size requested, it would mean there is no more signals.
          is_done = true;
        }
      } else {
        is_done = true;
      }
    } else {
      var message = jsonDict["message"];
      var return_code = jsonDict["retcode"];

      var title ="Error code: "+return_code;
      SpreadsheetApp.getActiveSpreadsheet().toast(message, title);

      switch (return_code) {
        case AUTO_IMPORT_URL_ERROR_CODE_AUTHKEY_DENIED:
          bannerSettings.setStatusText("feedback URL\nNo Longer Works", settingsSheet);
          return false;
        case AUTO_IMPORT_URL_ERROR_CODE_AUTH_TIMEOUT:
          bannerSettings.setStatusText("auth timeout", settingsSheet);
          return false;
        case AUTO_IMPORT_URL_ERROR_CODE_AUTH_INVALID:
          bannerSettings.setStatusText("auth invalid", settingsSheet);
          return false;
        case AUTO_IMPORT_URL_ERROR_CODE_REQUEST_PARAMS:
          bannerSettings.setStatusText("Change server setting", settingsSheet);
          return false;
        default:
          bannerSettings.setStatusText(`Unknown return code: ${return_code}`, settingsSheet);
          failed++;
          if (failed > 2){
            bannerSettings.setStatusText("Failed too many times", settingsSheet);
            // Preserve legacy behavior which did not treat this the same as other error code
            // cases
            return true;
          }
      }
    }
  }

  if (extractSignals.length > 0) {
    var now = new Date();
    var sixMonthBeforeNow = new Date(now.valueOf());
    sixMonthBeforeNow.setMonth(now.getMonth() - 6);
    var isValid = true;
    var outputString = "Found: "+extractSignals.length;
    if (!lastSignalDateAndTime) {
      // fresh history sheet no last date to check
      outputString += ", with signal history being empty"
    } else if (lastSignalDateAndTime < sixMonthBeforeNow) {
      // Check if last signal found is more than 6 months, no further validation
      outputString += ", last signal saved was 6 months ago, maybe missing signals inbetween"
    } else {
      if (signalTextString !== textSignal) {
        if (signalTextString !== oldTextSignal) {
          // API didn't reach to your last signal stored on the sheet, meaning the API is incomplete
          isValid = false;
          outputString = "Error your recently found signals did not reach to your last signal, found: "+extractSignals.length+", please try again miHoYo may have sent incomplete signal data.";
        }
      }
    }
    if (isValid) {
      extractSignals.reverse();
      bannerSheet.getRange(iLastRow, 1, extractSignals.length, 2).setValues(extractSignals);
    }
    bannerSettings.setStatusText(outputString, settingsSheet);
  } else {
    bannerSettings.setStatusText("Nothing to add", settingsSheet);
  }

  return true;
}

function getSignalHistoryUrl(selectedServer, queryBannerCode, languageSettings, numberOfSignalPerPage, authKey, gameBiz ="nap_global") {
    var base_url;
    if (selectedServer == "China") {
      base_url = AUTO_IMPORT_URL_CHINA;
    } else {
      base_url = AUTO_IMPORT_URL;
    }
    return `${base_url}?${AUTO_IMPORT_ADDITIONAL_QUERY.join("&")}&authkey=${authKey}&lang=${languageSettings['code']}&gacha_type=${queryBannerCode}&size=${numberOfSignalPerPage}&game_biz=${gameBiz}`;
}
