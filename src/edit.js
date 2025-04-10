/*
 * Version 0.02 made by yippym - 2025-04-10 01:05
 * https://github.com/Yippy/signal-tally-sheet
 */
function onEdit(e) {
    const sheet = e.range.getSheet(); 
    if(sheet.getName() == SIGNAL_TALLY_AGENTS_SHEET_NAME || sheet.getName() == SIGNAL_TALLY_W_ENGINES_SHEET_NAME) {
        if (e.value == "TRUE") {
            var allowableColumns = sheet.getRange(1,12).getValue();
            allowableColumns = String(allowableColumns).split(",");
            if (allowableColumns.includes(String(e.range.columnStart))) {
                sheet.getRange(e.range.rowStart, e.range.columnStart).setValue(false);
                var characterName = sheet.getRange(e.range.rowStart, e.range.columnStart+1).getValue();
                var indexRow = sheet.getRange(1,e.range.columnStart).getValue();
                sheet.getRange(indexRow, e.range.columnStart).setValue(characterName);
            }
        }
    }
}