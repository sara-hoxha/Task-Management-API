function newID(data){
    let maxID = 0
    data.forEach(id => {
        if(id > maxID){maxID = parseInt(id)}
    });
    return maxID + 1;
}

function checkHeaders(headers, headersPassed) {
    // if both are true, returns true
    return headers.length === headersPassed.length && 
            headers.every(item => headersPassed.includes(item));
}

function checkHeadersForUpdate(headersWithId, headersPassed, headerRequired) {
    // if both are true, returns true
    return  headersWithId.every(item => headersPassed.includes(item)) && headerRequired.every(item => headersPassed.includes(item));
        
}





// to log doPost logs
var LOG_SHEET_NAME = "Debug Logs";

function sheetLog(doc, msg) {
    Logger.log(msg); // Log to Apps Script Logs for debugging in the script editor

    // Get or create the Debug Logs sheet
        var logSheet = doc.getSheetByName(LOG_SHEET_NAME);
    if (!logSheet) {
        logSheet = doc.insertSheet(LOG_SHEET_NAME);
        logSheet.getRange(1, 1, 1, 2).setValues([["Timestamp", "Log Messages"]]);
    }

    // Append the new log message to the next row
    var nxtLogRow = logSheet.getLastRow() + 1;
    logSheet.getRange(nxtLogRow, 1, 1, 2).setValues([[new Date(), msg]]);
}
