function doPost(e) {
    // const sheet = ws.getSheetByName("Users")
    let sheet;
    const api = params["api"];

    if (api && params["users"])
        sheet = ws.getSheetByName("Users")
    else if(api && params["tasks"]){
        sheet = ws.getSheetByName("Tasks")
    } else if(api && params["projects"]){
        sheet = ws.getSheetByName("Projects")
    }
    let idData = sheet.getRange(2,1,sheet.getLastRow()-1).getValues()
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    // let idData = sheet.getRange(2,1,sheet.getLastRow()-1).getValues()
    // const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    const headersOriginalOrder = headers.slice(); // Copy headers to preserve original order
    const headersOriginalOrderWithId = headers.slice();
    const headersWithId = headers.sort()
    headersOriginalOrder.shift(); // Remove ID column header from copy
    headers.shift(); // Remove ID column header from original headers
    headers.sort() // Sort headers for comparison
    // const body = e.postData.contents // Get the POST request body as a string
    // Check if e.postData and e.postData.contents exist
    const body = e.postData && e.postData.contents ? e.postData.contents : null;
    let bodyJSON; // parsed JSON
    // const bodyJSON = JSON.parse(body);

    // for invalid JSON format
    try {
        bodyJSON = JSON.parse(body); // Attempt to parse JSON
    } catch (error) {
        return sendJSON_({
            "status": "error",
            "code": 400,
            "message": "Invalid JSON format"
        });
    }
    // const headersPassed = Object.keys(bodyJSON).sort();
    const params = e.parameters; // Get parameters from the event
    sheetLog(ws, "Request parameters: " + JSON.stringify(params));
    PropertiesService.getScriptProperties().setProperty("API_KEY", '8MBIbIdf3d8sxBhKCx2D');
    let apiKey = PropertiesService.getScriptProperties().getProperty("API_KEY"); ;
    if(!params["apiKey"] || (params["apiKey"] !== apiKey)) {
        return sendErrorResponse(403, "Access to the requested resource is forbidden")
    }
    else {
        if (api && params["users"]) {
            if (params["add"] && !Array.isArray(bodyJSON)) {
                return handleAddUserEndp(bodyJSON, sheet, idData, headers, headersOriginalOrder);
            } else if (params["batch"] && Array.isArray(bodyJSON)) {
                return handleBatchAddUserEndp(bodyJSON, sheet, idData, headers, headersOriginalOrder);
            } 
            else if (params["update"]) {
                return handleUpdateUserEndp(bodyJSON, sheet, headersWithId, headersOriginalOrderWithId);
            } else if (params["delete"]) {
                let idsParam = params["userID"];
                sheetLog(ws, "params['userID']: " + JSON.stringify(params["userID"]));
                return handleDeleteUserEndp(idsParam, sheet, idData);
            }
        }
        if (api && params["tasks"] && !(params["users"])) {
            if (params["add"] && !Array.isArray(bodyJSON)) {
                return handleAddTaskEndp(bodyJSON, sheet, idData, headers, headersOriginalOrder);
            } else if (params["batch"] && Array.isArray(bodyJSON)) {
                return handleBatchAddTaskEndp(bodyJSON, sheet, idData, headers, headersOriginalOrder);
            } 
            else if (params["update"]) {
                return handleUpdateTaskEndp(bodyJSON, sheet, headersWithId, headersOriginalOrderWithId);
            } else if (params["delete"]) {
                let idsParam = params["userID"];
                sheetLog(ws, "params['userID']: " + JSON.stringify(params["userID"]));
                return handleDeleteTaskEndp(idsParam, sheet, idData);
            }
        }
    }
    return sendErrorResponse(400, "Invalid endpoint or data format");
}

// Delete the function above after you are done with the implementation.

function handleBatchAddUserEndp(bodyJSON, sheet, idData, headers, headersOriginalOrder){
    // for batch endpoint
    if (bodyJSON.length === 0) {
        return sendJSON_({
            "status": "error",
            "code": 400,
            "message": "No data provided for batch addition"
        });
    }

    // Extract headers from the first user in the array
    const headersPassed = Object.keys(bodyJSON[0]).sort();
    // check the headers for each user too
    if (checkHeaders(headers, headersPassed)) {
        let newRows = [];
        let errorMessages = [];
    
        bodyJSON.forEach((user, index) => {
            const userHeadersPassed = Object.keys(user).sort();
    
            if (!checkHeaders(headers, userHeadersPassed)) {
                let missingColumns = headers.filter(h => !userHeadersPassed.includes(h));
                let unexpectedColumns = userHeadersPassed.filter(h => !headers.includes(h));
                
                errorMessages.push({
                    userIndex: index,
                    missingColumns: missingColumns,
                    unexpectedColumns: unexpectedColumns
                });
            } else {
                let arrayOfData = headersOriginalOrder.map(h => user[h]);
                let newUserID = newID(idData);
                // update the data, so each user gets a new id
                idData.push(newUserID);
                arrayOfData.unshift(newUserID);
                newRows.push(arrayOfData);
            }
        });
    
        if (errorMessages.length > 0) {
            return sendJSON_({
                "status": "error",
                "code": 400,
                "message": "One or more users have missing or extra data columns",
                "details": errorMessages
            });
        }
    
        if (newRows.length > 0) {
            sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
        }
    
        return sendJSON_({
            "status": "success",
            "code": 200,
            "message": "Batch data added successfully",
            "details": {
                "addedRecords": newRows.length
            }
        });
    } else {
            // Error response for incorrect headers
            let missingColumns = headers.filter(h => !headersPassed.includes(h));
            let unexpectedColumns = headersPassed.filter(h => !headers.includes(h));
            return sendJSON_({
                "status": "error",
                "code": 400,
                "message": "Missing required data columns or extra data columns provided",
                "details": {
                    "missingColumns": missingColumns,
                    "unexpectedColumns": unexpectedColumns
                }
                });
        }

}
