// Functions for User Endpoints
function handleDeleteUserEndp(idsParam, ssUsers, idData){
    let deletedUsers = [];
    let notFoundUsers = [];
    let deleteRows = [];

    if (!idsParam) {
        return sendJSON_({
            "status": "error",
            "code": 400,
            "message": "UserID is required for delete operation"
        });
    }
    sheetLog(ws, "Deleting user IDs: " + idsParam);
    // Separate the ids
    let userIds = idsParam.join(",").split(",").map(id => parseInt(id.trim()));
    sheetLog(ws, "Array of user IDs: " + JSON.stringify(userIds));

    // Initial validation
    userIds.forEach(function(reqId) {
        if (isNaN(parseInt(reqId))) {
            notFoundUsers.push({
                "error": "Invalid UserID format",
                "data": reqId
            });
            return;  
        }
    });

    // Early exit if all IDs are invalid
    if (notFoundUsers.length === userIds.length) {
        return sendJSON_({
            "status": "error",
            "code": 404,
            "message": "All UserIDs are invalid",
            "data": notFoundUsers
        });
    }

    // Process valid IDs
    userIds.forEach(function(reqId) {
        let userFound = false;
        if (isNaN(parseInt(reqId))) return;  // Skip invalid IDs

        idData.forEach(function(sheetId, rowIndex){
            if (parseInt(sheetId) === parseInt(reqId)){
                console.log("User found: " + reqId + " at row index: " + rowIndex);
                sheetLog(ws, "Array of user IDs: " + "User found: " + reqId + " at row index: " + rowIndex);
                deleteRows.push(rowIndex + 2)
                deletedUsers.push(reqId);
                userFound = true;
            }
        });
        if (!userFound) {
            notFoundUsers.push({
                "error": "User not found",
                "data": reqId
            });
        }
    });

    if (deleteRows.length > 0) {
        deleteRows.sort((a, b) => b - a);  // Sorting in descending order, to avoid row index issues while deleting
        deleteRows.forEach(function (row) {
            sheetLog(ws, "Deleting row: " + row);
            console.log("Deleting row: " + row);
            ssUsers.deleteRow(row);
        });
    }

    if (deletedUsers.length > 0) {
        return sendJSON_({
            "status": notFoundUsers.length > 0 ? "partial_success" : "success",
            "code": notFoundUsers.length > 0 ? 206 : 200,
            "message": notFoundUsers.length > 0 ? "Some users were not found" : "Users deleted successfully",
            "data": {
                "deleted": deletedUsers,  
                "notFound": notFoundUsers
            }
        });
    } else {
        return sendJSON_({
            "status": "error",
            "code": 404,
            "message": "No users found to delete"
        });
    }


}



function handleUpdateUserEndp(bodyJSON, ssUsers, headersWithId, headersOriginalOrderWithId){
    // Normalize bodyJSON to an array if it's not already(in only one user send case)
    let usersArray = Array.isArray(bodyJSON) ? bodyJSON : [bodyJSON];
    
    // Get headers from the first element in the array
    let headersPassed = Object.keys(usersArray[0]).sort();
    
    const userData = ssUsers.getRange(2, 1, ssUsers.getLastRow()-1, ssUsers.getLastColumn()).getValues()
    let headerRequired = ["UserID"];
    let updatedRows = [];
    let errorRows = []; // To store error details


    if (!checkHeadersForUpdate(headersWithId, headersPassed, headerRequired)) {
        let missingColumns = headerRequired.filter(h => !headersPassed.includes(h));
        let unexpectedColumns = headersPassed.filter(h => !headersWithId.includes(h));
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

    usersArray.forEach(function (user) {
        if (!user["UserID"]) {
            // If UserID is missing, add to errorRows
            errorRows.push({
                "error": "UserID is missing",
                "data": user
            });
            return; // Skip to the next user
        }
        
    // usersArray.forEach(function (user){
        if(checkHeadersForUpdate(headersWithId, headersPassed, headerRequired)){
            let userFound = false;
            let userId = parseInt(user["UserID"]);
            // console.log("type of dataID: " + typeof(userId))
            userData.forEach(function (row, rowIndex){
                if(userId === parseInt(row[0])){
                    console.log("type of dataID in the sheet: " + typeof(row[0]))
                    let updateRow = headersOriginalOrderWithId.map(function (header, headerIndex) {
                        // Use the updated value from 'user' if available; otherwise, keep the original value from 'row'.
                        return user[header] || row[headerIndex];
                    });
                    updatedRows.push({ rowIndex: rowIndex + 2, data: updateRow });
                    userFound = true;
                }
            });
            if (!userFound) {
                errorRows.push({
                    "error": "User not found",
                    "data": user
                });
            }
        }
    });


    // Apply all updates at once after the loop
    if (updatedRows.length > 0) {
        updatedRows.forEach(function (row) {
            ssUsers.getRange(row["rowIndex"], 1, 1, headersOriginalOrderWithId.length).setValues([row.data]);
            delete row["rowIndex"]
        });
    }
    // Send a detailed response
    if (updatedRows.length > 0 || errorRows.length > 0) {
        return sendJSON_({
            "status": errorRows.length > 0 ? "partial_success" : "success",
            "code": errorRows.length > 0 ? 206 : 200, // Use 206 for partial success
            "message": errorRows.length > 0 ? "Some users failed to update" : "Data updated successfully",
            "data": {
                "updated": updatedRows,
                "errors": errorRows
            }
        });
    } else {
        return sendJSON_({
            "status": "error",
            "code": 404,
            "message": "No users found or updated"
        });
    }

}




function handleAddUserEndp(bodyJSON, ssUsers, idData, headers, headersOriginalOrder){
    // for add endpoint
    const headersPassed = Object.keys(bodyJSON).sort();
    if (checkHeaders(headers, headersPassed)) {
        let newUserID = newID(idData);
        const arrayOfData = headersOriginalOrder.map(h => bodyJSON[h]);
        arrayOfData.unshift(newUserID);
        ssUsers.appendRow(arrayOfData);

        bodyJSON.id = newUserID;
        return sendJSON_({
            "status": "success",
            "code": 200,
            "message": "Data added successfully", data: bodyJSON
        })
    } else {
        // If the headers are not correct
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
        })
    }
}


function handleBatchAddUserEndp(bodyJSON, ssUsers, idData, headers, headersOriginalOrder){
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
            ssUsers.getRange(ssUsers.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
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

