// function doPost(e) {
//     // const sheet = ws.getSheetByName("Users")
//     let sheet;
//     const api = params["api"];

//     if (api && params["users"])
//         sheet = ws.getSheetByName("Users")
//     else if(api && params["tasks"]){
//         sheet = ws.getSheetByName("Tasks")
//     } else if(api && params["projects"]){
//         sheet = ws.getSheetByName("Projects")
//     }
//     let idData = sheet.getRange(2,1,sheet.getLastRow()-1).getValues()
//     const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
//     // let idData = sheet.getRange(2,1,sheet.getLastRow()-1).getValues()
//     // const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
//     const headersOriginalOrder = headers.slice(); // Copy headers to preserve original order
//     const headersOriginalOrderWithId = headers.slice();
//     const headersWithId = headers.sort()
//     headersOriginalOrder.shift(); // Remove ID column header from copy
//     headers.shift(); // Remove ID column header from original headers
//     headers.sort() // Sort headers for comparison
//     // const body = e.postData.contents // Get the POST request body as a string
//     // Check if e.postData and e.postData.contents exist
//     const body = e.postData && e.postData.contents ? e.postData.contents : null;
//     let bodyJSON; // parsed JSON
//     // const bodyJSON = JSON.parse(body);

//     // for invalid JSON format
//     try {
//         bodyJSON = JSON.parse(body); // Attempt to parse JSON
//     } catch (error) {
//         return sendJSON_({
//             "status": "error",
//             "code": 400,
//             "message": "Invalid JSON format"
//         });
//     }
//     // const headersPassed = Object.keys(bodyJSON).sort();
//     const params = e.parameters; // Get parameters from the event
//     sheetLog(ws, "Request parameters: " + JSON.stringify(params));
//     PropertiesService.getScriptProperties().setProperty("API_KEY", '8MBIbIdf3d8sxBhKCx2D');
//     let apiKey = PropertiesService.getScriptProperties().getProperty("API_KEY"); ;
//     if(!params["apiKey"] || (params["apiKey"] !== apiKey)) {
//         return sendErrorResponse(403, "Access to the requested resource is forbidden")
//     }
//     else {
//         if (api && params["users"]) {
//             if (params["add"] && !Array.isArray(bodyJSON)) {
//                 return handleAddUserEndp(bodyJSON, sheet, idData, headers, headersOriginalOrder);
//             } else if (params["batch"] && Array.isArray(bodyJSON)) {
//                 return handleBatchAddUserEndp(bodyJSON, sheet, idData, headers, headersOriginalOrder);
//             } 
//             else if (params["update"]) {
//                 return handleUpdateUserEndp(bodyJSON, sheet, headersWithId, headersOriginalOrderWithId);
//             } else if (params["delete"]) {
//                 let idsParam = params["userID"];
//                 sheetLog(ws, "params['userID']: " + JSON.stringify(params["userID"]));
//                 return handleDeleteUserEndp(idsParam, sheet, idData);
//             }
//         }
//         if (api && params["tasks"] && !(params["users"])) {
//             if (params["add"] && !Array.isArray(bodyJSON)) {
//                 return handleAddTaskEndp(bodyJSON, sheet, idData, headers, headersOriginalOrder);
//             } else if (params["batch"] && Array.isArray(bodyJSON)) {
//                 return handleBatchAddTaskEndp(bodyJSON, sheet, idData, headers, headersOriginalOrder);
//             } 
//             else if (params["update"]) {
//                 return handleUpdateTaskEndp(bodyJSON, sheet, headersWithId, headersOriginalOrderWithId);
//             } else if (params["delete"]) {
//                 let idsParam = params["userID"];
//                 sheetLog(ws, "params['userID']: " + JSON.stringify(params["userID"]));
//                 return handleDeleteTaskEndp(idsParam, sheet, idData);
//             }
//         }
//     }
//     return sendErrorResponse(400, "Invalid endpoint or data format");
// }

// Delete the function above after you are done with the implementation.

function handleDeleteTaskEndp(idsParam, sheet, idData){
    let deletedTasks = [];
    let notFoundTasks = [];
    let deleteRows = [];

    if (!idsParam) {
        return sendJSON_({
            "status": "error",
            "code": 400,
            "message": "TaskID is required for delete operation"
        });
    }
    sheetLog(ws, "Deleting task IDs: " + idsParam);
    // Separate the ids
    let taskIds = idsParam.join(",").split(",").map(id => parseInt(id.trim()));
    sheetLog(ws, "Array of task IDs: " + JSON.stringify(taskIds));

    // Initial validation
    taskIds.forEach(function(reqId) {
        if (isNaN(parseInt(reqId))) {
            notFoundTasks.push({
                "error": "Invalid TaskID format",
                "data": reqId
            });
            return;  
        }
    });

    // Early exit if all IDs are invalid
    if (notFoundTasks.length === taskIds.length) {
        return sendJSON_({
            "status": "error",
            "code": 404,
            "message": "All TaskIDs are invalid",
            "data": notFoundTasks
        });
    }

    // Process valid IDs
    taskIds.forEach(function(reqId) {
        let taskFound = false;
        if (isNaN(parseInt(reqId))) return;  // Skip invalid IDs

        idData.forEach(function(sheetId, rowIndex){
            if (parseInt(sheetId) === parseInt(reqId)){
                console.log("Task found: " + reqId + " at row index: " + rowIndex);
                sheetLog(ws, "Array of task IDs: " + "Task found: " + reqId + " at row index: " + rowIndex);
                deleteRows.push(rowIndex + 2)
                deletedTasks.push(reqId);
                taskFound = true;
            }
        });
        if (!taskFound) {
            notFoundTasks.push({
                "error": "Task not found",
                "data": reqId
            });
        }
    });

    if (deleteRows.length > 0) {
        deleteRows.sort((a, b) => b - a);  // Sorting in descending order, to avoid row index issues while deleting
        deleteRows.forEach(function (row) {
            sheetLog(ws, "Deleting row: " + row);
            console.log("Deleting row: " + row);
            sheet.deleteRow(row);
        });
    }

    if (deletedTasks.length > 0) {
        return sendJSON_({
            "status": notFoundTasks.length > 0 ? "partial_success" : "success",
            "code": notFoundTasks.length > 0 ? 206 : 200,
            "message": notFoundTasks.length > 0 ? "Some tasks were not found" : "Tasks deleted successfully",
            "data": {
                "deleted": deletedTasks,  
                "notFound": notFoundTasks
            }
        });
    } else {
        return sendJSON_({
            "status": "error",
            "code": 404,
            "message": "No tasks found to delete"
        });
    }


}



function handleUpdateTaskEndp(bodyJSON, sheet, headersWithId, headersOriginalOrderWithId){
    // Normalize bodyJSON to an array if it's not already(in only one task send case)
    let tasksArray = Array.isArray(bodyJSON) ? bodyJSON : [bodyJSON];
    
    // Get headers from the first element in the array
    let headersPassed = Object.keys(tasksArray[0]).sort();
    
    const taskData = sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).getValues()
    let headerRequired = ["TaskID"];
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

    tasksArray.forEach(function (task) {
        if (!task["TaskID"]) {
            // If TaskID is missing, add to errorRows
            errorRows.push({
                "error": "TaskID is missing",
                "data": task
            });
            return; // Skip to the next task
        }
        
    // tasksArray.forEach(function (task){
        if(checkHeadersForUpdate(headersWithId, headersPassed, headerRequired)){
            let taskFound = false;
            let taskId = parseInt(task["TaskID"]);
            // console.log("type of dataID: " + typeof(taskId))
            taskData.forEach(function (row, rowIndex){
                if(taskId === parseInt(row[0])){
                    console.log("type of dataID in the sheet: " + typeof(row[0]))
                    let updateRow = headersOriginalOrderWithId.map(function (header, headerIndex) {
                        // Use the updated value from 'task' if available; otherwise, keep the original value from 'row'.
                        return task[header] || row[headerIndex];
                    });
                    updatedRows.push({ rowIndex: rowIndex + 2, data: updateRow });
                    taskFound = true;
                }
            });
            if (!taskFound) {
                errorRows.push({
                    "error": "Task not found",
                    "data": task
                });
            }
        }
    });


    // Apply all updates at once after the loop
    if (updatedRows.length > 0) {
        updatedRows.forEach(function (row) {
            sheet.getRange(row["rowIndex"], 1, 1, headersOriginalOrderWithId.length).setValues([row.data]);
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




function handleAddTaskEndp(bodyJSON, sheet, idData, headers, headersOriginalOrder){
    // for add endpoint
    const headersPassed = Object.keys(bodyJSON).sort();
    if (checkHeaders(headers, headersPassed)) {
        let newTaskID = newID(idData);
        const arrayOfData = headersOriginalOrder.map(h => bodyJSON[h]);
        arrayOfData.unshift(newTaskID);
        sheet.appendRow(arrayOfData);

        bodyJSON.id = newTaskID;
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

function handleBatchAddTaskEndp(bodyJSON, sheet, idData, headers, headersOriginalOrder){
    // for batch endpoint
    if (bodyJSON.length === 0) {
        return sendJSON_({
            "status": "error",
            "code": 400,
            "message": "No data provided for batch addition"
        });
    }

    // Extract headers from the first task in the array
    const headersPassed = Object.keys(bodyJSON[0]).sort();
    // check the headers for each task too
    if (checkHeaders(headers, headersPassed)) {
        let newRows = [];
        let errorMessages = [];
    
        bodyJSON.forEach((task, index) => {
            const taskHeadersPassed = Object.keys(task).sort();
    
            if (!checkHeaders(headers, taskHeadersPassed)) {
                let missingColumns = headers.filter(h => !taskHeadersPassed.includes(h));
                let unexpectedColumns = taskHeadersPassed.filter(h => !headers.includes(h));
                
                errorMessages.push({
                    taskIndex: index,
                    missingColumns: missingColumns,
                    unexpectedColumns: unexpectedColumns
                });
            } else {
                let arrayOfData = headersOriginalOrder.map(h => task[h]);
                let newTaskID = newID(idData);
                // update the data, so each Task gets a new id
                idData.push(newTaskID);
                arrayOfData.unshift(newTaskID);
                newRows.push(arrayOfData);
            }
        });
    
        if (errorMessages.length > 0) {
            return sendJSON_({
                "status": "error",
                "code": 400,
                "message": "One or more tasks have missing or extra data columns",
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
