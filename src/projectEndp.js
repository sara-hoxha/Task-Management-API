function handleDeleteProjectEndp(idsParam, sheet, idData){
    let deletedProjects = [];
    let notFoundProjects = [];
    let deleteRows = [];

    if (!idsParam) {
        return sendJSON_({
            "status": "error",
            "code": 400,
            "message": "ProjectID is required for delete operation"
        });
    }
    sheetLog(ws, "Deleting project IDs: " + idsParam);
    // Separate the ids
    let projectIds = idsParam.join(",").split(",").map(id => parseInt(id.trim()));
    sheetLog(ws, "Array of project IDs: " + JSON.stringify(projectIds));

    // Initial validation
    projectIds.forEach(function(reqId) {
        if (isNaN(parseInt(reqId))) {
            notFoundProjects.push({
                "error": "Invalid ProjectID format",
                "data": reqId
            });
            return;  
        }
    });

    // Early exit if all IDs are invalid
    if (notFoundProjects.length === projectIds.length) {
        return sendJSON_({
            "status": "error",
            "code": 404,
            "message": "All ProjectIDs are invalid",
            "data": notFoundProjects
        });
    }

    // Process valid IDs
    projectIds.forEach(function(reqId) {
        let projectFound = false;
        if (isNaN(parseInt(reqId))) return;  // Skip invalid IDs

        idData.forEach(function(sheetId, rowIndex){
            if (parseInt(sheetId) === parseInt(reqId)){
                console.log("Project found: " + reqId + " at row index: " + rowIndex);
                sheetLog(ws, "Array of project IDs: " + "Project found: " + reqId + " at row index: " + rowIndex);
                deleteRows.push(rowIndex + 2)
                deletedProjects.push(reqId);
                projectFound = true;
            }
        });
        if (!projectFound) {
            notFoundProjects.push({
                "error": "Project not found",
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

    if (deletedProjects.length > 0) {
        return sendJSON_({
            "status": notFoundProjects.length > 0 ? "partial_success" : "success",
            "code": notFoundProjects.length > 0 ? 206 : 200,
            "message": notFoundProjects.length > 0 ? "Some projects were not found" : "Projects deleted successfully",
            "data": {
                "deleted": deletedProjects,  
                "notFound": notFoundProjects
            }
        });
    } else {
        return sendJSON_({
            "status": "error",
            "code": 404,
            "message": "No projects found to delete"
        });
    }


}


function handleUpdateProjectEndp(bodyJSON, sheet, headersWithId, headersOriginalOrderWithId){
    // Normalize bodyJSON to an array if it's not already(for only one project send case)
    let projectsArray = Array.isArray(bodyJSON) ? bodyJSON : [bodyJSON];
    
    // Get headers from the first element in the array
    let headersPassed = Object.keys(projectsArray[0]).sort();
    
    const projectData = sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).getValues()
    let headerRequired = ["ProjectID"];
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

    projectsArray.forEach(function (project) {
        if (!project["ProjectID"]) {
            // If ProjectID is missing, add to errorRows
            errorRows.push({
                "error": "ProjectID is missing",
                "data": project
            });
            return; // Skip to the next project
        }
        
    // projectsArray.forEach(function (project){
        if(checkHeadersForUpdate(headersWithId, headersPassed, headerRequired)){
            let projectFound = false;
            let projectId = parseInt(project["ProjectID"]);
            // console.log("type of dataID: " + typeof(projectId))
            projectData.forEach(function (row, rowIndex){
                if(projectId === parseInt(row[0])){
                    console.log("type of dataID in the sheet: " + typeof(row[0]))
                    let updateRow = headersOriginalOrderWithId.map(function (header, headerIndex) {
                        // Use the updated value from 'project' if available; otherwise, keep the original value from 'row'.
                        return project[header] || row[headerIndex];
                    });
                    updatedRows.push({ rowIndex: rowIndex + 2, data: updateRow });
                    projectFound = true;
                }
            });
            if (!projectFound) {
                errorRows.push({
                    "error": "Project not found",
                    "data": project
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
            "message": errorRows.length > 0 ? "Some projects failed to update" : "Data updated successfully",
            "data": {
                "updated": updatedRows,
                "errors": errorRows
            }
        });
    } else {
        return sendJSON_({
            "status": "error",
            "code": 404,
            "message": "No projects found or updated"
        });
    }

}




function handleAddProjectEndp(bodyJSON, sheet, idData, headers, headersOriginalOrder){
    // for add endpoint
    const headersPassed = Object.keys(bodyJSON).sort();
    if (checkHeaders(headers, headersPassed)) {
        let newProjectID = newID(idData);
        const arrayOfData = headersOriginalOrder.map(h => bodyJSON[h]);
        arrayOfData.unshift(newProjectID);
        sheet.appendRow(arrayOfData);

        bodyJSON.id = newProjectID;
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








function handleBatchAddProjectEndp(bodyJSON, sheet, idData, headers, headersOriginalOrder){
    // for batch endpoint
    if (bodyJSON.length === 0) {
        return sendJSON_({
            "status": "error",
            "code": 400,
            "message": "No data provided for batch addition"
        });
    }

    // Extract headers from the first project in the array
    const headersPassed = Object.keys(bodyJSON[0]).sort();
    // check the headers for each project too
    if (checkHeaders(headers, headersPassed)) {
        let newRows = [];
        let errorMessages = [];
    
        bodyJSON.forEach((project, index) => {
            const projectHeadersPassed = Object.keys(project).sort();
    
            if (!checkHeaders(headers, projectHeadersPassed)) {
                let missingColumns = headers.filter(h => !projectHeadersPassed.includes(h));
                let unexpectedColumns = projectHeadersPassed.filter(h => !headers.includes(h));
                
                errorMessages.push({
                    projectIndex: index,
                    missingColumns: missingColumns,
                    unexpectedColumns: unexpectedColumns
                });
            } else {
                let arrayOfData = headersOriginalOrder.map(h => project[h]);
                let newProjectID = newID(idData);
                // update the data, so each Project gets a new id
                idData.push(newProjectID);
                arrayOfData.unshift(newProjectID);
                newRows.push(arrayOfData);
            }
        });
    
        if (errorMessages.length > 0) {
            return sendJSON_({
                "status": "error",
                "code": 400,
                "message": "One or more projects have missing or extra data columns",
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
