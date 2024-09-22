
















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
