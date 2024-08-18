// https://script.google.com/macros/***IKaouy/exec?user=true&id=idnumber




// adding users in bulk - array of objects
// [
    // {
    //     "username": "jsmith",
    //     "email": "jsmith@example.com",
    //     "age": 30
    // },
    // {
    //     "username": "ajones",
    //     "email": "ajones@example.com",
    //     "age": 25
    // }
// ]
// adding only one user - only onbject will be send

// Add a Single User
// https://yourapp.com?api&users&add
// Add Multiple Users
// https://yourapp.com?api&users&batch
// Put - Edit user information
// https://yourapp.com?api&users&put


function doPost(e) {
    const ssUsers = ws.getSheetByName("Users")
    let idData = ssUsers.getRange(2,1,ssUsers.getLastRow()-1).getValues()
    const headers = ssUsers.getRange(1, 1, 1, ssUsers.getLastColumn()).getValues()[0]
    const headersOriginalOrder = headers.slice(); // Copy headers to preserve original order
    headersOriginalOrder.shift(); // Remove ID column header from copy
    headers.shift(); // Remove ID column header from original headers
    headers.sort() // Sort headers for comparison
    const body = e.postData.contents // Get the POST request body as a string
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

    if (params["api"] && params["users"] && params["add"] && !Array.isArray(bodyJSON)) {
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
    if (params["api"] && params["users"] && params["batch"] && Array.isArray(bodyJSON)) {
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
}


// function doPost(e){
//     const ssUsers = ws.getSheetByName("Users")
//     let idData = ssUsers.getRange(2,1,ssUsers.getLastRow()-1).getValues()
//     const headers = ssUsers.getRange(1, 1, 1, ssUsers.getLastColumn()).getValues()[0]
//     const headersOriginalOrder = headers.slice();
//     headersOriginalOrder.shift(); // remove id columns header
    
//     headers.shift();
//     headers.sort() // Sort headers for comparison

//     const body = e.postData.contents
//     let bodyJSON;
//     // const bodyJSON = JSON.parse(body);

//     // for invalid JSON format
//     try {
//         bodyJSON = JSON.parse(body);
//     } catch (error) {
//         return sendJSON_({
//             "status": "error",
//             "code": 400,
//             "message": "Invalid JSON format"
//         });
//     }

//     // const headersPassed = Object.keys(bodyJSON).sort();
//     const params = e.parameters;

//     if (params["api"] && params["users"] && params["add"] && !Array.isArray(bodyJSON)) {
//         // for add endpoint
//         const headersPassed = Object.keys(bodyJSON).sort();
//         if (checkHeaders(headers, headersPassed)) {
//             let newUserID = newID(idData);
//             const arrayOfData = headersOriginalOrder.map(h => bodyJSON[h]);
//             arrayOfData.unshift(newUserID);
//             ssUsers.appendRow(arrayOfData);

//             bodyJSON.id = newUserID;
//             return sendJSON_({
//                 "status": "success",
//                 "code": 200,
//                 "message": "Data added successfully", data: bodyJSON
//             })
//         } else {
//             // If the headers are not correct
//             // Error response for incorrect headers
//             let missingColumns = headers.filter(h => !headersPassed.includes(h));
//             let unexpectedColumns = headersPassed.filter(h => !headers.includes(h));
//             return sendJSON_({
//                 "status": "error",
//                 "code": 400,
//                 "message": "Missing required data columns or extra data columns provided",
//                 "details": {
//                     "missingColumns": missingColumns,
//                     "unexpectedColumns": unexpectedColumns
//                 }
//             })
//         }

//     // for batch endpoint
//     } else if (params["api"] && params["users"] && params["batch"] && Array.isArray(bodyJSON)) {
//         if (bodyJSON.length === 0) {
//             return sendJSON_({
//                 "status": "error",
//                 "code": 400,
//                 "message": "No data provided for batch addition"
//             });
//         }

//         // Extract headers from the first user in the array
//         const headersPassed = Object.keys(bodyJSON[0]).sort();

//         if (checkHeaders(headers, headersPassed)) {
//             let newRows = [];
//             let errorMessages = [];
        
//             bodyJSON.forEach((user, index) => {
//                 const userHeadersPassed = Object.keys(user).sort();
        
//                 if (!checkHeaders(headers, userHeadersPassed)) {
//                     let missingColumns = headers.filter(h => !userHeadersPassed.includes(h));
//                     let unexpectedColumns = userHeadersPassed.filter(h => !headers.includes(h));
                    
//                     errorMessages.push({
//                         userIndex: index,
//                         missingColumns: missingColumns,
//                         unexpectedColumns: unexpectedColumns
//                     });
//                 } else {
//                     let arrayOfData = headersOriginalOrder.map(h => user[h]);
//                     let newUserID = newID(idData);
//                     idData.push(newUserID);
//                     arrayOfData.unshift(newUserID);
//                     newRows.push(arrayOfData);
//                 }
//             });
        
//             if (errorMessages.length > 0) {
//                 return sendJSON_({
//                     "status": "error",
//                     "code": 400,
//                     "message": "One or more users have missing or extra data columns",
//                     "details": errorMessages
//                 });
//             }
        
//             if (newRows.length > 0) {
//                 ssUsers.getRange(ssUsers.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
//             }
        
//             return sendJSON_({
//                 "status": "success",
//                 "code": 200,
//                 "message": "Batch data added successfully",
//                 "details": {
//                     "addedRecords": newRows.length
//                 }
//             });
//         } else {
//             // Error response for incorrect headers
//             let missingColumns = headers.filter(h => !headersPassed.includes(h));
//             let unexpectedColumns = headersPassed.filter(h => !headers.includes(h));
//             return sendJSON_({
//                 "status": "error",
//                 "code": 400,
//                 "message": "Missing required data columns or extra data columns provided",
//                 "details": {
//                     "missingColumns": missingColumns,
//                     "unexpectedColumns": unexpectedColumns
//                 }
//                 });
//             }
//     // if params are incorrect
//     } else {
//         return sendJSON_({
//             "status": "error",
//             "code": 400,
//             "message": "Invalid endpoint or data format"
//         });
//     }



// }





