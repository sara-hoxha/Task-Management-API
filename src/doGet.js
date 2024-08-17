// todo
// User Management

// POST /users: Create a new user
// GET /users/
// : Retrieve user details
// PUT /users/
// : Update user information
// DELETE /users/
// : Delete a user

// https://script.google.com/macros/***IKaouy/exec?user=true&id=idnumber


const ws = SpreadsheetApp.openById("1BeOR30LcEp7Oxt4lOv45BCquHCQNqsCPfQG34-qOWRs");

function doGet(e) {
    if (e.parameter.api) {
        return handleApiGet(e);
    } else{
        return "Wrong Endpoint"
    }

    // } else {
    //     return HtmlService.createHtmlOutputFromFile('Index');
    // }
}


// /user/userid
function handleApiGet(e){
    // ?api&username=jsmith&age=21
    // e.parameters
    const ssUsers = ws.getSheetByName("Users")
    let usersData = ssUsers.getRange(1, 1, ssUsers.getLastRow(), ssUsers.getLastColumn()).getValues();
    let headers = usersData[0]
    usersData.shift();
    const params = e.parameters;

    if (params["user"] && params["user"][0] === "true") {
        if (params["id"] ) {
            const userId = parseInt(params["id"][0]);
            if (!isNaN(userId)) {
                usersData = usersData.filter(function(row){
                    if(row[0] === userId) return row
                })
            } else {
                // If the ID is not a valid number
                return sendJSON_({error: "Invalid user ID" });
            }
        } else {
            // If the 'id' parameter is missing
            return sendJSON_({error: "User ID not provided" });
        }
    
    }

    
    let newData = usersData.map(user => {
        let response = {}
        headers.forEach((header, index) =>{
            response[header] = user[index];
        });
        return response;
    });
    
    response = [{
        status: "success",
        code: 200,
        message: "Request was successful", data: newData}]
    // console.log(response)
        
    
    return sendJSON_(response)
}


// const response = [{status: 200, data: jsonArray}];

// return sendJSON_(response);

function sendJSON_(jsonResponse){
    return ContentService
        .createTextOutput(JSON.stringify(jsonResponse))
        .setMimeType(ContentService.MimeType.JSON);
}