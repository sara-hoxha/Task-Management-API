const ws = SpreadsheetApp.openById("1BeOR30LcEp7Oxt4lOv45BCquHCQNqsCPfQG34-qOWRs");

function doGet(e) {
    const params = e.parameters;
    sheetLog(ws, "e.parameters: " + JSON.stringify(e.parameters))
    const api = params["api"];
    let sheet;
    


    // Determine which sheet to use based on the API type
    if (api && params["users"]) {
        sheet = ws.getSheetByName("Users");
        return handleApiGet(sheet, params, "users");
    } else if (api && params["tasks"]) {
        sheet = ws.getSheetByName("Tasks");
        return handleApiGet(sheet, params, "tasks");
    } else if (api && params["projects"]) {
        sheet = ws.getSheetByName("Projects");
        return handleApiGet(sheet, params, "projects");
    } else {
        return sendErrorResponse(400, "Invalid API endpoint");
    }
}

function mockDoGet() {
    // tasks&status=Pending
  // Simulated parameters for testing
  const params = {
    api: "true",
    tasks: "true",
    id: "1", 
    // limit: "10",
    // page: "1"
  };

  const e = { parameters: params }; // Mock event object
  const sheet = ws.getSheetByName("Tasks"); // Assume "Users" sheet exists
  return handleApiGet(sheet, params, "tasks");
}

function testHandleApiGet() {
  const response = mockDoGet();
  Logger.log(response); // Log the response to see the output
}


// GET /?api&users&limit=5&page=1
// GET /?api&tasks&id=123
// GET /?api&tasks&status=Pending&limit=10&page=2
// GET /?api&projects&owner=37


// A handler for the 'GET' API with filtering and pagination
function handleApiGet(sheet, params, type) {
    const allData = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
    const headers = allData[0];
    // new array with all data except the headers
    const records = allData.slice(1);
    
    sheetLog(ws, "params: " + JSON.stringify(params))
    sheetLog(ws, "params[id]: " + JSON.stringify(params["id"]))
    
    let filteredData = records;
    sheetLog(ws, "allData: " + records);
    sheetLog(ws, "All Data: " + JSON.stringify(allData));

    // Filter data based on parameters
    if (params["id"]) {
        sheetLog(ws, "params[id]: " + JSON.stringify(params["id"][0]))
        const idIndex = headers.indexOf(type === "users" ? "UserID" : type === "tasks" ? "TaskID" : "ProjectID");
        // const idIndex = 0; // ID is in the first column in all sheets
        // filters column 1 based on ID passed
        // let idsParam = params["id"];
        // sheetLog(ws, "params['id']: " + JSON.stringify(params["id"]));
        filteredData = filteredData.filter((row) => {return row[idIndex].toString() === params["id"][0]});
        sheetLog(ws, "Filtered by ID: " + filteredData);
    }

    // Additional filtering for users, tasks, and projects
    if (type === "tasks" && params["status"]) {
        sheetLog(ws, "params[status]: " + JSON.stringify(params["status"][0]))
        const statusIndex = headers.indexOf("Status");
        sheetLog(ws, "statusIndex for tasks: " + statusIndex );
        // const statusIndex = 6; // status is on 7th column
        filteredData = filteredData.filter((row) => {return row[statusIndex].toLower() === params["status"][0]});
        sheetLog(ws, "Filtered by Status: " + filteredData);
    }

    if (type === "projects" && params["owner"]) {
        sheetLog(ws, "params[owner]: " + JSON.stringify(params["owner"][0]))
        const ownerIndex = headers.indexOf("Owner");
        // const ownerIndex = 4; // owner is on 5th column
        filteredData = filteredData.filter((row) => {return row[ownerIndex].toString() === params["owner"][0]});
        sheetLog(ws, "Filtered by Owner: " + filteredData);
    }

    // Pagination
    // the number of records I want to show per page
    const limit = params["limit"] ? parseInt(params["limit"], 10) : 10;
    
    //  the page number the user wants to view (e.g page 1)
    const page = params["page"] ? parseInt(params["page"], 10) : 1;
    // the index of the first record on the current page
    const start = (page - 1) * limit;
    // the data for the current page
    const paginatedData = filteredData.slice(start, start + limit);

    sheetLog(ws, "Paginated Data: " + paginatedData);

    // Return data as JSON
    const jsonData = paginatedData.map(row => {
        let record = {};
        headers.forEach((header, index) => {
            record[header] = row[index];
        });
        return record;
    });

    sheetLog(ws, "Final JSON Data: " + JSON.stringify(jsonData));

    return sendJSON_({
        "status": "success",
        "code": 200,
        "data": jsonData,
        "pagination": {
            "currentPage": page,
            "totalPages": Math.ceil(filteredData.length / limit),
            "totalRecords": filteredData.length
        }
    });
}

