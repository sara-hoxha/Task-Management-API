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
