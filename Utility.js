function getDataToInfo() {
    let metadataSheet = getSheetInstance("MetaData");

    //sheet name json
    let databaseTables = {}
    databaseTables["userTable"] = {}
    databaseTables["profileTable"] = {}
    databaseTables["taskTable"] = {}
    databaseTables["responseTable"] = {}
        //databaseTables["wingTable"] = {}

    databaseTables["userTable"]["name"] = metadataSheet.getRange("B2").getValue();
    databaseTables["profileTable"]["name"] = metadataSheet.getRange("B3").getValue();
    databaseTables["taskTable"]["name"] = metadataSheet.getRange("B4").getValue();
    databaseTables["responseTable"]["name"] = metadataSheet.getRange("B5").getValue();
    //databaseTables["wingTable"] ["name"] = metadataSheet.getRange("B6").getValue();

    for (let table in databaseTables) {
        databaseTables[table]["range"] = getSheetBounds(databaseTables[table]["name"]);
    }

    return databaseTables;
}

function getSheetInstance(sheetName) {
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}

function getSheetBounds(sheetName) {
    let sheet = getSheetInstance(sheetName);

    let startingCell = sheet.getRange(1, 1); // Row 1, Column 1 (A2)

    let lastRow = sheet.getLastRow(); // Last row with data
    let lastColumn = sheet.getLastColumn(); // Last column with data
    let endingCell = sheet.getRange(lastRow, lastColumn);

    return startingCell.getA1Notation() + ':' + endingCell.getA1Notation();
}

// second Approach
function fetchFilteredDataWithMap(sheetName, filters) {
    let sheet = getSheetInstance(sheetName);
    let range = sheet.getDataRange();
    let headers = range.offset(0, 0, 1).getValues()[0]; // Only get the header row
    let data = range.offset(1, 0).getValues(); // Get all data without headers

    // Precompute column indices
    let filterIndices = Object.keys(filters).reduce((indices, key) => {
        let colIndex = headers.indexOf(key);
        if (colIndex === -1) throw new Error(`Field '${key}' not found`);
        indices[key] = colIndex;
        return indices;
    }, {});

    // Apply filters
    let filteredData = data.filter(row =>
        Object.entries(filterIndices).every(([key, colIndex]) => row[colIndex] === filters[key])
    );

    // Convert filtered rows to objects
    return filteredData.map(row =>
        headers.reduce((obj, header, i) => {
            obj[header] = row[i];
            return obj;
        }, {})
    );
}


function fetchAllSheetData(sheetName) {
    console.log('sheetname==>' + sheetName);
    let sheet = getSheetInstance(sheetName);
    let range = sheet.getDataRange();
    let values = range.getValues();

    if (values.length < 2) {
        throw new Error("Sheet is empty or only contains headers.");
    }

    let headers = values[0];
    let dataRows = values.slice(1);

    return dataRows.map(row =>
        headers.reduce((obj, header, index) => {
            obj[header] = row[index];
            return obj;
        }, {})
    );
}

function filterDataFromListObjects(dataList, filters) {
    if (!Array.isArray(dataList) || typeof filters !== "object") {
        throw new Error("Invalid input. Expecting an array of objects and a filters object.");
    }

    return dataList.filter(row =>
        Object.entries(filters).every(([key, value]) => row[key] === value)
    );
}

//-----------------------------Create Response Utility--------------------------------------------------
function createResponses(tasks, users) {
    // Validate inputs
    if (!Array.isArray(tasks) || !Array.isArray(users)) {
        throw new Error("Invalid input. Tasks, Profiles, and Users must be arrays of objects.");
    }

    let responses = []; // Array to store the generated response records
    let responseCounter = 1; // Counter to generate unique ResponseId

    // Iterate over the tasks
    tasks.forEach(task => {
        let profileId = task.ProfileId; // Get the ProfileId from the task

        // Find all users associated with this ProfileId
        let associatedUsers = users.filter(user => user.ProfileId === profileId);

        // Create a response record for each associated user
        associatedUsers.forEach(user => {
            let responseRecord = {
                "ResponseId": generateUniqueId(responseCounter++), // unique
                "Frequency": task.Frequency, // frequency
                "TaskId": task.TaskId,
                "Assigned To": user.UserId, // Map UserId to AssignedTo
                "Progress Status": "Not Started", // Default status
                "Remarks": "", // Initially empty
                "Evidence": "", // Initially empty
                "Start Date": new Date().toISOString().split('T')[0], // Start date is current date
                "Due Date": calculateDueDate(task.Frequency), // Calculate based on task frequency
                "Email Trigger": "None", // email trigger {None, Reminder, Escalation1, Escalation2}
                "Status": "Active" // Default active
            };

            responses.push(responseRecord); // Add the generated response to the array
        });
    });

    return responses; // Return the generated responses
}

// Helper function to calculate due dates based on frequency
function calculateDueDate(frequency) {
    let currentDate = new Date();
    switch (frequency.toLowerCase()) {
        case "weekly":
            currentDate.setDate(currentDate.getDate() + 7);
            break;
        case "monthly":
            currentDate.setMonth(currentDate.getMonth() + 1);
            break;
        case "quarterly":
            currentDate.setMonth(currentDate.getMonth() + 3);
            break;
        case "yearly":
            currentDate.setFullYear(currentDate.getFullYear() + 1); // Add 1 year
            break;
        case "fortnightly":
            currentDate.setDate(currentDate.getDate() + 14); // Add 14 days
            break;
        case "biyearly":
            currentDate.setMonth(currentDate.getMonth() + 6); // Add 6 months
            break;

        default:
            throw new Error(`Invalid frequency: ${frequency}`);
    }
    return currentDate.toISOString().split("T")[0]; // Return only the date part
}

// Helper function to generate a unique ID
function generateUniqueId(counter) {
    return `RESP-${counter}-${new Date().getTime()}`; // Unique format: RESP-<counter>-<timestamp>
}


//*****************************************************************************************************/

//---------------------------------Insert data Utility--------------------------------------------------
function insertResponses(responses) {
    if (!Array.isArray(responses) || responses.length === 0) {
        throw new Error("Invalid data. The responses array is empty or not an array.");
    }

    let sheet = getSheetInstance("Responses");
    if (!sheet) {
        throw new Error("The 'Responses' sheet does not exist.");
    }

    // Get headers from the sheet
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Convert responses to a 2D array
    let dataToInsert = responses.map(response =>
        headers.map(header => response[header] || "")
    );

    // Append the data to the sheet
    sheet.getRange(sheet.getLastRow() + 1, 1, dataToInsert.length, dataToInsert[0].length).setValues(dataToInsert);

    Logger.log(`${dataToInsert.length} response records successfully inserted.`);
}

//***********************************************************************************************************/



/*-------------------------------------------------------------------------------------------------------
// first approach
function fetchFilteredDataWithMap(sheetName, filters) {
  // Open the spreadsheet and get the sheet
  let sheet = getSheetInstance(sheetName);

  // Get all data (including headers)
  let data = sheet.getDataRange().getValues();
  let headers = data[0]; // Extract headers from the first row
  let rows = data.slice(1); // Data without headers

  // Filter rows based on the provided filters
  let filteredRows = rows.filter(row => {
    return Object.keys(filters).every(filterKey => {
      let columnIndex = headers.indexOf(filterKey); // Find column index for the filter key
      if (columnIndex === -1) {
        throw new Error(`Field '${filterKey}' not found in headers`);
      }
      return row[columnIndex] === filters[filterKey]; // Match value
    });
  });

  // Convert filtered rows to objects
  let filteredObjects = filteredRows.map(row => {
    return headers.reduce((obj, header, index) => {
      obj[header] = row[index];
      return obj;
    }, {});
  });

  return filteredObjects;
}

----------------------------------------------------------------------------------------------------------*/
function getRequests() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        if (!ss) return JSON.stringify({ status: 'fail', message: 'Spreadsheet not accessible.' });

        const sheet = ss.getSheetByName('Request');
        if (!sheet) return JSON.stringify({ status: 'success', data: [] });

        const values = sheet.getDataRange().getValues();
        if (values.length < 2) return JSON.stringify({ status: 'success', data: [] });

        const headers = values[0].map(h => h.toString().trim());
        const rows = values.slice(1);
        const result = rows.map(r => {
            const obj = {};
            headers.forEach((h, i) => obj[h] = (r[i] == null ? '' : r[i]));
            return obj;
        });

        return JSON.stringify({ status: 'success', data: result });
    } catch (e) {
        console.error('getRequests:', e);
        return JSON.stringify({ status: 'fail', message: String(e) });
    }
}

// returns only requests where RequestFrom === cached userId
function getMyRequests() {
    try {
        const user = getCacheValue('userData'); // your existing helper
        if (!user || !user.UserId) return JSON.stringify({ status: 'fail', message: 'User not available.' });

        const allResText = getRequests();
        const allRes = (typeof allResText === 'string') ? JSON.parse(allResText) : allResText;
        if (allRes.status !== 'success') return JSON.stringify({ status: 'success', data: [] });

        const filtered = allRes.data.filter(r => (r.RequestFrom || '').toString().trim() === user.UserId.toString().trim());
        return JSON.stringify({ status: 'success', data: filtered });

    } catch (e) {
        console.error('getMyRequests:', e);
        return JSON.stringify({ status: 'fail', message: String(e) });
    }
}