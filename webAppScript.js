function testabc() {
    const data = fetchUserTaskData();
    console.log('data==>' + JSON.stringify(data));
}
//------------------------------ Allowed Domains List ----------------------------------------
const allowedDomains = ['dpsnashik.in', 'dpsnagpurcity.com', 'dpsvaranasi.com', 'dpshinjawadi.com']; // Allowed domains

//------------------------------ index.html controllers ----------------------------------------
function doGet() {
    return HtmlService.createHtmlOutputFromFile('index').setTitle('Access Validation');
}

//-----------------------------check User Access------------------------------------------------
function checkUserAccess() {
    try {
        const userEmail = Session.getActiveUser().getEmail(); // Get current user's email
        if (!userEmail) {
            return { status: 'error', message: 'Unable to retrieve user email. Please try to login again.' };
        }

        const userDomain = userEmail.split('@')[1]; // Extract domain from email

        // Check if user's domain is allowed
        if (!allowedDomains.includes(userDomain)) {
            return {
                status: 'fail',
                message: `Access denied. Your domain (${userDomain}) is not authorized.`,
            };
        }
        ``
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
        if (!sheet) {
            return { status: 'error', message: 'Something went wrong.' };
        }

        const data = sheet.getDataRange().getValues();
        if (data.length === 0) {
            return { status: 'error', message: 'Something went wrong.' };
        }

        const headers = data[0];
        const emailIndex = headers.indexOf('Email');

        if (emailIndex === -1) {
            return { status: 'error', message: 'Something went wrong.' };
        }

        // Find user row
        const userRow = data.find((row, index) => index > 0 && row[emailIndex] === userEmail);

        if (userRow) {
            // Map headers to corresponding values
            const userData = {};
            headers.forEach((header, index) => {
                userData[header] = userRow[index];
            });

            const filterForProfile = {
                "ProfileId": userData["ProfileId"]
            }
            const profileData = fetchFilteredDataWithMap('Profiles', filterForProfile);

            if (profileData.length > 0) {
                userData["ProfileName"] = profileData[0]["Profile Name"];
            }

            createAndUpdateCache('userData', userData);

            return { status: 'success', message: `Welcome, ${userEmail}.` };
        } else {
            return { status: 'fail', message: 'You do not have the authority to access this app.' };
        }

    } catch (error) {
        return { status: 'error', message: `Unexpected error: ${error.message}` };
    }
}

function getDashboardHTML() {
    return HtmlService.createHtmlOutputFromFile('userDashboard').getContent();
}

//---------------------------------- Fetch User Response Data-----------------------------------------------

function fetchUsersResponseData() {
    try {
        const userData = getCacheValue('userData');
        const userId = userData['UserId'];
        const filter = {
            'Assigned To': userId
        }
        let filteredResponseData = getFilteredAndOptimizedData('Responses', filter);
        if (filteredResponseData != null || filteredResponseData != undefined) {
            return JSON.stringify({ status: 'success', data: filteredResponseData });
        } else {
            return JSON.stringify({ status: 'fail' });
        }
    } catch (error) {
        return JSON.stringify({ status: 'fail' });
    }
}

function getFilteredAndOptimizedData(sheetName, filters) {
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

//--------------------------------- Fetch Task Data---------------------------------------------------

function fetchUserTaskData() {
    try {
        const userData = getCacheValue('userData');
        const profileId = userData['ProfileId'];
        const filter = {
            'ProfileId': profileId
        }
        let filteredTaskData = getFilteredAndOptimizedData('Tasks', filter);
        if (filteredTaskData != null || filteredTaskData != undefined) {
            let tasksDataToReturn = {};
            filteredTaskData.forEach(task => {
                tasksDataToReturn[task["TaskId"]] = {
                    goal: task["Goal"],
                    taskName: task["Task Name"],
                    keyResult: task["Key Results"],
                    taskEvidence: task["Task Evidence"],
                    frequency: task["Frequency"]
                }
            });
            return JSON.stringify({ status: 'success', data: tasksDataToReturn });
        } else {
            return JSON.stringify({ status: 'fail' });
        }
    } catch (error) {
        console.log('error==>' + error);
        return JSON.stringify({ status: 'fail' });
    }
}

//---------------------------------- Generic Functions ------------------------------------------------------

//-------------------- Cache values set and retrieve----------------------------------
function createAndUpdateCache(key, valueObject, expiration = 600) {
    try {
        if (!key || !valueObject) {
            throw new Error('Invalid cache key or value');
        }
        const cache = CacheService.getScriptCache();
        const jsonString = JSON.stringify(valueObject);
        cache.put(key, jsonString, expiration);
    } catch (error) {
        console.error(`Cache Error: ${error.message}`);
    }
}

function getCacheValue(key) {
    try {
        if (!key) {
            throw new Error('Invalid cache key');
        }
        const cache = CacheService.getScriptCache();
        const jsonString = cache.get(key);
        return jsonString ? JSON.parse(jsonString) : null;
    } catch (error) {
        console.error(`Cache Read Error: ${error.message}`);
        return null;
    }
}

//---------------------------------- UserDashboard Controller ------------------------------------------------

function fetchUserName() {
    let userData = getCacheValue('userData');
    return userData['User Name'];
}

function fetchUserDataForProfile() {
    return getCacheValue('userData');
}

async function updateUserResponsesInSheet(recordData) {
    //return {status: "success", message: 'test success message'};
    const respData = trimResponseRecords(recordData);
    const result = await updateMultipleRecords(respData);
    return result;
}

function trimResponseRecords(records) {
    const allowedKeys = new Set(["ResponseId", "Progress Status", "Remarks", "Evidence"]);

    return records.map(record => {
        return Object.fromEntries(
            Object.entries(record).filter(([key, _]) => allowedKeys.has(key))
        );
    });
}

function updateMultipleRecords(records) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Responses");
    if (!sheet) {
        Logger.log("Sheet not found!");
        return { status: "failed", message: "Something went wrong..." };
    }

    var data = sheet.getDataRange().getValues(); // Read all data at once
    var headers = data[0]; // First row is headers
    var idIndex = headers.indexOf("ResponseId");

    if (idIndex === -1) {
        Logger.log("ResponseId column not found!");
        return { status: "failed", message: "Something went wrong..." };
    }

    var columnIndexes = {}; // Store column indexes for quick access
    headers.forEach((col, i) => columnIndexes[col] = i);

    var rowMap = {}; // Map ResponseId to row index
    for (var i = 1; i < data.length; i++) {
        rowMap[data[i][idIndex]] = i + 1; // Store actual row number
    }

    var updates = []; // Store only the necessary updates

    records.forEach(record => {
        var rowNum = rowMap[record.ResponseId];
        if (rowNum !== undefined) {
            Object.keys(record).forEach(key => {
                if (key !== "ResponseId" && columnIndexes[key] !== undefined) {
                    updates.push({
                        row: rowNum,
                        col: columnIndexes[key] + 1, // Convert index to 1-based column number
                        value: record[key]
                    });
                }
            });
        }
    });

    // **Batch update only specific cells**
    if (updates.length > 0) {
        var ranges = updates.map(update => sheet.getRange(update.row, update.col));
        var values = updates.map(update => [update.value]);
        ranges.forEach((range, i) => range.setValue(values[i][0]));
    }

    SpreadsheetApp.flush(); // Ensure updates happen immediately
    Logger.log("Updated " + updates.length + " specific cells successfully.");
    return { status: "success", message: "Tasks are updated..." }

}


function getUserTasks() {
    const email = Session.getActiveUser().getEmail();
    if (!email) throw new Error('No user logged in.');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const [uHdr, ...uRows] = ss.getSheetByName('Users').getDataRange().getValues();
    const eIdx = uHdr.indexOf('Email'),
        pIdx = uHdr.indexOf('ProfileId');
    const me = uRows.find(r => r[eIdx] === email);
    if (!me) return [];

    const myPid = me[pIdx];
    const [tHdr, ...tRows] = ss.getSheetByName('Tasks').getDataRange().getValues();
    const idx = name => tHdr.indexOf(name);

    return tRows
        .filter(r => r[idx('ProfileId')] === myPid)
        .map(r => ({
            taskId: r[idx('TaskId')],
            goal: r[idx('Goal')], // ← new
            frequency: r[idx('Frequency')],
            taskName: r[idx('Task Name')],
            keyResults: r[idx('Key Results')]
        }));
}