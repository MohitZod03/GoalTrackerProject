function testDeactivationForAlltypes() {
    weeklyDeactivateResponses();
}

// Frequency-specific functions
function yearlyDeactivateResponsesTypeA() { updateResponseStatus("Yearly", "Buffer", "Inactive", "A"); }

function yearlyDeactivateResponsesTypeB() { updateResponseStatus("Yearly", "Buffer", "Inactive", "B"); }

function yearlyBufferResponsesTypeA() { updateResponseStatus("Yearly", "Active", "Buffer", "A"); }

function yearlyBufferResponsesTypeB() { updateResponseStatus("Yearly", "Active", "Buffer", "B"); }

function quarterlyDeactivateResponsesTypeA() { updateResponseStatus("Quarterly", "Buffer", "Inactive", "A"); }

function quarterlyDeactivateResponsesTypeB() { updateResponseStatus("Quarterly", "Buffer", "Inactive", "B"); }

function quarterlyBufferResponsesTypeA() { updateResponseStatus("Quarterly", "Active", "Buffer", "A"); }

function quarterlyBufferResponsesTypeB() { updateResponseStatus("Quarterly", "Active", "Buffer", "B"); }

function monthlyDeactivateResponses() { updateResponseStatus("Monthly", "Buffer", "Inactive", "none"); }

function monthlyBufferResponses() { updateResponseStatus("Monthly", "Active", "Buffer", "none"); }

function fortnightlyDeactivateResponses() { updateResponseStatus("Fortnightly", "Buffer", "Inactive", "none"); }

function fortnightlyBufferResponses() { updateResponseStatus("Fortnightly", "Active", "Buffer", "none"); }

function weeklyDeactivateResponses() { updateResponseStatus("Weekly", "Buffer", "Inactive", "none"); }

function weeklyBufferResponses() { updateResponseStatus("Weekly", "Active", "Buffer", "none"); }

function halfYearlyDeactivateResponsesTypeA() { updateResponseStatus("Biyearly", "Buffer", "Inactive", "A"); }

function halfYearlyDeactivateResponsesTypeB() { updateResponseStatus("Biyearly", "Buffer", "Inactive", "B"); }

function halfYearlyBufferResponsesTypeA() { updateResponseStatus("Biyearly", "Active", "Buffer", "A"); }

function halfYearlyBufferResponsesTypeB() { updateResponseStatus("Biyearly", "Active", "Buffer", "B"); }

// Core reusable function
function updateResponseStatus(frequency, currentStatus, newStatus, schoolType) {
    try {
        const databaseTables = getDataToInfo();
        processResponses(databaseTables, frequency, currentStatus, newStatus, schoolType);
    } catch (error) {
        console.log("Error: " + error);
    }
}

function processResponses(databaseTables, freq, currentStatus, newStatus, schoolType) {
    console.log("database===>" + JSON.stringify(databaseTables));
    if (
        databaseTables ? .responseTable ? .name == null ||
        databaseTables ? .userTable ? .name == null ||
        freq == null ||
        currentStatus == null ||
        newStatus == null
    ) {
        throw new Error("Table name, frequency, current status, or new status is undefined or null.");
    }

    const sheet = getSheetInstance(databaseTables ? .responseTable ? .name);
    if (!sheet) {
        throw new Error("Sheet not found: " + databaseTables.responseTable.name);
    }

    const range = sheet.getDataRange(); // Get all data including headers
    const data = range.getValues(); // Fetch all data as a 2D array
    const headers = data[0]; // Assume the first row contains headers

    // Find indices for 'Status' and 'Frequency' columns
    const statusColumnIndex = headers.indexOf("Status");
    const frequencyColumnIndex = headers.indexOf("Frequency");
    const assignedToColumnIndex = headers.indexOf("Assigned To");

    if (statusColumnIndex === -1 || frequencyColumnIndex === -1) {
        throw new Error("'Status' or 'Frequency' column not found in headers.");
    }

    let rowsToUpdate = []; // Keep track of rows that need to be updated

    const userSheetName = databaseTables.userTable.name;
    let userList;
    let userSet;
    let schoolCondition = false;

    if (freq == 'Yearly' || freq == 'Biyearly' || freq == 'Quarterly') {
        schoolCondition = true;
        let filter = {
            "Type": schoolType
        }
        userData = fetchFilteredDataWithMap(userSheetName, filter);
        userList = userData.map(user => user.UserId);
        userSet = new Set(userList);
    }

    // Iterate over data starting from the second row (data rows)
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const status = row[statusColumnIndex];
        const frequency = row[frequencyColumnIndex];
        const assignedTo = row[assignedToColumnIndex];

        if (status === currentStatus && frequency === freq && (!schoolCondition || userSet.has(assignedTo))) {
            rowsToUpdate.push(i + 1); // Add 1-based index for Sheets API
        }
    }

    // Update only the relevant rows
    rowsToUpdate.forEach((rowNumber) => {
        sheet.getRange(rowNumber, statusColumnIndex + 1).setValue(newStatus); // Update 'Status'
    });

    console.log(`Updated ${rowsToUpdate.length} rows from ${currentStatus} to ${newStatus} for frequency ${freq}.`);
}