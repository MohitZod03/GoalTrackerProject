function updateSupervisors() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Users");
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // Dynamically determine the column indices
    const userIdCol = headers.indexOf("UserId");
    const emailCol = headers.indexOf("Email");
    const managerIdCol = headers.indexOf("ManagerId");
    const supervisorIdsCol = headers.indexOf("SupervisorIds");

    // Check if required columns are present
    if (
        userIdCol === -1 ||
        emailCol === -1 ||
        managerIdCol === -1 ||
        supervisorIdsCol === -1
    ) {
        Logger.log("Required columns not found in the sheet.");
        return;
    }

    const userMap = createUserMapData(data, userIdCol, emailCol); // Create UserId -> Email map

    // Get supervisor emails for all users
    const supervisorIds = getSupervisorsForAllUsers(
        data,
        userMap,
        userIdCol,
        managerIdCol
    );

    // Update the SupervisorIds in the sheet
    updateSupervisorIdsInSheet(sheet, supervisorIds, supervisorIdsCol);
}

/**
 * Creates a map with UserId as key and Email as value
 */
function createUserMapData(data, userIdCol, emailCol) {
    const userMap = new Map();
    for (let i = 1; i < data.length; i++) {
        const userId = data[i][userIdCol]; // UserId from dynamic column
        const email = data[i][emailCol]; // Email from dynamic column
        if (userId && email) {
            // Ensure valid UserId and Email
            userMap.set(userId, email);
        }
    }
    return userMap;
}

/**
 * Retrieves the supervisor emails for each user
 */
function getSupervisorsForAllUsers(data, userMap, userIdCol, managerIdCol) {
    const supervisorIds = [];

    // Loop through each user to get their supervisors
    for (let i = 1; i < data.length; i++) {
        const userId = data[i][userIdCol]; // UserId from the dynamic column
        let managerId = data[i][managerIdCol]; // ManagerId from the dynamic column
        let supervisors = [];

        Logger.log("Processing UserId: " + userId);

        // Traverse up the hierarchy for this user
        while (managerId) {
            const managerEmail = userMap.get(managerId);

            Logger.log(
                "ManagerId: " + managerId + " -> ManagerEmail: " + managerEmail
            );

            if (managerEmail) {
                supervisors.push(managerEmail);
            } else {
                Logger.log("No email found for ManagerId: " + managerId);
            }

            managerId = getManagerId(managerId, data, userIdCol, managerIdCol); // Get the next manager
        }

        Logger.log(
            "Supervisors for UserId " + userId + ": " + supervisors.join(", ")
        );
        supervisorIds.push(supervisors.join(", "));
    }

    return supervisorIds;
}

/**
 * Retrieves the ManagerId for a given userId
 */
function getManagerId(userId, data, userIdCol, managerIdCol) {
    for (let i = 1; i < data.length; i++) {
        if (data[i][userIdCol] === userId) {
            return data[i][managerIdCol]; // Return the ManagerId from dynamic column
        }
    }
    return null;
}

/**
 * Updates the SupervisorIds in the sheet
 */
function updateSupervisorIdsInSheet(sheet, supervisorIds, supervisorIdsCol) {
    for (let i = 0; i < supervisorIds.length; i++) {
        if (supervisorIds[i]) {
            sheet.getRange(i + 2, supervisorIdsCol + 1).setValue(supervisorIds[i]); // Add 1 for 1-based index
        }
    }
}