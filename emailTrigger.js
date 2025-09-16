function testEmailTriggerFunction() {
    weeklyEmailReminderTrigger();
    //testEmailSend();
}

function testEmailSend() {
    const responseRecords = [
        { "Email": "bambalyash@gmail.com", "User Name": "Test 1", "ManagerEmail": "aiproject@dpsnashik.in", "Frequency": "Weekly", "Tasks": "Task Weekly 1; Task Weekly 2" },
    ];
    sendDynamicEmailNotifications(responseRecords, 'escalation');
}

/****************************** Email Triggers *************************** */
async function combinedYearlyReminderTrigger() {
    await yearlyEmailReminderTriggerTypeA();
    await yearlyEmailReminderTriggerTypeB();
}
async function combinedHalfYearlyReminderTrigger() {
    await halfYearlyReminderTriggerTypeA();
    await halfYearlyReminderTriggerTypeB();
}
async function combinedQuarterlyReminderTrigger() {
    await quarterlyEmailReminderTriggerTypeA();
    await quarterlyEmailReminderTriggerTypeB();
}
async function yearlyEmailReminderTriggerTypeA() { await emailTriggerStart("Yearly", "reminder", "A"); }
async function yearlyEmailReminderTriggerTypeB() { await emailTriggerStart("Yearly", "reminder", "B"); }
async function quarterlyEmailReminderTriggerTypeA() { await emailTriggerStart("Quarterly", "reminder", "A"); }
async function quarterlyEmailReminderTriggerTypeB() { await emailTriggerStart("Quarterly", "reminder", "B"); }

function monthlyEmailReminderTrigger() { emailTriggerStart("Monthly", "reminder", "none"); }

function fortnightlyEmailReminderTrigger() { emailTriggerStart("Fortnightly", "reminder", "none"); }

function weeklyEmailReminderTrigger() { emailTriggerStart("Weekly", "reminder", "none"); }
async function halfYearlyReminderTriggerTypeA() { await emailTriggerStart("Biyearly", "reminder", "A"); }
async function halfYearlyReminderTriggerTypeB() { await emailTriggerStart("Biyearly", "reminder", "B"); }

async function combinedYearlyEscalation1Trigger() {
    await yearlyEscalation1TriggerTypeA();
    await yearlyEscalation1TriggerTypeB();
}
async function combinedHalfYearlyEscalation1Trigger() {
    await halfYearlyEscalation1TriggerTypeA();
    await halfYearlyEscalation1TriggerTypeB();
}
async function combinedQuarterlyEscalation1Trigger() {
    await quarterlyEscalation1TriggerTypeA();
    await quarterlyEscalation1TriggerTypeB();
}
async function yearlyEscalation1TriggerTypeA() { await emailTriggerStart("Yearly", "escalation1", "A"); }
async function yearlyEscalation1TriggerTypeB() { await emailTriggerStart("Yearly", "escalation1", "B"); }
async function quarterlyEscalation1TriggerTypeA() { await emailTriggerStart("Quarterly", "escalation1", "A"); }
async function quarterlyEscalation1TriggerTypeB() { await emailTriggerStart("Quarterly", "escalation1", "B"); }

function monthlyEscalation1Trigger() { emailTriggerStart("Monthly", "escalation1", "none"); }

function fortnightlyEscalation1Trigger() { emailTriggerStart("Fortnightly", "escalation1", "none"); }

function weeklyEscalation1Trigger() { emailTriggerStart("Weekly", "escalation1", "none"); }
async function halfYearlyEscalation1TriggerTypeA() { await emailTriggerStart("Biyearly", "escalation1", "A"); }
async function halfYearlyEscalation1TriggerTypeB() { await emailTriggerStart("Biyearly", "escalation1", "B"); }

async function combinedYearlyEscalation2Trigger() {
    await yearlyEscalation2TriggerTypeA();
    await yearlyEscalation2TriggerTypeB();
}
async function combinedHalfYearlyEscalation2Trigger() {
    await halfYearlyEscalation2TriggerTypeA();
    await halfYearlyEscalation2TriggerTypeB();
}
async function combinedQuarterlyEscalation2Trigger() {
    await quarterlyEscalation2TriggerTypeA();
    await quarterlyEscalation2TriggerTypeB();
}
async function yearlyEscalation2TriggerTypeA() { await emailTriggerStart("Yearly", "escalation2", "A"); }
async function yearlyEscalation2TriggerTypeB() { await emailTriggerStart("Yearly", "escalation2", "B"); }
async function quarterlyEscalation2TriggerTypeA() { await emailTriggerStart("Quarterly", "escalation2", "A"); }
async function quarterlyEscalation2TriggerTypeB() { await emailTriggerStart("Quarterly", "escalation2", "B"); }

function monthlyEscalation2Trigger() { emailTriggerStart("Monthly", "escalation2", "none"); }

function fortnightlyEscalation2Trigger() { emailTriggerStart("Fortnightly", "escalation2", "none"); }

function weeklyEscalation2Trigger() { emailTriggerStart("Weekly", "escalation2", "none"); }
async function halfYearlyEscalation2TriggerTypeA() { await emailTriggerStart("Biyearly", "escalation2", "A"); }
async function halfYearlyEscalation2TriggerTypeB() { await emailTriggerStart("Biyearly", "escalation2", "B"); }

/************************** Main Trigger Function ************************** */
// Define constants
const validStatus = ["done", "na"];

// Configuration for CC email ranges
const CC_EMAIL_CONFIG = {
    reminder: {
        startIndex: 0,
        endIndex: 1 // Will include managers from index 0 to 2 (inclusive)
    },
    escalation1: {
        startIndex: 0,
        endIndex: 2 // Will include managers from index 0 to 3 (inclusive)
    },
    escalation2: {
        startIndex: 0,
        endIndex: -1 // -1 indicates to include all managers
    }
};

function emailTriggerStart(frequency, notificationType, schoolType) {
    try {
        const databaseTables = getDataToInfo();
        const statusFilter = getStatusFilter(notificationType); // Get status filter based on notification type
        const emailTriggerObj = generateResponseRecords(databaseTables, frequency, notificationType, statusFilter, schoolType);
        console.log('Generated data:', JSON.stringify(emailTriggerObj));
        sendDynamicEmailNotifications(emailTriggerObj, notificationType);
    } catch (error) {
        console.error('Error:', error);
    }
}

/******************************* Helper Functions ******************************* */
// Determine the status filter based on the notification type
function getStatusFilter(notificationType) {
    switch (notificationType) {
        case "reminder":
            return "Active"; // Reminder emails for Active status
        case "escalation1":
            return "Buffer"; // Escalation 1 emails for Buffer status
        case "escalation2":
            return "Buffer"; // Escalation 2 emails for Buffer status
        default:
            throw new Error(`Invalid notification type: ${notificationType}`);
    }
}

// Generate response records based on status filter
function generateResponseRecords(databaseTables, frequency, notificationType, statusFilter, schoolType) {
    validateInputs(databaseTables, frequency);

    const responseSheetName = databaseTables.responseTable.name;
    const userSheetName = databaseTables.userTable.name;
    const tasksSheetName = databaseTables.taskTable.name;

    console.log("database===>", JSON.stringify(databaseTables));

    // Fetch all data from the sheets
    let responseData;
    let userData;
    let tasksData;
    let schoolCondition = false;
    let userList = null;
    if (frequency == 'Yearly' || frequency == 'Biyearly' || frequency == 'Quarterly') {
        schoolCondition = true;
        let filter = {
            "Type": schoolType
        }
        userData = fetchFilteredDataWithMap(userSheetName, filter);
        userList = userData.map(user => user.UserId);
        console.log('school condition==>' + schoolCondition);
        console.log("userlist===>" + userList);
    } else {
        userData = fetchAllSheetData(userSheetName);
    }
    responseData = fetchAllSheetData(responseSheetName);
    tasksData = fetchAllSheetData(tasksSheetName);


    // Create user and task maps for efficient lookups
    const userMap = createUserMap(userData);
    const taskMap = createTaskMap(tasksData);

    // Filter responses
    const filteredResponses = filterResponses(responseData, frequency, statusFilter, validStatus, schoolCondition, userList);

    // Update the Escalation or Reminder column for filtered tasks
    updateResponseRecords(responseSheetName, filteredResponses, notificationType);

    // Group tasks by user
    const userTasksMap = groupTasksByUser(filteredResponses, taskMap);

    // Build response records
    return buildResponseRecords(userTasksMap, userMap, notificationType);
}

// Validate input parameters
function validateInputs(databaseTables, frequency) {
    if (!databaseTables ? .responseTable ? .name ||
        !databaseTables ? .userTable ? .name ||
        !databaseTables ? .taskTable ? .name ||
        !frequency
    ) {
        throw new Error("Table names or frequency are undefined or null.");
    }
}

// Create a user map for efficient UserId -> User object lookup
function createUserMap(userData) {
    return userData.reduce((map, user) => {
        map[user.UserId] = user;
        return map;
    }, {});
}

// Create a task map for efficient TaskId -> Task Name lookup
function createTaskMap(tasksData) {
    return tasksData.reduce((map, task) => {
        map[task.TaskId] = task["Task Name"] + ' ( ' + task['Key Results'] + ' )'; // Assuming 'Task Name' column exists
        return map;
    }, {});
}

// Optimized function using Set for faster lookup
function filterResponses(responseData, frequency, statusFilter, validStatus, schoolCondition, userList) {
    const userSet = new Set(userList); // Convert array to Set for O(1) lookups

    return responseData.filter(response =>
        response.Status === statusFilter &&
        response.Frequency === frequency &&
        !validStatus.includes(response["Progress Status"] ? .toLowerCase()) &&
        (!schoolCondition || userSet.has(response["Assigned To"])) // O(1) lookup
    );
}

function updateResponseRecords(responseSheetName, filteredResponses, notificationType) {
    const sheet = getSheetInstance(responseSheetName);

    if (!sheet) {
        throw new Error(`Sheet ${responseSheetName} not found.`);
    }

    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();

    const columnName = "Email Trigger";

    // Get the index of the Escalation or Reminder column
    const headers = data[0].map(header => header.toLowerCase()); // Normalize header case
    const notificationColumnIndex = headers.indexOf(columnName.toLowerCase());

    if (notificationColumnIndex === -1) {
        throw new Error(`${columnName} column not found in ${responseSheetName}.`);
    }

    // Create a map for row lookups
    const idToRowIndexMap = data.reduce((map, row, index) => {
        if (index > 0) { // Skip header row
            map[row[0]] = index; // Assuming 'Id' is in the first column
        }
        return map;
    }, {});

    // Prepare batch updates
    const updates = [];
    filteredResponses.forEach(response => {
        const rowIndex = idToRowIndexMap[response.ResponseId];
        if (rowIndex !== undefined) {
            data[rowIndex][notificationColumnIndex] = notificationType; // Update the data array
            updates.push(rowIndex + 1); // Track updated row index (1-based for sheet)
        } else {
            console.warn(`Response ID ${response.ResponseId} not found in the sheet.`);
        }
    });

    // Write updated data back to the sheet
    if (updates.length > 0) {
        const updateRange = sheet.getRange(2, 1, data.length - 1, data[0].length); // Skip header row
        updateRange.setValues(data.slice(1)); // Exclude headers for writing
    }
}

// Group tasks by user, using task names
function groupTasksByUser(filteredResponses, taskMap) {
    const userTasksMap = {};

    filteredResponses.forEach(response => {
        const userId = response["Assigned To"]; // Use 'Assigned To' to identify users
        const taskId = response["TaskId"]; // Use 'TaskId' to look up task name
        const responseId = response["ResponseId"]; // Get ResponseId to include with task
        const progressStatus = response["Progress Status"]; // Get Progress Status for task

        if (!userTasksMap[userId]) {
            userTasksMap[userId] = {
                UserId: userId,
                Frequency: response["Frequency"], // Ensure consistent frequency value
                Tasks: [],
                ResponseIds: [], // Store ResponseIds corresponding to tasks
                ProgressStatuses: [], // Store Progress Statuses corresponding to tasks
            };
        }

        const taskName = taskMap[taskId];
        if (taskName) {
            userTasksMap[userId].Tasks.push(taskName);
            userTasksMap[userId].ResponseIds.push(responseId);
            userTasksMap[userId].ProgressStatuses.push(progressStatus); // Store Progress Status
        }
    });

    return userTasksMap;
}

// Build response records with templates for different notification types
function buildResponseRecords(userTasksMap, userMap, notificationType) {
    const isEscalation1 = notificationType === "escalation1";
    const isEscalation2 = notificationType === "escalation2";

    return Object.values(userTasksMap)
        .map(record => {
            const user = userMap[record.UserId];

            if (user) {
                // Create an array of tasks with their corresponding ResponseIds and Progress Status
                const tasksWithIds = record.Tasks.map((task, index) => {
                    // Try to get Progress Status from the filteredResponses if available
                    // We'll assume record.ProgressStatuses exists and is aligned with Tasks/ResponseIds
                    const progressStatus = record.ProgressStatuses && record.ProgressStatuses[index] ? record.ProgressStatuses[index] : '';
                    // Format: [Status: ...] Task Name [Task Id: ...]
                    return `[Status: ${progressStatus}] ${task} [Task Id: ${record.ResponseIds[index]}]`;
                });

                return {
                    Email: user.Email,
                    "User Name": user["User Name"], // Assuming Users sheet has 'User Name'
                    ManagerEmail: user["SupervisorIds"], // Assuming SupervisorIds contains manager emails
                    Frequency: record.Frequency,
                    Tasks: tasksWithIds.join("; ") // Join tasks with ResponseIds and Progress Status
                };
            }

            return null; // Return null if user not found (optional)
        })
        .filter(record => record !== null); // Filter out null values
}

function sendDynamicEmailNotifications(responseRecords, notificationType) {
    if (!Array.isArray(responseRecords) || responseRecords.length === 0) {
        console.log("No response records provided for email notifications.");
        return;
    }

    // Validate notificationType
    if (!["reminder", "escalation1", "escalation2"].includes(notificationType)) {
        console.error(`Invalid notificationType: ${notificationType}. Must be 'reminder' or 'escalation'.`);
        return;
    }

    responseRecords.forEach(record => {
        try {
            const userEmail = record["Email"]; // User's email
            const userName = record["User Name"]; // User's name
            const managerEmails = record["ManagerEmail"] || ""; // Manager email(s)
            const frequency = record["Frequency"]; // Task frequency
            const tasks = record["Tasks"] || ""; // Tasks field

            // If user's email or frequency is missing, log and skip this record
            if (!userEmail || !frequency) {
                console.warn(`Missing required fields for record: ${JSON.stringify(record)}. Skipping.`);
                return;
            }

            // Process manager emails based on notification type
            let ccEmails = [];
            if (managerEmails) {
                const managerEmailArray = managerEmails
                    .split(",")
                    .map(email => email.trim())
                    .filter(email => email); // Filter out empty or invalid emails

                const config = CC_EMAIL_CONFIG[notificationType];
                if (config.endIndex === -1) {
                    // For escalation2, include all manager emails
                    ccEmails = managerEmailArray;
                } else {
                    // For reminder and escalation1, use the configured range
                    // Ensure we don't exceed array bounds
                    const endIndex = Math.min(config.endIndex, managerEmailArray.length - 1);
                    ccEmails = managerEmailArray.slice(config.startIndex, endIndex + 1);
                }
            }

            // Format tasks as point-wise list
            const taskListHtml = tasks
                .split(";")
                .map(task => `<li>${task.trim()}</li>`) // Create list items for each task
                .join("");

            const taskListText = tasks
                .split(";")
                .map(task => `- ${task.trim()}`) // Create bullet points for text email
                .join("\n");

            const subject = notificationType === "reminder" ?
                `Reminder: Tasks (${frequency})` :
                `Escalation: Tasks (${frequency})`;

            const frequencyText = frequencyTextForMail(frequency);

            const bodyHtml = notificationType === "reminder" ?
                `
        <p>Dear ${userName},</p>
        <p>This is a reminder for your following ${frequency} tasks that are pending for current ${frequencyText}.</p>
        <ul>${taskListHtml}</ul>
        <p>Please make sure to update the progress as soon as possible.</p>
        <p>Best Regards,<br/>Team DPS</p>
        <P>Note: Kindly ignore, if already done.</p>
      ` :
                `
        <p>Dear ${userName},</p>
        <p>This email serves as an escalation regarding the following ${frequency} tasks that are pending for last ${frequencyText}.</p>
        <ul>${taskListHtml}</ul>
        <p>Kindly ensure that the progress on these tasks is updated at the earliest to avoid any delays. Your prompt attention to this matter will be greatly appreciated.</p>
        <p>Best Regards,<br/>Team DPS</p>
        <P>Note: Kindly ignore, if already done.</p>
        `;

            const bodyText = notificationType === "reminder" ?
                `
        Dear ${userName},

        This is a reminder for your following ${frequency} tasks that are pending for current ${frequencyText}.
        ${taskListText}

        Please make sure to update the progress as soon as possible.

        Best Regards,
        Team PDS

        Note: Kindly ignore, if already done.
      ` :
                `
        Dear ${userName},

        This email serves as an escalation regarding the following ${frequency} tasks that are pending for last ${frequencyText}.
        ${taskListText}

        Kindly ensure that the progress on these tasks is updated at the earliest to avoid any delays. Your prompt attention to this matter will be greatly appreciated.

        Best Regards,
        Team PDS

        Note: Kindly ignore, if already done.
      `;

            // Send email
            GmailApp.sendEmail(userEmail, subject, bodyText, {
                htmlBody: bodyHtml,
                cc: ccEmails.join(","), // Join CC emails into a comma-separated string
            });

            console.log(`Email sent to ${userEmail} with CC to ${ccEmails.join(", ") || "none"}`);
        } catch (error) {
            console.error(`Failed to send email for record: ${JSON.stringify(record)}. Error: ${error.message}`);
        }
    });
}

function frequencyTextForMail(frequency) {
    switch (frequency) {
        case "Weekly":
            return "week";
        case "Fortnightly":
            return "fortnight";
        case "Monthly":
            return "month";
        case "Quarterly":
            return "quarter";
        case "Half Yearly":
        case "Biyearly": // Handle both synonyms
            return "half year";
        case "Yearly":
            return "year";
        default:
            return "unknown frequency"; // Fallback for unrecognized values
    }
}

/******************************************************************************************************************************* */

// Replace existing sendAccessRequestEmail with this function.
// Only change: remove the human-readable "Task ID / Task Name / Key Results" block from the email body.
// Everything else and validations remain unchanged.
function sendAccessRequestEmail(payload) {
    try {
        var selectedTasks = [];
        var changesPayload = null;

        if (payload === undefined || payload === null) {
            return { status: 'fail', message: 'No payload provided' };
        }

        if (Array.isArray(payload)) {
            selectedTasks = payload;
        } else {
            selectedTasks = payload.tasks || payload.data || [];
            changesPayload = payload.changes || payload.request || null;
        }

        if (typeof changesPayload === 'string' && changesPayload) {
            try { changesPayload = JSON.parse(changesPayload); } catch (e) { /* leave as-is */ }
        }

        if (!selectedTasks || selectedTasks.length === 0) {
            return { status: 'fail', message: 'No tasks provided.' };
        }

        var userData = null;
        try { userData = getCacheValue('userData'); } catch (e) { userData = null; }
        if (!userData || !userData.Email || !userData.UserId) {
            return { status: 'fail', message: 'User info missing. Reload app.' };
        }

        var myTaskIds = new Set((getUserTasks() || []).map(function(t) { return (t.taskId || t.TaskId || '').toString(); }));
        var valid = selectedTasks.filter(function(t) {
            return myTaskIds.has((t.taskId || t.TaskId || '').toString());
        });
        if (!valid || valid.length === 0) {
            return { status: 'fail', message: 'Selected tasks not valid for this user.' };
        }

        // Resolve supervisor list (can be emails or UserIds)
        var supField = userData['SupervisorIds'] || '';
        if (!supField && typeof fetchFilteredDataWithMap === 'function') {
            var rec = (fetchFilteredDataWithMap('Users', { Email: userData['Email'] }) || [])[0];
            supField = rec ? (rec['SupervisorIds'] || '') : '';
        }
        if (!supField) {
            return { status: 'fail', message: 'Supervisor not set.' };
        }

        // Build array of supervisor identifiers (trim, filter empty)
        var first = supField.split(',').map(function(s) { return s.trim(); }).filter(Boolean);
        if (!first.length) return { status: 'fail', message: 'Supervisor list empty after parsing.' };

        // Resolve each identifier to a user record when possible
        var supRecords = [];
        first.forEach(function(id) {
            try {
                var resolved = null;
                if (id.indexOf('@') !== -1) {
                    // looks like email
                    resolved = (typeof fetchFilteredDataWithMap === 'function') ?
                        (fetchFilteredDataWithMap('Users', { Email: id }) || [])[0] || { Email: id, UserId: id } :
                        { Email: id, UserId: id };
                } else {
                    // probably a UserId
                    resolved = (typeof fetchFilteredDataWithMap === 'function') ?
                        (fetchFilteredDataWithMap('Users', { UserId: id }) || [])[0] || { UserId: id } :
                        { UserId: id };
                }
                // ensure we at least have an Email or UserId
                if (resolved) {
                    // If resolved object lacks Email but id looked like an email, populate it
                    if ((!resolved.Email || resolved.Email === '') && id.indexOf('@') !== -1) resolved.Email = id;
                    // If resolved lacks UserId, set to email fallback
                    if ((!resolved.UserId || resolved.UserId === '') && resolved.Email) resolved.UserId = resolved.Email;
                    supRecords.push(resolved);
                }
            } catch (e) {
                // ignore resolution error but keep fallback
                supRecords.push({ Email: id, UserId: id });
            }
        });

        if (!supRecords.length) {
            return { status: 'fail', message: 'Supervisor ID/email(s) not resolved.' };
        }

        // Primary is first in hierarchy; rest are CC
        var primary = supRecords[0];
        var primaryEmail = primary.Email || '';
        var primaryUserId = primary.UserId || primaryEmail || first[0];

        // Build CC emails list from the rest (keep any raw emails/records)
        var ccRecords = supRecords.slice(1);
        var ccEmails = ccRecords.map(function(r) { return r.Email; }).filter(function(e) { return !!e; });

        // Build list of requestToIds to save (comma-separated UserIds or fallback to emails)
        var requestToIds = primary.UserId ? primary.UserId : '';
        var supName = primary['User Name'] || (primaryEmail ? primaryEmail.split('@')[0] : primaryUserId);
        var userName = userData['User Name'] || userData['Email'];
        var subject = 'Request to Grant Edit Access — ' + userName;

        // Helper to extract field-old/new pairs for a given taskId.
        function extractChangesForTask(changes, taskId) {
            var out = [];
            if (!changes) return out;
            if (Array.isArray(changes)) {
                changes.forEach(function(entry) {
                    if (entry && entry[taskId]) {
                        var inner = entry[taskId];
                        var map = {};
                        for (var k in inner) {
                            if (!inner.hasOwnProperty(k)) continue;
                            var m = k.match(/^(.*)_(old|new)$/i);
                            if (m) {
                                var base = m[1];
                                map[base] = map[base] || { old: '', new: '' };
                                if (/old$/i.test(k)) map[base].old = inner[k];
                                if (/new$/i.test(k)) map[base].new = inner[k];
                            } else {
                                map[k] = map[k] || { old: '', new: '' };
                                map[k].new = inner[k];
                            }
                        }
                        for (var f in map) {
                            out.push({ field: f, oldVal: map[f].old, newVal: map[f].new });
                        }
                    }
                });
                return out;
            }
            if (typeof changes === 'object') {
                if (changes[taskId]) {
                    var inner2 = changes[taskId];
                    var map2 = {};
                    for (var k2 in inner2) {
                        if (!inner2.hasOwnProperty(k2)) continue;
                        var m2 = k2.match(/^(.*)_(old|new)$/i);
                        if (m2) {
                            var base2 = m2[1];
                            map2[base2] = map2[base2] || { old: '', new: '' };
                            if (/old$/i.test(k2)) map2[base2].old = inner2[k2];
                            if (/new$/i.test(k2)) map2[base2].new = inner2[k2];
                        } else {
                            map2[k2] = map2[k2] || { old: '', new: '' };
                            map2[k2].new = inner2[k2];
                        }
                    }
                    for (var ff in map2) {
                        out.push({ field: ff, oldVal: map2[ff].old, newVal: map2[ff].new });
                    }
                    return out;
                }
                if (Array.isArray(changes.request)) {
                    return extractChangesForTask(changes.request, taskId);
                }
                for (var k3 in changes) {
                    if (!changes.hasOwnProperty(k3)) continue;
                    try {
                        if (k3 === taskId && typeof changes[k3] === 'object') {
                            var inner3 = changes[k3];
                            var map3 = {};
                            for (var p in inner3) {
                                if (!inner3.hasOwnProperty(p)) continue;
                                var m3 = p.match(/^(.*)_(old|new)$/i);
                                var base3 = m3 ? m3[1] : p;
                                map3[base3] = map3[base3] || { old: '', new: '' };
                                if (m3 && /old$/i.test(p)) map3[base3].old = inner3[p];
                                else if (m3 && /new$/i.test(p)) map3[base3].new = inner3[p];
                                else map3[base3].new = inner3[p];
                            }
                            for (var q in map3) out.push({ field: q, oldVal: map3[q].old, newVal: map3[q].new });
                            return out;
                        }
                    } catch (e) { /* ignore */ }
                }
            }
            return out;
        }

        // Build list of only those selected tasks that have changes
        var tasksWithChanges = [];
        valid.forEach(function(task) {
            var id = (task.taskId || task.TaskId || '').toString();
            var ch = extractChangesForTask(changesPayload, id);
            var meaningful = ch.filter(function(x) { return (x.oldVal !== undefined && x.oldVal !== '') || (x.newVal !== undefined && x.newVal !== ''); });
            if (meaningful.length) tasksWithChanges.push({ id: id, changes: meaningful });
        });

        if (!tasksWithChanges.length) {
            return { status: 'fail', message: 'No changes found for selected tasks.' };
        }

        // Build plain-text body: bullet-separated tasks, only TaskID and field_Old/field_New lines
        var plain = [];
        plain.push('Dear ' + supName + ',\n');
        plain.push(userName + ' is requesting edit access for the following task(s):\n');

        tasksWithChanges.forEach(function(t) {
            plain.push('• Task ID: ' + t.id);
            t.changes.forEach(function(c) {
                plain.push('  ' + c.field + '_Old: ' + (c.oldVal === undefined ? '' : c.oldVal));
                plain.push('  ' + c.field + '_New: ' + (c.newVal === undefined ? '' : c.newVal));
            });
            plain.push(''); // blank line between tasks
        });

        plain.push('Kindly review the request and grant the necessary permissions.\n');
        plain.push('Best Regards,');
        plain.push('Goal Tracker\n');
        plain.push('Note: This is a system-generated email, please do not reply directly.');

        var bodyText = plain.join('\n');

        // Build a minimal HTML body so Task ID appears bold for HTML-capable clients
        var html = [];
        html.push('<p>Dear ' + escapeHtmlForHtml(supName) + ',</p>');
        html.push('<p><strong>' + escapeHtmlForHtml(userName) + '</strong> is requesting edit access for the following task(s):</p>');

        tasksWithChanges.forEach(function(t) {
            html.push('<p>• <strong>Task ID: ' + escapeHtmlForHtml(t.id) + '</strong></p>');
            html.push('<ul style="margin-top:0;margin-bottom:8px;padding-left:18px;">');
            t.changes.forEach(function(c) {
                html.push('<li>' + escapeHtmlForHtml(c.field + '_Old: ' + (c.oldVal === undefined ? '' : c.oldVal)) + '</li>');
                html.push('<li>' + escapeHtmlForHtml(c.field + '_New: ' + (c.newVal === undefined ? '' : c.newVal)) + '</li>');
            });
            html.push('</ul>');
        });

        html.push('<p>Kindly review the request and grant the necessary permissions.</p>');
        html.push('<p>Best Regards,<br/>Goal Tracker</p>');
        html.push('<p><em>Note: This is a system-generated email, please do not reply directly.</em></p>');

        var bodyHtml = html.join('\n');

        // Send email with primary as TO and rest as CC
        if (!primaryEmail || primaryEmail === '') {
            // If primary doesn't have an email, try to fallback to first identifier that looks like an email
            var fallback = first.find(function(s) { return s.indexOf('@') !== -1; }) || primaryUserId;
            primaryEmail = fallback;
        }
        GmailApp.sendEmail(primaryEmail, subject, bodyText, { htmlBody: bodyHtml, cc: ccEmails.join(',') });

        // Save request JSON to sheet (stringify changesPayload if present) - requestTo will be comma-separated UserIds/emails
        var changesJsonString = '';
        try {
            changesJsonString = (changesPayload && typeof changesPayload !== 'string') ? JSON.stringify(changesPayload) : String(changesPayload || '');
        } catch (e) {
            changesJsonString = String(changesPayload || '');
        }
        var saveStatus = saveAccessRequest(userData.UserId, requestToIds, changesJsonString);

        var ccLog = ccEmails.length ? (' with CC to ' + ccEmails.join(', ')) : '';
        if (saveStatus && saveStatus.ok) {
            return { status: 'success', message: 'Request sent to ' + primaryEmail + ccLog + ' and saved (' + saveStatus.msg + ').', requestId: saveStatus.msg };
        } else if (saveStatus && !saveStatus.ok) {
            return { status: 'success', message: 'Email sent to ' + primaryEmail + ccLog + ' but saving request failed: ' + saveStatus.msg };
        } else {
            return { status: 'success', message: 'Request sent to ' + primaryEmail + ccLog + '.' };
        }

    } catch (e) {
        console.error('sendAccessRequestEmail:', e);
        return { status: 'fail', message: 'Error sending email: ' + (e.message || e) };
    }
}
// safe escape helper for HTML
function escapeHtmlForHtml(str) {
    if (str === null || str === undefined) return '';
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}



// Save function (unchanged except it will store the JSON string passed as requestTask)
function saveAccessRequest(requestFromId, requestToId, requestTask) {
    var status = { ok: true, msg: '' };
    try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        if (!ss) throw new Error('Spreadsheet not accessible.');

        var sheetName = 'Request';
        var reqSheet = ss.getSheetByName(sheetName);

        // Minimal required headers
        var baseHeaders = ['RequestId', 'RequestFrom', 'RequestTo', 'RequestTask', 'Status', 'DataChanged'];

        // Create sheet with base headers if missing
        if (!reqSheet) {
            reqSheet = ss.insertSheet(sheetName);
            reqSheet.getRange(1, 1, 1, baseHeaders.length).setValues([baseHeaders]);
        }

        // Read headers and build map
        var finalLastCol = Math.max(1, reqSheet.getLastColumn());
        var finalHeaders = reqSheet.getRange(1, 1, 1, finalLastCol).getValues()[0] || [];
        var headerMap = {};
        for (var i = 0; i < finalHeaders.length; i++) {
            var key = (finalHeaders[i] || '').toString().trim();
            if (key) headerMap[key.toLowerCase()] = i;
        }

        // Ensure DataChanged header exists (append it if not)
        if (headerMap['datachanged'] === undefined) {
            reqSheet.insertColumnAfter(reqSheet.getLastColumn());
            reqSheet.getRange(1, reqSheet.getLastColumn()).setValue('DataChanged');
            // refresh headers and map
            finalLastCol = Math.max(1, reqSheet.getLastColumn());
            finalHeaders = reqSheet.getRange(1, 1, 1, finalLastCol).getValues()[0] || [];
            headerMap = {};
            for (var i = 0; i < finalHeaders.length; i++) {
                var k = (finalHeaders[i] || '').toString().trim();
                if (k) headerMap[k.toLowerCase()] = i;
            }
        }

        // Normalize requestTask into JS object/array/string
        var parsed = null;
        var wasString = false;
        if (requestTask === undefined || requestTask === null) {
            parsed = null;
        } else if (typeof requestTask === 'string') {
            wasString = true;
            try { parsed = JSON.parse(requestTask); } catch (e) { parsed = requestTask; }
        } else {
            parsed = requestTask;
        }

        // Helpers
        function isArrayOfTaskEntries(x) {
            return Array.isArray(x) && x.length > 0 && x.every(function(el) { return el && typeof el === 'object' && Object.keys(el).length === 1; });
        }

        function makePerTaskObj(taskId, innerObj) {
            var out = { tasks: taskId, request: [] };
            var wrapper = {};
            wrapper[taskId] = innerObj || {};
            out.request.push(wrapper);
            return out;
        }

        // Build list of per-task payloads to save (one element => one row)
        var tasksToSave = [];

        // Case 1: object with request array of single-task objects [{TaskId: {...}}...]
        if (parsed && typeof parsed === 'object' && Array.isArray(parsed.request) && isArrayOfTaskEntries(parsed.request)) {
            parsed.request.forEach(function(entry) {
                var tid = Object.keys(entry)[0];
                tasksToSave.push(makePerTaskObj(tid, entry[tid] || {}));
            });
        }
        // Case 2: an array directly like [{TaskId: {...}}, ...]
        else if (isArrayOfTaskEntries(parsed)) {
            parsed.forEach(function(entry) {
                var tid2 = Object.keys(entry)[0];
                tasksToSave.push(makePerTaskObj(tid2, entry[tid2] || {}));
            });
        }
        // Case 3: object mapping { "TaskId00005": {...}, "TaskId00008": {...} }
        else if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) {
            var keys = Object.keys(parsed);
            // if looks like multi-task mapping, split; otherwise keep as single row
            var looksLikeMulti = keys.length > 1 && keys.every(function(k) { return /^TaskId/i.test(k) || k.indexOf('Task') !== -1; });
            if (looksLikeMulti) {
                keys.forEach(function(k) {
                    tasksToSave.push(makePerTaskObj(k, parsed[k]));
                });
            } else {
                tasksToSave.push(parsed);
            }
        }
        // Case 4: fallback for string or other shapes => save single row preserving original value
        else {
            tasksToSave.push(parsed === null ? (requestTask || '') : parsed);
        }

        if (!tasksToSave || tasksToSave.length === 0) tasksToSave = [requestTask || ''];

        // Append one row per task object
        var createdIds = [];
        for (var idx = 0; idx < tasksToSave.length; idx++) {
            var single = tasksToSave[idx];

            // Prepare row with same number of columns as header
            var rowArr = new Array(finalHeaders.length).fill('');

            var requestId = 'REQ-' + (new Date()).getTime() + '-' + Math.floor(Math.random() * 9000 + 1000);
            var statusValue = 'Requested';
            var dataChangedValue = 'No';

            var requestTaskValue = single;
            if (requestTaskValue !== null && requestTaskValue !== undefined && typeof requestTaskValue !== 'string') {
                try { requestTaskValue = JSON.stringify(requestTaskValue); } catch (e) { requestTaskValue = String(requestTaskValue); }
            }

            if (headerMap['requestid'] !== undefined) rowArr[headerMap['requestid']] = requestId;
            else rowArr[0] = requestId;

            if (headerMap['requestfrom'] !== undefined) rowArr[headerMap['requestfrom']] = requestFromId;
            if (headerMap['requestto'] !== undefined) rowArr[headerMap['requestto']] = requestToId;
            if (headerMap['requesttask'] !== undefined) rowArr[headerMap['requesttask']] = requestTaskValue;
            if (headerMap['status'] !== undefined) rowArr[headerMap['status']] = statusValue;
            if (headerMap['datachanged'] !== undefined) rowArr[headerMap['datachanged']] = dataChangedValue;

            reqSheet.appendRow(rowArr);
            createdIds.push(requestId);
            // allow tiny gap so timestamp changes if fast loop (not essential but makes ids more unique)
            Utilities.sleep(6);
        }

        status.ok = true;
        status.msg = createdIds.join(',');
    } catch (err) {
        console.error('Failed to write Request sheet:', err);
        status.ok = false;
        status.msg = String(err);
    }
    return status;
}