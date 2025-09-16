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

        // Try parse if changesPayload is a JSON string
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

        // Resolve supervisor
        var supField = userData['SupervisorIds'] || '';
        if (!supField && typeof fetchFilteredDataWithMap === 'function') {
            var rec = (fetchFilteredDataWithMap('Users', { Email: userData['Email'] }) || [])[0];
            supField = rec ? (rec['SupervisorIds'] || '') : '';
        }
        if (!supField) {
            return { status: 'fail', message: 'Supervisor not set.' };
        }

        var first = supField.split(',').map(function(s) { return s.trim(); }).filter(Boolean)[0];
        var supRecord = null;
        if (first && first.indexOf('@') !== -1) {
            supRecord = (typeof fetchFilteredDataWithMap === 'function') ?
                (fetchFilteredDataWithMap('Users', { Email: first }) || [])[0] || { Email: first } :
                { Email: first };
        } else {
            supRecord = (typeof fetchFilteredDataWithMap === 'function') ?
                (fetchFilteredDataWithMap('Users', { UserId: first }) || [])[0] || null :
                null;
        }
        if (!supRecord || !supRecord.Email || !supRecord.UserId) {
            return { status: 'fail', message: 'Supervisor ID/email not resolved.' };
        }

        var supEmail = supRecord.Email;
        var supName = supRecord['User Name'] || supEmail.split('@')[0];
        var userName = userData['User Name'] || userData['Email'];
        var subject = 'Request to Grant Edit Access — ' + userName;

        // Helper: extract field-old/new pairs for a given taskId.
        // Returns array of { field, oldVal, newVal }.
        function extractChangesForTask(changes, taskId) {
            var out = [];
            if (!changes) return out;

            // If changes is an array: [{ "TaskId00005": { "frequency_old":"..", "frequency_new":".." } }, ...]
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
                                // if keys do not follow _old/_new, treat as new value only
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

            // If changes is an object mapping: { "TaskId00005": { ... } } or wrapper { request: [...] }
            if (typeof changes === 'object') {
                // direct mapping
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

                // wrapper shape: { request: [ { TaskId... } ] }
                if (Array.isArray(changes.request)) {
                    return extractChangesForTask(changes.request, taskId);
                }

                // fallback: scan request-like keys
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
            // include only if at least one old or new present
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

        // Send email with both plain text and html
        GmailApp.sendEmail(supEmail, subject, bodyText, { htmlBody: bodyHtml });

        // Save request JSON to sheet (stringify changesPayload if present)
        var changesJsonString = '';
        try {
            changesJsonString = (changesPayload && typeof changesPayload !== 'string') ? JSON.stringify(changesPayload) : String(changesPayload || '');
        } catch (e) {
            changesJsonString = String(changesPayload || '');
        }
        var saveStatus = saveAccessRequest(userData.UserId, supRecord.UserId, changesJsonString);

        if (saveStatus && saveStatus.ok) {
            return { status: 'success', message: 'Request sent to ' + supEmail + ' and saved (' + saveStatus.msg + ').', requestId: saveStatus.msg };
        } else if (saveStatus && !saveStatus.ok) {
            return { status: 'success', message: 'Email sent to ' + supEmail + ' but saving request failed: ' + saveStatus.msg };
        } else {
            return { status: 'success', message: 'Request sent to ' + supEmail + '.' };
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

        // Desired minimal headers (used only when creating sheet)
        var baseHeaders = ['RequestId', 'RequestFrom', 'RequestTo', 'RequestTask', 'Status', 'DataChanged'];

        if (!reqSheet) {
            // create sheet with minimal headers
            reqSheet = ss.insertSheet(sheetName);
            reqSheet.getRange(1, 1, 1, baseHeaders.length).setValues([baseHeaders]);
        } else {
            // Read current header row (read enough columns to include current last column)
            var lastCol = Math.max(1, reqSheet.getLastColumn());
            var headers = reqSheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];

            // Helper to find header index case-insensitive
            function findIndexCaseInsensitive(arr, name) {
                name = (name || '').toString().trim().toLowerCase();
                for (var i = 0; i < arr.length; i++) {
                    if ((arr[i] || '').toString().trim().toLowerCase() === name) return i;
                }
                return -1;
            }

            // 1) Remove Deadline column if present
            var deadlineIdx = findIndexCaseInsensitive(headers, 'deadline');
            if (deadlineIdx !== -1) {
                // deleteColumn uses 1-based indexes
                reqSheet.deleteColumn(deadlineIdx + 1);
                // Refresh header array after deletion
                lastCol = Math.max(1, reqSheet.getLastColumn());
                headers = reqSheet.getRange(1, 1, 1, lastCol).getValues()[0] || [];
            }

            // 2) Ensure DataChanged exists - if not, insert after Status if possible, otherwise append at end
            var dataChangedIdx = findIndexCaseInsensitive(headers, 'datachanged') !== -1 ? findIndexCaseInsensitive(headers, 'datachanged') : -1;
            if (dataChangedIdx === -1) {
                var statusIdx = findIndexCaseInsensitive(headers, 'status');
                if (statusIdx !== -1) {
                    // insert a column after Status
                    reqSheet.insertColumnAfter(statusIdx + 1); // using 1-based
                    reqSheet.getRange(1, statusIdx + 2).setValue('DataChanged');
                    // no need to shift other data manually (Sheets handles it)
                } else {
                    // append at the end
                    reqSheet.insertColumnAfter(reqSheet.getLastColumn());
                    reqSheet.getRange(1, reqSheet.getLastColumn()).setValue('DataChanged');
                }
            }
        }

        // Re-read headers and mapping before append
        var finalLastCol = Math.max(1, reqSheet.getLastColumn());
        var finalHeaders = reqSheet.getRange(1, 1, 1, finalLastCol).getValues()[0] || [];

        // build header -> index map
        var headerMap = {};
        for (var i = 0; i < finalHeaders.length; i++) {
            var key = (finalHeaders[i] || '').toString().trim();
            if (key) headerMap[key.toLowerCase()] = i;
        }

        // Prepare row array with the same length as header row
        var rowArr = new Array(finalHeaders.length).fill('');

        // Create request id and values
        var requestId = 'REQ-' + (new Date()).getTime() + '-' + Math.floor(Math.random() * 9000 + 1000);
        var statusValue = 'Requested';
        var dataChangedValue = 'No';

        // Ensure requestTask is string (stringify if object)
        var requestTaskValue = requestTask;
        if (requestTaskValue !== null && requestTaskValue !== undefined && typeof requestTaskValue !== 'string') {
            try { requestTaskValue = JSON.stringify(requestTaskValue); } catch (e) { requestTaskValue = String(requestTaskValue); }
        }

        // Put values into their header columns if present
        if (headerMap['requestid'] !== undefined) rowArr[headerMap['requestid']] = requestId;
        else rowArr[0] = requestId; // fallback to first column

        if (headerMap['requestfrom'] !== undefined) rowArr[headerMap['requestfrom']] = requestFromId;
        if (headerMap['requestto'] !== undefined) rowArr[headerMap['requestto']] = requestToId;
        if (headerMap['requesttask'] !== undefined) rowArr[headerMap['requesttask']] = requestTaskValue;
        if (headerMap['status'] !== undefined) rowArr[headerMap['status']] = statusValue;
        if (headerMap['datachanged'] !== undefined) rowArr[headerMap['datachanged']] = dataChangedValue;
        else {
            // if somehow DataChanged header is missing (shouldn't after ensure step), append at end
            rowArr.push(dataChangedValue);
        }

        // Append new row
        reqSheet.appendRow(rowArr);

        status.msg = requestId;
    } catch (err) {
        console.error('Failed to write Request sheet:', err);
        status.ok = false;
        status.msg = String(err);
    }
    return status;
}