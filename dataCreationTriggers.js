//------------------------------------- Data insertion-----------------------------------------------
function testDataCreationForAllFrequency() {
    weeklyTaskInsert();
    // yearlyTaskInsertTypeA();
    // monthlyTaskInsert();
    // weeklyTaskInsert();
    // halfYearlyTaskInsertTypeA();
    // quarterlyTaskInsertTypeA();
}

function yearlyTaskInsertTypeA() { taskInsert("Yearly", "A"); } // 3 month before // 1.5 months before session end
function yearlyTaskInsertTypeB() { taskInsert("Yearly", "B"); }

function quarterlyTaskInsertTypeA() { taskInsert("Quarterly", "A"); } // 10 days before end // 5 of next month
function quarterlyTaskInsertTypeB() { taskInsert("Quarterly", "B"); }

function monthlyTaskInsert() { taskInsert("Monthly", "none"); } // 25 rem // 5 next month

function fortnightlyTaskInsert() { taskInsert("Fortnightly", "none"); } // 12 remi // 18 esc

function weeklyTaskInsert() { taskInsert("Weekly", "none"); } // thursday  //Monday 12 am

function halfYearlyTaskInsertTypeA() { taskInsert("Biyearly", "A"); } // month starting // 5 of next month
function halfYearlyTaskInsertTypeB() { taskInsert("Biyearly", "B"); }

function taskInsert(frequency, shcoolType) {
    try {
        if (frequency == null || frequency == undefined) { throw new Error('Frequency is either null or undefined'); }
        let databaseTables = getDataToInfo();
        console.log('data==>' + JSON.stringify(databaseTables));

        createData(databaseTables, frequency, shcoolType);
    } catch (error) {
        // add the error issue in the sheet
        console.log('error occurred==>' + error);
    }
}

// freq and sequence of trigger needs to be changed tasks will be active for little long than task creation -> me
// figure out for different session timing -> me
// send emil with template -> me -Done
// clarity on the escalation 1 or 2 -> i/p

//**********************************************************************************************************/
function createData(databaseTables, frequency, shcoolType) {
    // fetch Task data
    let userData;
    const taskData = fetchTaskData(databaseTables, frequency);
    if (taskData == null || taskData == undefined) { throw new Error("Task data is returned null"); }
    console.log('task data==>' + JSON.stringify(taskData))

    // fetch all User data
    let filter = {
        "Active": "Y"
    }
    if (frequency == 'Yearly' || frequency == 'Biyearly' || frequency == 'Quarterly') {
        filter["Type"] = shcoolType;
        userData = fetchFilteredDataWithMap(databaseTables ? .userTable ? .name, filter);
    } else {
        userData = fetchFilteredDataWithMap(databaseTables ? .userTable ? .name, filter);
    }
    console.log('user data ==>' + JSON.stringify(userData));

    const createdData = createResponses(taskData, userData);
    console.log('created response data - ' + 'Record Cont: ' + createData ? .length + JSON.stringify(createdData));

    if (Array.isArray(createdData) && createdData.length > 0) {
        insertResponses(createdData);

    }

}

function fetchTaskData(databaseTables, frequency) {
    const taskFilter = {
        "Frequency": frequency
    }

    if (databaseTables ? .taskTable ? .name) {
        const filteredTaskData = fetchFilteredDataWithMap(databaseTables.taskTable.name, taskFilter);
        return filteredTaskData;
    }

    return null;
}