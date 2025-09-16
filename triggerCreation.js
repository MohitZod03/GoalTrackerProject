/**
 * Calling triggers functions to create triggers
 */

function putCallingFunction() {
    //deleteAllTimeTriggers();
    //createWeeklyTriggers();
    //setupFortnightlyTriggers();
    //setupMonthlyTriggers();
    //setupQuarterlyTriggers();
    //testDateTriggers(4, 1); // Uncomment to test what would run on April 1st
}

/**
 * Deleting trigger
 */

function deleteAllTimeTriggers() {
    const allTriggers = ScriptApp.getProjectTriggers();
    allTriggers.forEach(trigger => {
        if (trigger.getTriggerSource() === ScriptApp.TriggerSource.CLOCK) {
            ScriptApp.deleteTrigger(trigger);
        }
    });
}

/**
 * Deleting trigger end
 */

/**
 * Weekly Triggers 
 */
function createWeeklyTriggers() {
    const triggerConfigs = [
        { functionName: "weeklyBufferResponses", day: ScriptApp.WeekDay.MONDAY, hour: 0 },
        { functionName: "weeklyEscalation1Trigger", day: ScriptApp.WeekDay.MONDAY, hour: 1 },
        { functionName: "weeklyTaskInsert", day: ScriptApp.WeekDay.MONDAY, hour: 2 },
        { functionName: "weeklyEscalation2Trigger", day: ScriptApp.WeekDay.TUESDAY, hour: 0 },
        { functionName: "weeklyDeactivateResponses", day: ScriptApp.WeekDay.TUESDAY, hour: 1 },
        { functionName: "weeklyEmailReminderTrigger", day: ScriptApp.WeekDay.THURSDAY, hour: 0 }
    ];

    triggerConfigs.forEach(config => {
        ScriptApp.newTrigger(config.functionName)
            .timeBased()
            .onWeekDay(config.day)
            .atHour(config.hour)
            .create();
    });
}

/**
 * Weekly Trigger End
 */


/**
 * Fortnightly Trigger
 */

function setupFortnightlyTriggers() {
    const fortnightlyTriggers = [
        { day: 1, hour: 0, func: 'fortnightlyBufferResponses' },
        { day: 16, hour: 0, func: 'fortnightlyBufferResponses' },

        { day: 1, hour: 1, func: 'fortnightlyEscalation1Trigger' },
        { day: 16, hour: 1, func: 'fortnightlyEscalation1Trigger' },

        { day: 1, hour: 2, func: 'fortnightlyTaskInsert' },
        { day: 16, hour: 2, func: 'fortnightlyTaskInsert' },

        { day: 3, hour: 0, func: 'fortnightlyEscalation2Trigger' },
        { day: 18, hour: 0, func: 'fortnightlyEscalation2Trigger' },

        { day: 3, hour: 1, func: 'fortnightlyDeactivateResponses' },
        { day: 18, hour: 1, func: 'fortnightlyDeactivateResponses' },

        { day: 10, hour: 0, func: 'fortnightlyEmailReminderTrigger' },
        { day: 25, hour: 0, func: 'fortnightlyEmailReminderTrigger' }
    ];

    // Create monthly time-based triggers
    fortnightlyTriggers.forEach(entry => {
        ScriptApp.newTrigger(entry.func)
            .timeBased()
            .onMonthDay(entry.day)
            .atHour(entry.hour)
            .everyMonths(1)
            .create();
    });
}

/**
 * Fortnightly Trigger End
 */

/**
 * Monthly Trigger
 */

const MONTHLY_TRIGGER_CONFIG = [
    { day: 1, hour: 0, func: 'monthlyBufferResponses' },
    { day: 1, hour: 1, func: 'monthlyEscalation1Trigger' },
    { day: 1, hour: 2, func: 'monthlyTaskInsert' },
    { day: 5, hour: 0, func: 'monthlyEscalation2Trigger' },
    { day: 5, hour: 1, func: 'monthlyDeactivateResponses' },
    { day: 20, hour: 0, func: 'monthlyEmailReminderTrigger' }
];

function setupMonthlyTriggers() {
    MONTHLY_TRIGGER_CONFIG.forEach(entry => {
        ScriptApp.newTrigger(entry.func)
            .timeBased()
            .onMonthDay(entry.day)
            .atHour(entry.hour)
            .everyMonths(1)
            .create();
    });
}

/**
 * Monthly Trigger End
 */


//   /**
//    * Quarterly Trigger 
//    */

// // CONFIGURATION
// const QUARTER_START_MONTH = 4; // Fiscal year starting month: April = 4, Jan = 1, etc.

// const QUARTERLY_TRIGGER_CONFIG = [
//   { offsetMonth: 0, day: 1, hour: 0, func: 'quarterlyBufferResponses' },
//   { offsetMonth: 0, day: 1, hour: 1, func: 'quarterlyEscalation1Trigger' },
//   { offsetMonth: 0, day: 1, hour: 2, func: 'quarterlyTaskInsert' },

//   { offsetMonth: 0, day: 5, hour: 0, func: 'quarterlyEscalation2Trigger' },
//   { offsetMonth: 0, day: 5, hour: 1, func: 'quarterlyDeactivateResponses' },

//   { offsetMonth: 2, day: 1, hour: 0, func: 'quarterlyEmailReminderTrigger' } // 3rd month of the quarter
// ];

// /**
//  * Sets up the quarterly trigger system.
//  * @param {number} quarterStartMonth - The first month of the fiscal year (1-12)
//  * @param {string|Date} [customStartDate=null] - Optional custom date to use as reference instead of current date
//  * @param {boolean} [forceInitialQuarter=false] - If true, forces creation of triggers starting from the specified quarterStartMonth
//  */
// function setupQuarterlyTriggers(quarterStartMonth = QUARTER_START_MONTH, customStartDate = null, forceInitialQuarter = true) {
//   // Delete any existing deactivated quarterly triggers
//   deleteDeactivatedQuarterlyTriggers();

//   // Get current date or use custom start date if provided
//   const now = customStartDate ? new Date(customStartDate) : new Date();
//   const currentYear = now.getFullYear();
//   const currentMonth = now.getMonth(); // 0-based month

//   // Calculate which quarter we're in based on the quarterStartMonth
//   const baseMonth = quarterStartMonth - 1; // Convert to 0-based month

//   let targetMonth, targetYear;

//   if (forceInitialQuarter) {
//     // For initial setup, use the exact specified quarter start month
//     targetMonth = baseMonth;
//     targetYear = currentYear;

//     // If the specified month is already in the past this year, use next year
//     if (baseMonth < currentMonth) {
//       targetYear = currentYear + 1;
//     }
//   } else {
//     // Calculate the current quarter and the next quarter (for regular operation)
//     const monthsSinceQuarterStart = (currentMonth - baseMonth + 12) % 12;
//     const currentQuarter = Math.floor(monthsSinceQuarterStart / 3);

//     // Calculate the next quarter
//     const nextQuarter = (currentQuarter + 1) % 4;
//     targetYear = currentYear + (nextQuarter === 0 && currentQuarter === 3 ? 1 : 0);

//     // Calculate the first month of the next quarter
//     targetMonth = (baseMonth + (nextQuarter * 3)) % 12;
//   }

//   console.log(`Setting up triggers starting from month: ${targetMonth + 1}/${targetYear}`);

//   // Create triggers for the target quarter
//   QUARTERLY_TRIGGER_CONFIG.forEach(entry => {
//     const totalMonths = targetMonth + entry.offsetMonth;
//     const month = totalMonths % 12;
//     const year = targetYear + Math.floor(totalMonths / 12);

//     const triggerDate = new Date(year, month, entry.day, entry.hour, 0, 0);

//     // Only create triggers if they're in the future
//     if (triggerDate > now) {
//       ScriptApp.newTrigger(entry.func)
//         .timeBased()
//         .at(triggerDate)
//         .create();

//       console.log(`Created trigger for ${entry.func} at ${triggerDate.toISOString()}`);
//     } else {
//       console.log(`Skipped creating trigger for ${entry.func} at ${triggerDate.toISOString()} (date is in the past)`);
//     }
//   });

//   // Calculate when the target quarter ends
//   const targetQuarterEndMonth = (targetMonth + 3 - 1) % 12; // Last month of the quarter
//   const targetQuarterEndYear = targetYear + (targetQuarterEndMonth < targetMonth ? 1 : 0);

//   // Calculate the date one day before the end of the target quarter
//   const lastDayOfQuarter = new Date(targetQuarterEndYear, targetQuarterEndMonth + 1, 0);
//   const oneDayBeforeQuarterEnd = new Date(lastDayOfQuarter);
//   oneDayBeforeQuarterEnd.setDate(lastDayOfQuarter.getDate() - 1);

//   // Set up a trigger to run setupNextQuarterlyTriggers one day before the quarter ends
//   // Only create it if it's in the future
//   if (oneDayBeforeQuarterEnd > now) {
//     ScriptApp.newTrigger('setupNextQuarterlyTriggers')
//       .timeBased()
//       .at(oneDayBeforeQuarterEnd)
//       .create();

//     console.log(`Scheduled to create next quarterly triggers on ${oneDayBeforeQuarterEnd.toISOString()}`);
//   } else {
//     // If one day before quarter end is already passed, schedule for the next quarter
//     const nextTargetMonth = (targetMonth + 3) % 12;
//     const nextTargetYear = targetYear + (nextTargetMonth < targetMonth ? 1 : 0);

//     const nextQuarterEndMonth = (nextTargetMonth + 3 - 1) % 12;
//     const nextQuarterEndYear = nextTargetYear + (nextQuarterEndMonth < nextTargetMonth ? 1 : 0);

//     const lastDayOfNextQuarter = new Date(nextQuarterEndYear, nextQuarterEndMonth + 1, 0);
//     const oneDayBeforeNextQuarterEnd = new Date(lastDayOfNextQuarter);
//     oneDayBeforeNextQuarterEnd.setDate(lastDayOfNextQuarter.getDate() - 1);

//     ScriptApp.newTrigger('setupNextQuarterlyTriggers')
//       .timeBased()
//       .at(oneDayBeforeNextQuarterEnd)
//       .create();

//     console.log(`Target quarter end date already passed. Scheduled for next quarter instead: ${oneDayBeforeNextQuarterEnd.toISOString()}`);
//   }
// }

// /**
//  * Sets up triggers for the next quarter. This function is called automatically
//  * one day before the end of each quarter.
//  */
// function setupNextQuarterlyTriggers() {
//   // For automatic scheduling, we don't want to force the initial quarter
//   setupQuarterlyTriggers(QUARTER_START_MONTH, null, false);
// }

// // Function to delete deactivated quarterly triggers
// function deleteDeactivatedQuarterlyTriggers() {
//   const allTriggers = ScriptApp.getProjectTriggers();
//   const quarterlyFunctions = QUARTERLY_TRIGGER_CONFIG.map(config => config.func);

//   allTriggers.forEach(trigger => {
//     const functionName = trigger.getHandlerFunction();

//     // Check if this is a quarterly function
//     if (quarterlyFunctions.includes(functionName)) {
//       const triggerDate = trigger.getTriggerSourceId() ? new Date(trigger.getTriggerSourceId()) : null;

//       // If the trigger date is in the past, delete it
//       if (triggerDate && triggerDate < new Date()) {
//         ScriptApp.deleteTrigger(trigger);
//         console.log(`Deleted past quarterly trigger: ${functionName} scheduled for ${triggerDate.toISOString()}`);
//       }
//     }
//   });
// }


// /**
//  * Quarterly Trigger End
//  */

/**
 * Yearly and Bi-Yearly Trigger Configurations
 */

const QUARTERLY_TRIGGER_CONFIG = {

    TYPE_A: {
        AM_12: [
            { month: 4, day: 1, func: 'quarterlyBufferResponsesTypeA' },
            { month: 7, day: 1, func: 'quarterlyBufferResponsesTypeA' },
            { month: 10, day: 1, func: 'quarterlyBufferResponsesTypeA' },
            { month: 1, day: 1, func: 'quarterlyBufferResponsesTypeA' },
            { month: 4, day: 5, func: 'quarterlyEscalation2TriggerTypeA' },
            { month: 7, day: 5, func: 'quarterlyEscalation2TriggerTypeA' },
            { month: 10, day: 5, func: 'quarterlyEscalation2TriggerTypeA' },
            { month: 1, day: 5, func: 'quarterlyEscalation2TriggerTypeA' },
            { month: 6, day: 1, func: 'quarterlyEmailReminderTriggerTypeA' },
            { month: 9, day: 1, func: 'quarterlyEmailReminderTriggerTypeA' },
            { month: 12, day: 1, func: 'quarterlyEmailReminderTriggerTypeA' },
            { month: 3, day: 1, func: 'quarterlyEmailReminderTriggerTypeA' },
        ],
        AM_1: [
            { month: 4, day: 1, func: 'quarterlyEscalation1TriggerTypeA' },
            { month: 7, day: 1, func: 'quarterlyEscalation1TriggerTypeA' },
            { month: 10, day: 1, func: 'quarterlyEscalation1TriggerTypeA' },
            { month: 1, day: 1, func: 'quarterlyEscalation1TriggerTypeA' },
            { month: 4, day: 5, func: 'quarterlyDeactivateResponsesTypeA' },
            { month: 7, day: 5, func: 'quarterlyDeactivateResponsesTypeA' },
            { month: 10, day: 5, func: 'quarterlyDeactivateResponsesTypeA' },
            { month: 1, day: 5, func: 'quarterlyDeactivateResponsesTypeA' },
        ],
        AM_2: [
            { month: 4, day: 1, func: 'quarterlyTaskInsertTypeA' },
            { month: 7, day: 1, func: 'quarterlyTaskInsertTypeA' },
            { month: 10, day: 1, func: 'quarterlyTaskInsertTypeA' },
            { month: 1, day: 1, func: 'quarterlyTaskInsertTypeA' },
        ]
    },
    TYPE_B: {
        AM_12: [
            { month: 6, day: 1, func: 'quarterlyBufferResponsesTypeB' },
            { month: 9, day: 1, func: 'quarterlyBufferResponsesTypeB' },
            { month: 12, day: 1, func: 'quarterlyBufferResponsesTypeB' },
            { month: 3, day: 1, func: 'quarterlyBufferResponsesTypeB' },
            { month: 6, day: 5, func: 'quarterlyEscalation2TriggerTypeB' },
            { month: 9, day: 5, func: 'quarterlyEscalation2TriggerTypeB' },
            { month: 12, day: 5, func: 'quarterlyEscalation2TriggerTypeB' },
            { month: 3, day: 5, func: 'quarterlyEscalation2TriggerTypeB' },
            { month: 8, day: 1, func: 'quarterlyEmailReminderTriggerTypeB' },
            { month: 11, day: 1, func: 'quarterlyEmailReminderTriggerTypeB' },
            { month: 2, day: 1, func: 'quarterlyEmailReminderTriggerTypeB' },
            { month: 5, day: 1, func: 'quarterlyEmailReminderTriggerTypeB' },
        ],
        AM_1: [
            { month: 6, day: 1, func: 'quarterlyEscalation1TriggerTypeB' },
            { month: 9, day: 1, func: 'quarterlyEscalation1TriggerTypeB' },
            { month: 12, day: 1, func: 'quarterlyEscalation1TriggerTypeB' },
            { month: 3, day: 1, func: 'quarterlyEscalation1TriggerTypeB' },
            { month: 6, day: 5, func: 'quarterlyDeactivateResponsesTypeB' },
            { month: 9, day: 5, func: 'quarterlyDeactivateResponsesTypeB' },
            { month: 12, day: 5, func: 'quarterlyDeactivateResponsesTypeB' },
            { month: 3, day: 5, func: 'quarterlyDeactivateResponsesTypeB' },
        ],
        AM_2: [
            { month: 6, day: 1, func: 'quarterlyTaskInsertTypeB' },
            { month: 9, day: 1, func: 'quarterlyTaskInsertTypeB' },
            { month: 12, day: 1, func: 'quarterlyTaskInsertTypeB' },
            { month: 3, day: 1, func: 'quarterlyTaskInsertTypeB' },
        ]
    }
};


// Bi-Yearly Trigger Configuration
const BIYEARLY_TRIGGER_CONFIG = {
    // Type A: April and October cycle
    TYPE_A: {
        // 12 AM Functions
        AM_12: [
            { month: 4, day: 1, func: 'halfYearlyBufferResponsesTypeA' }, // April 1
            { month: 10, day: 1, func: 'halfYearlyBufferResponsesTypeA' }, // October 1
            { month: 4, day: 5, func: 'halfYearlyEscalation1TriggerTypeA' }, // April 5
            { month: 10, day: 5, func: 'halfYearlyEscalation1TriggerTypeA' }, // October 5
            { month: 4, day: 15, func: 'halfYearlyEscalation2TriggerTypeA' }, // April 15
            { month: 10, day: 15, func: 'halfYearlyEscalation2TriggerTypeA' }, // October 15
            { month: 3, day: 20, func: 'halfYearlyReminderTriggerTypeA' }, // March 20
            { month: 9, day: 20, func: 'halfYearlyReminderTriggerTypeA' }, // September 20
        ],
        // 1 AM Functions
        AM_1: [
            { month: 4, day: 1, func: 'halfYearlyTaskInsertTypeA' }, // April 1
            { month: 10, day: 1, func: 'halfYearlyTaskInsertTypeA' }, // October 1
            { month: 4, day: 15, func: 'halfYearlyDeactivateResponsesTypeA' }, // April 15
            { month: 10, day: 15, func: 'halfYearlyDeactivateResponsesTypeA' }, // October 15
        ]
    },
    // Type B: May and November cycle
    TYPE_B: {
        // 12 AM Functions
        AM_12: [
            { month: 5, day: 1, func: 'halfYearlyBufferResponsesTypeB' }, // May 1
            { month: 11, day: 1, func: 'halfYearlyBufferResponsesTypeB' }, // November 1
            { month: 5, day: 5, func: 'halfYearlyEscalation1TriggerTypeB' }, // May 5
            { month: 11, day: 5, func: 'halfYearlyEscalation1TriggerTypeB' }, // November 5
            { month: 5, day: 15, func: 'halfYearlyEscalation2TriggerTypeB' }, // May 15
            { month: 11, day: 15, func: 'halfYearlyEscalation2TriggerTypeB' }, // November 15
            { month: 4, day: 1, func: 'halfYearlyReminderTriggerTypeB' }, // April 1
            { month: 10, day: 1, func: 'halfYearlyReminderTriggerTypeB' }, // October 1
        ],
        // 1 AM Functions
        AM_1: [
            { month: 5, day: 1, func: 'halfYearlyTaskInsertTypeB' }, // May 1
            { month: 11, day: 1, func: 'halfYearlyTaskInsertTypeB' }, // November 1
            { month: 6, day: 1, func: 'halfYearlyTaskInsertTypeB' }, // June 1 (additional date specified)
            { month: 5, day: 15, func: 'halfYearlyDeactivateResponsesTypeB' }, // May 15
            { month: 11, day: 15, func: 'halfYearlyDeactivateResponsesTypeB' }, // November 15
        ]
    }
};

// Yearly Trigger Configuration
const YEARLY_TRIGGER_CONFIG = {
    // Type A
    TYPE_A: {
        // 12 AM Functions
        AM_12: [
            { month: 4, day: 1, func: 'yearlyBufferResponsesTypeA' }, // April 1
            { month: 1, day: 5, func: 'yearlyEmailReminderTriggerTypeA' }, // January 5
            { month: 2, day: 5, func: 'yearlyEscalation1TriggerTypeA' }, // February 5
            { month: 3, day: 5, func: 'yearlyEscalation2TriggerTypeA' }, // March 5
        ],
        // 1 AM Functions
        AM_1: [
            { month: 4, day: 1, func: 'yearlyTaskInsertTypeA' }, // April 1
            { month: 4, day: 30, func: 'yearlyDeactivateResponsesTypeA' }, // April 30
        ]
    },
    // Type B
    TYPE_B: {
        // 12 AM Functions
        AM_12: [
            { month: 6, day: 1, func: 'yearlyBufferResponsesTypeB' }, // June 1
            { month: 1, day: 10, func: 'yearlyEmailReminderTriggerTypeB' }, // January 10
            { month: 2, day: 25, func: 'yearlyEscalation1TriggerTypeB' }, // February 25
            { month: 3, day: 25, func: 'yearlyEscalation2TriggerTypeB' }, // March 25
        ],
        // 1 AM Functions
        AM_1: [
            { month: 6, day: 1, func: 'yearlyTaskInsertTypeB' }, // June 1
            { month: 4, day: 25, func: 'yearlyDeactivateResponsesTypeB' }, // April 25
        ]
    }
};

/**
 * Function to set up in your 12 AM trigger
 * This function will check if any functions need to be executed for the current date at 12 AM
 */
function dailyCheck12AM() {
    const now = new Date();
    const currentMonth = now.getMonth() + 1; // getMonth() is 0-indexed, so add 1
    const currentDay = now.getDate();

    console.log(`Running daily trigger check for 12 AM on ${currentMonth}/${currentDay}`);

    // Quarterly Type A
    checkAndRunFunctions(QUARTERLY_TRIGGER_CONFIG.TYPE_A.AM_12, currentMonth, currentDay, "Quarterly Type A - 12 AM");
    // Quarterly Type B
    checkAndRunFunctions(QUARTERLY_TRIGGER_CONFIG.TYPE_B.AM_12, currentMonth, currentDay, "Quarterly Type B - 12 AM");

    // Check Bi-Yearly Type A functions - 12 AM
    checkAndRunFunctions(BIYEARLY_TRIGGER_CONFIG.TYPE_A.AM_12, currentMonth, currentDay, "Bi-Yearly Type A");

    // Check Bi-Yearly Type B functions - 12 AM
    checkAndRunFunctions(BIYEARLY_TRIGGER_CONFIG.TYPE_B.AM_12, currentMonth, currentDay, "Bi-Yearly Type B");

    // Check Yearly Type A functions - 12 AM
    checkAndRunFunctions(YEARLY_TRIGGER_CONFIG.TYPE_A.AM_12, currentMonth, currentDay, "Yearly Type A");

    // Check Yearly Type B functions - 12 AM
    checkAndRunFunctions(YEARLY_TRIGGER_CONFIG.TYPE_B.AM_12, currentMonth, currentDay, "Yearly Type B");
}

/**
 * Function to set up in your 1 AM trigger
 * This function will check if any functions need to be executed for the current date at 1 AM
 */
function dailyCheck1AM() {
    const now = new Date();
    const currentMonth = now.getMonth() + 1; // getMonth() is 0-indexed, so add 1
    const currentDay = now.getDate();

    console.log(`Running daily trigger check for 1 AM on ${currentMonth}/${currentDay}`);

    // Quarterly Type A
    checkAndRunFunctions(QUARTERLY_TRIGGER_CONFIG.TYPE_A.AM_1, currentMonth, currentDay, "Quarterly Type A - 1 AM");
    // Quarterly Type B
    checkAndRunFunctions(QUARTERLY_TRIGGER_CONFIG.TYPE_B.AM_1, currentMonth, currentDay, "Quarterly Type B - 1 AM");

    // Check Bi-Yearly Type A functions - 1 AM
    checkAndRunFunctions(BIYEARLY_TRIGGER_CONFIG.TYPE_A.AM_1, currentMonth, currentDay, "Bi-Yearly Type A");

    // Check Bi-Yearly Type B functions - 1 AM
    checkAndRunFunctions(BIYEARLY_TRIGGER_CONFIG.TYPE_B.AM_1, currentMonth, currentDay, "Bi-Yearly Type B");

    // Check Yearly Type A functions - 1 AM
    checkAndRunFunctions(YEARLY_TRIGGER_CONFIG.TYPE_A.AM_1, currentMonth, currentDay, "Yearly Type A");

    // Check Yearly Type B functions - 1 AM
    checkAndRunFunctions(YEARLY_TRIGGER_CONFIG.TYPE_B.AM_1, currentMonth, currentDay, "Yearly Type B");
}

/**
 * Function to set up in your 2 AM trigger
 * This function will check if any functions need to be executed for the current date at 2 AM
 */
function dailyCheck2AM() {
    const now = new Date();
    const currentMonth = now.getMonth() + 1; // getMonth() is 0-indexed, so add 1
    const currentDay = now.getDate();

    console.log(`Running daily trigger check for 2 AM on ${currentMonth}/${currentDay}`);

    // Quarterly Type A
    checkAndRunFunctions(QUARTERLY_TRIGGER_CONFIG.TYPE_A.AM_2, currentMonth, currentDay, "Quarterly Type A - 2 AM");
    // Quarterly Type B
    checkAndRunFunctions(QUARTERLY_TRIGGER_CONFIG.TYPE_B.AM_2, currentMonth, currentDay, "Quarterly Type B - 2 AM");
}

/**
 * Helper function to check and run functions based on the current date
 * @param {Array} functionList - List of functions with their month and day
 * @param {number} currentMonth - Current month (1-12)
 * @param {number} currentDay - Current day of the month
 * @param {string} triggerType - The type of trigger (for logging)
 */
function checkAndRunFunctions(functionList, currentMonth, currentDay, triggerType) {
    functionList.forEach(item => {
        if (item.month === currentMonth && item.day === currentDay) {
            try {
                console.log(`Executing ${triggerType} function: ${item.func}`);
                // Call the function dynamically
                this[item.func]();
            } catch (e) {
                console.error(`Error executing ${item.func}: ${e.message}`);
            }
        }
    });
}

/**
 * Test function to manually check what would run on a specific date
 * @param {number} month - Month (1-12)
 * @param {number} day - Day of the month
 */
function testDateTriggers(month, day) {
    console.log(`Testing triggers for date: ${month}/${day}`);

    console.log("--- Quarterly Type A - 12 AM Functions ---");
    checkAndRunFunctions(QUARTERLY_TRIGGER_CONFIG.TYPE_A.AM_12, month, day, "Quarterly Type A - 12 AM");
    console.log("--- Quarterly Type B - 12 AM Functions ---");
    checkAndRunFunctions(QUARTERLY_TRIGGER_CONFIG.TYPE_B.AM_12, month, day, "Quarterly Type B - 12 AM");

    console.log("--- Bi-Yearly Type A - 12 AM Functions ---");
    checkAndRunFunctions(BIYEARLY_TRIGGER_CONFIG.TYPE_A.AM_12, month, day, "Bi-Yearly Type A");

    console.log("--- Bi-Yearly Type B - 12 AM Functions ---");
    checkAndRunFunctions(BIYEARLY_TRIGGER_CONFIG.TYPE_B.AM_12, month, day, "Bi-Yearly Type B");

    console.log("--- Yearly Type A - 12 AM Functions ---");
    checkAndRunFunctions(YEARLY_TRIGGER_CONFIG.TYPE_A.AM_12, month, day, "Yearly Type A");

    console.log("--- Yearly Type B - 12 AM Functions ---");
    checkAndRunFunctions(YEARLY_TRIGGER_CONFIG.TYPE_B.AM_12, month, day, "Yearly Type B");

    console.log("--- Quarterly Type A - 1 AM Functions ---");
    checkAndRunFunctions(QUARTERLY_TRIGGER_CONFIG.TYPE_A.AM_1, month, day, "Quarterly Type A - 1 AM");
    console.log("--- Quarterly Type B - 1 AM Functions ---");
    checkAndRunFunctions(QUARTERLY_TRIGGER_CONFIG.TYPE_B.AM_1, month, day, "Quarterly Type B - 1 AM");

    console.log("--- Bi-Yearly Type A - 1 AM Functions ---");
    checkAndRunFunctions(BIYEARLY_TRIGGER_CONFIG.TYPE_A.AM_1, month, day, "Bi-Yearly Type A");

    console.log("--- Bi-Yearly Type B - 1 AM Functions ---");
    checkAndRunFunctions(BIYEARLY_TRIGGER_CONFIG.TYPE_B.AM_1, month, day, "Bi-Yearly Type B");

    console.log("--- Yearly Type A - 1 AM Functions ---");
    checkAndRunFunctions(YEARLY_TRIGGER_CONFIG.TYPE_A.AM_1, month, day, "Yearly Type A");

    console.log("--- Yearly Type B - 1 AM Functions ---");
    checkAndRunFunctions(YEARLY_TRIGGER_CONFIG.TYPE_B.AM_1, month, day, "Yearly Type B");
}

//this below trigger responsible for the hourly request data changed

function headerIndexMap(headers) {
    const map = {};
    headers.forEach((h, i) => {
        const key = (h || '').toString().trim().toLowerCase();
        if (key) map[key] = i;
    });
    return map;
}

function parseRequestTask(raw) {
    const out = [];
    if (raw === undefined || raw === null) return out;
    let obj = null;

    if (typeof raw === 'object') {
        obj = raw;
    } else {
        const s = ('' + raw).trim();
        if (s.startsWith('{') || s.startsWith('[')) {
            try {
                obj = JSON.parse(s);
            } catch (e) {
                obj = null;
            }
        }
    }

    // If parsed object
    if (obj && typeof obj === 'object') {
        // case: {"request":[{"TaskId00005":{...}}, ...], "tasks":"TaskId00005"}
        if (Array.isArray(obj.request) && obj.request.length) {
            obj.request.forEach(item => {
                if (!item || typeof item !== 'object') return;
                const keys = Object.keys(item);
                if (!keys.length) return;
                const tid = keys[0];
                out.push({ taskId: tid, changes: item[tid] || {} });
            });
            return out;
        }

        // case: obj.tasks present (csv or single) and optionally details in obj[tid]
        if (obj.tasks) {
            const tids = ('' + obj.tasks).split(/[,;]+/).map(s => s.trim()).filter(Boolean);
            tids.forEach(tid => {
                if (obj[tid] && typeof obj[tid] === 'object') out.push({ taskId: tid, changes: obj[tid] });
                else out.push({ taskId: tid, changes: {} });
            });
            return out;
        }

        // if object of { TaskId: changes, ... }
        Object.entries(obj).forEach(([k, v]) => {
            if (!k) return;
            if (typeof v === 'object') out.push({ taskId: k, changes: v });
        });
        return out;
    }

    // Fallback: raw CSV list of task ids
    const s = ('' + raw).trim();
    if (s) {
        const tids = s.split(/[,;]+/).map(x => x.trim()).filter(Boolean);
        tids.forEach(tid => out.push({ taskId: tid, changes: {} }));
    }
    return out;
}

/**
 * Fetch rows in Request sheet where Status = "Requested" && DataChanged = "No"
 * returns array: { rowIndex (1-based), RequestId, RequestTaskRaw, parsedChanges: [{taskId, changes}, ...], originalRow }
 */
function fetchRequestsToProcess() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // NOTE: using singular sheet name 'Request' per your setup
    const reqSheet = ss.getSheetByName('Request');
    if (!reqSheet) throw new Error('Request sheet not found (expected name "Request")');

    const values = reqSheet.getDataRange().getValues();
    if (values.length < 2) return [];

    const headers = values[0].map(h => (h || '').toString());
    const hmap = headerIndexMap(headers);

    // normalized header lookup handles slight name differences
    const colStatus = (hmap['status'] !== undefined) ? hmap['status'] : -1;
    const colDataChanged = (hmap['datachanged'] !== undefined) ? hmap['datachanged'] :
        (hmap['data changed'] !== undefined ? hmap['data changed'] : -1);
    const colRequestTask = (hmap['requesttask'] !== undefined) ? hmap['requesttask'] :
        (hmap['request task'] !== undefined ? hmap['request task'] : -1);
    const colRequestId = (hmap['requestid'] !== undefined) ? hmap['requestid'] :
        (hmap['request id'] !== undefined ? hmap['request id'] : -1);

    if (colStatus < 0 || colDataChanged < 0 || colRequestTask < 0) {
        throw new Error('Request sheet must contain headers: Status, DataChanged, RequestTask (check header row)');
    }

    const out = [];
    for (let r = 1; r < values.length; r++) {
        const row = values[r];
        const status = ('' + (row[colStatus] || '')).trim().toLowerCase();
        const dataChanged = ('' + (row[colDataChanged] || '')).trim().toLowerCase();
        if (status === 'approved' && dataChanged === 'no') {
            const rawReqTask = row[colRequestTask];
            const parsed = parseRequestTask(rawReqTask);
            out.push({
                rowIndex: r + 1,
                RequestId: colRequestId >= 0 ? row[colRequestId] : null,
                RequestTaskRaw: rawReqTask,
                parsedChanges: parsed,
                originalRow: row
            });
        }
    }
    return out;
}

/**
 * Normalize frequency to the exact allowed validation values.
 * Allowed target tokens (exact): "Weekly","Biyearly","Quarterly","Monthly","Yearly","Fortnightly"
 */
function normalizeFrequencyValue(raw) {
    if (raw === undefined || raw === null) return undefined;
    const s = ('' + raw).trim().toLowerCase();

    const map = {
        // weekly
        'weekly': 'Weekly',
        'week': 'Weekly',

        // biyearly (sheet uses spelled "Biyearly")
        'biyearly': 'Biyearly',
        'bi-yearly': 'Biyearly',
        'bi yearly': 'Biyearly',
        'biyearly': 'Biyearly',
        'biannual': 'Biyearly',
        'biannualy': 'Biyearly',
        'bi-year': 'Biyearly',
        'bi year': 'Biyearly',
        'bi-yearly': 'Biyearly',
        'bi-yearly': 'Biyearly',
        'biannually': 'Biyearly',
        'bi yearly': 'Biyearly',

        // quarterly / monthly / yearly
        'quarterly': 'Quarterly',
        'monthly': 'Monthly',
        'yearly': 'Yearly',

        // fortnightly
        'fortnightly': 'Fortnightly',
        'fortnight': 'Fortnightly'
    };

    if (map[s]) return map[s];

    // If exact match ignoring case exists in allowed set, return canonical
    const allowed = ['Weekly', 'Biyearly', 'Quarterly', 'Monthly', 'Yearly', 'Fortnightly'];
    for (const a of allowed) {
        if (a.toLowerCase() === s) return a;
    }

    // fallback: return trimmed original so caller can decide to write or skip
    return ('' + raw).trim();
}

/** (Replace your applyRequestChanges with this) */
function applyRequestChanges() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // tasks and request sheet names (use your singular 'Request' as you have)
    const tasksSheet = ss.getSheetByName('Tasks');
    const reqSheet = ss.getSheetByName('Request');
    if (!tasksSheet) throw new Error('Tasks sheet not found (expected name "Tasks")');
    if (!reqSheet) throw new Error('Request sheet not found (expected name "Request")');

    const requests = fetchRequestsToProcess();
    if (!requests || requests.length === 0) {
        Logger.log('No Requests to process.');
        return { requestsProcessed: 0, tasksUpdated: 0 };
    }

    // Load tasks sheet fully
    const tasksVals = tasksSheet.getDataRange().getValues();
    if (tasksVals.length < 1) throw new Error('Tasks sheet appears empty or missing header row.');
    const taskHeaders = tasksVals[0].map(h => (h || '').toString());
    const tmap = headerIndexMap(taskHeaders);

    const colTaskId = (tmap['taskid'] !== undefined) ? tmap['taskid'] : (tmap['task id'] !== undefined ? tmap['task id'] : -1);
    const colTaskName = (tmap['task name'] !== undefined) ? tmap['task name'] : (tmap['taskname'] !== undefined ? tmap['taskname'] : -1);
    const colKeyResults = (tmap['key results'] !== undefined) ? tmap['key results'] : (tmap['keyresults'] !== undefined ? tmap['keyresults'] : -1);
    const colFrequency = (tmap['frequency'] !== undefined) ? tmap['frequency'] : -1;
    const colEvidence = (tmap['task evidence'] !== undefined) ? tmap['task evidence'] : (tmap['taskevidence'] !== undefined ? tmap['taskevidence'] : -1);

    if (colTaskId < 0) throw new Error('Tasks sheet must contain a "TaskId" column header.');

    // build map TaskId -> row index in tasksVals (array index)
    const taskIdToIndex = {};
    for (let i = 1; i < tasksVals.length; i++) {
        const id = ('' + (tasksVals[i][colTaskId] || '')).trim();
        if (id) taskIdToIndex[id] = i;
    }

    let tasksUpdated = 0;
    const updatedTaskIds = new Set();

    function getNewValueFromChanges(changes, fieldNames) {
        if (!changes || typeof changes !== 'object') return undefined;
        const lower = {};
        Object.keys(changes).forEach(k => lower[k.toLowerCase()] = changes[k]);
        for (const base of fieldNames) {
            const key1 = (base + '_new').toLowerCase();
            if (lower.hasOwnProperty(key1)) return lower[key1];
            const key2 = (base + 'new').toLowerCase();
            if (lower.hasOwnProperty(key2)) return lower[key2];
        }
        for (const base of fieldNames) {
            if (lower.hasOwnProperty(base.toLowerCase())) return lower[base.toLowerCase()];
        }
        return undefined;
    }

    // iterate each request and its parsed changes
    for (const req of requests) {
        for (const entry of req.parsedChanges) {
            const tid = (entry.taskId || '').toString().trim();
            if (!tid) continue;
            const idx = taskIdToIndex[tid];
            if (idx === undefined) {
                Logger.log('TaskId not found in Tasks sheet: ' + tid);
                continue;
            }

            const changes = entry.changes || {};

            // pull new values if present
            const newFreqRaw = getNewValueFromChanges(changes, ['frequency', 'freq']);
            const newTaskName = getNewValueFromChanges(changes, ['taskname', 'task name', 'task']);
            const newKeyResult = getNewValueFromChanges(changes, ['keyresult', 'key results', 'key_results']);
            const newEvidence = getNewValueFromChanges(changes, ['evidence', 'task evidence', 'taskevidence']);

            let applied = false;

            if (newFreqRaw !== undefined && colFrequency >= 0) {
                // normalize to exact allowed label before writing (prevents data validation error)
                const normalizedFreq = normalizeFrequencyValue(newFreqRaw);
                // Only write if non-empty normalized value
                if (normalizedFreq !== undefined && String(normalizedFreq).trim() !== '') {
                    tasksVals[idx][colFrequency] = normalizedFreq;
                    applied = true;
                } else {
                    Logger.log('Skipping frequency write for ' + tid + ' because normalization returned empty for raw value: ' + newFreqRaw);
                }
            }

            if (newTaskName !== undefined && colTaskName >= 0) {
                tasksVals[idx][colTaskName] = newTaskName;
                applied = true;
            }
            if (newKeyResult !== undefined && colKeyResults >= 0) {
                tasksVals[idx][colKeyResults] = newKeyResult;
                applied = true;
            }
            if (newEvidence !== undefined && colEvidence >= 0) {
                tasksVals[idx][colEvidence] = newEvidence;
                applied = true;
            }

            if (applied) {
                tasksUpdated++;
                updatedTaskIds.add(tid);
            }
        }
    }

    // Write back tasks sheet only if any updates
    if (tasksUpdated > 0) {
        // IMPORTANT: Writing tasks first. If a data validation error still occurs it will be thrown here.
        const writeRange = tasksSheet.getRange(1, 1, tasksVals.length, tasksVals[0].length);
        writeRange.setValues(tasksVals);
        Logger.log('Tasks updated for ' + updatedTaskIds.size + ' distinct task(s), total field changes count approx: ' + tasksUpdated);
    } else {
        Logger.log('No task fields to update from parsed requests.');
    }

    // Now mark processed request rows DataChanged = "Yes"
    const reqVals = reqSheet.getDataRange().getValues();
    const reqHeaders = reqVals[0].map(h => (h || '').toString());
    const rmap = headerIndexMap(reqHeaders);
    const colDataChanged = (rmap['datachanged'] !== undefined) ? rmap['datachanged'] :
        (rmap['data changed'] !== undefined ? rmap['data changed'] : -1);
    if (colDataChanged < 0) throw new Error('Request sheet missing DataChanged column header.');

    for (const req of requests) {
        const row = req.rowIndex; // 1-based
        const arrIdx = row - 1;
        if (reqVals[arrIdx]) {
            reqVals[arrIdx][colDataChanged] = 'Yes';
        }
    }

    reqSheet.getRange(1, 1, reqVals.length, reqVals[0].length).setValues(reqVals);

    return {
        requestsProcessed: requests.length,
        tasksUpdatedCount: tasksUpdated,
        updatedTaskIds: Array.from(updatedTaskIds)
    };
}