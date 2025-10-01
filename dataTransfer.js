function transferThenRemoveDisabledResponses() {
    const t = transferDisabledResponses_CopyOnly(),
        r = removeResponses();
    Logger.log(`Transferred ${t.copiedCount} responses, removed ${r.removedCount}`);
    return { transferred: t, removed: r };
}

/* ---------------- core worker: copy only ---------------- */

function transferDisabledResponses_CopyOnly() {
    const taskIds = getDisabledTaskIds();
    if (!taskIds || taskIds.length === 0) {
        Logger.log('No disabled TaskIds found (both flags true). Nothing to copy.');
        return { copiedCount: 0, copiedIds: [] };
    }
    return copyResponsesForTaskIds(taskIds);
}

function copyResponsesForTaskIds(taskIds) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const responsesSheet = ss.getSheetByName('Responses');
    const backupSheet = getOrCreateSheet(ss, 'Backup');

    if (!responsesSheet) throw new Error('Responses sheet not found.');

    // Read responses
    const respData = responsesSheet.getDataRange().getValues();
    if (respData.length < 2) {
        Logger.log('No responses present.');
        return { copiedCount: 0, copiedIds: [] };
    }

    const respHeaders = respData[0].map(h => String(h || '').trim());
    const respRows = respData.slice(1);
    const rIdx = makeIndexFinder(respHeaders);
    const RESP_ID_IDX = rIdx(['responseid', 'response id', 'response_id', 'response-id']);
    const RESP_TASKID_IDX = rIdx(['taskid', 'task id', 'task_id']);

    if (RESP_ID_IDX === -1) throw new Error('ResponseId column not found in Responses sheet.');
    if (RESP_TASKID_IDX === -1) throw new Error('TaskId column not found in Responses sheet.');

    // Ensure Backup headers exist (exactly as Responses headers) and find where ResponseId is in backup
    ensureHeadersForBackup(backupSheet, respHeaders);
    const backupHeaders = backupSheet.getRange(1, 1, 1, backupSheet.getLastColumn()).getValues()[0].map(h => String(h || '').trim());
    const bIdx = makeIndexFinder(backupHeaders);
    const BACKUP_RESP_ID_IDX = bIdx(['responseid', 'response id', 'response_id', 'response-id']);
    if (BACKUP_RESP_ID_IDX === -1) throw new Error('ResponseId column not found in Backup after ensureHeaders.');

    // Build set of existing ResponseIds in Backup
    const existingBackupIds = new Set();
    const backupLastRow = backupSheet.getLastRow();
    if (backupLastRow > 1) {
        const idVals = backupSheet.getRange(2, BACKUP_RESP_ID_IDX + 1, backupLastRow - 1, 1).getValues();
        idVals.forEach(r => existingBackupIds.add(String(r[0] || '').trim()));
    }

    // Build target TaskId set
    const targetTaskSet = new Set(taskIds.map(id => String(id || '').trim()).filter(Boolean));

    // Collect rows to copy (for TaskIds in target set and not already in backup)
    const rowsToCopy = [];
    const copiedIds = [];
    for (let i = 0; i < respRows.length; i++) {
        const row = respRows[i];
        const taskId = String(row[RESP_TASKID_IDX] || '').trim();
        if (!targetTaskSet.has(taskId)) continue;

        const respId = String(row[RESP_ID_IDX] || '').trim();
        if (!respId) continue;

        if (!existingBackupIds.has(respId)) {
            rowsToCopy.push(row.slice()); // copy row
            copiedIds.push(respId);
            existingBackupIds.add(respId); // avoid duplicates within this run
        }
    }

    if (rowsToCopy.length === 0) {
        Logger.log('No matching responses to copy (either none for disabled TaskIds or already backed up).');
        return { copiedCount: 0, copiedIds: [] };
    }

    // Ensure backup header width is at least responses header width (if smaller, overwrite header to match)
    const respCols = respHeaders.length;
    if (backupSheet.getLastColumn() < respCols) {
        backupSheet.getRange(1, 1, 1, respCols).setValues([respHeaders]);
    }

    // Pad each row to exactly respCols (so columns align)
    const rowsAligned = rowsToCopy.map(row => {
        const copy = row.slice();
        while (copy.length < respCols) copy.push('');
        if (copy.length > respCols) copy.length = respCols; // trim extra if any
        return copy;
    });

    // Append in one operation
    const appendAt = backupSheet.getLastRow() + 1;
    backupSheet.getRange(appendAt, 1, rowsAligned.length, respCols).setValues(rowsAligned);

    Logger.log(`Copied ${rowsAligned.length} response(s) to Backup. IDs: ${copiedIds.join(', ')}`);
    return { copiedCount: rowsAligned.length, copiedIds };
}

/* ------------- helper: get TaskIds where both flags true ------------- */

function getDisabledTaskIds() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tasksSheet = ss.getSheetByName('Tasks');
    if (!tasksSheet) throw new Error('Tasks sheet not found.');

    const data = tasksSheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const headers = data[0].map(h => String(h || '').trim());
    const rows = data.slice(1);
    const idx = makeIndexFinder(headers);

    const TASKID_IDX = idx(['taskid', 'task id', 'task_id']);
    const DISABLE_CREATION_IDX = idx(['disable creation', 'disablecreation', 'disable_creation']);
    const DISABLE_TRIGGER_IDX = idx(['disable trigger', 'disabletrigger', 'disable_trigger']);

    if (TASKID_IDX === -1) throw new Error('TaskId column not found in Tasks sheet.');

    const result = [];
    for (let r = 0; r < rows.length; r++) {
        const row = rows[r];
        const taskId = String(row[TASKID_IDX] || '').trim();
        if (!taskId) continue;

        const rawCreate = safeCell(row, DISABLE_CREATION_IDX);
        const rawTrigger = safeCell(row, DISABLE_TRIGGER_IDX);
        if (isTruthy(rawCreate) && isTruthy(rawTrigger)) {
            result.push(taskId);
        }
    }
    return result;
}

/* ---------------- small utilities ---------------- */

function makeIndexFinder(headerRow) {
    const normalized = headerRow.map(h => normalizeHeader(String(h || '')));
    return function(variants) {
        if (!Array.isArray(variants)) variants = [variants];
        for (let v of variants) {
            const nv = normalizeHeader(String(v || ''));
            for (let i = 0; i < normalized.length; i++)
                if (normalized[i] === nv) return i;
        }
        // fallback: partial word match
        for (let v of variants) {
            const parts = String(v || '').toLowerCase().split(/\W+/).filter(Boolean);
            for (let i = 0; i < headerRow.length; i++) {
                const h = String(headerRow[i] || '').toLowerCase();
                if (parts.every(p => h.indexOf(p) !== -1)) return i;
            }
        }
        return -1;
    };
}

function normalizeHeader(h) {
    return String(h || '').trim().toLowerCase().replace(/[_\s\-]+/g, '');
}

function safeCell(row, idx) {
    if (typeof idx !== 'number' || idx < 0) return '';
    return row[idx];
}

function isTruthy(value) {
    const raw = String(value || '').trim().toLowerCase();
    return ['true', '1', 'yes', 'y', 'on'].includes(raw);
}

function getOrCreateSheet(ss, name) {
    let s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name);
    return s;
}

function ensureHeadersForBackup(sheet, responsesHeaders) {
    // If sheet empty -> create header row equal to responses headers
    if (sheet.getLastRow() === 0) {
        sheet.getRange(1, 1, 1, responsesHeaders.length).setValues([responsesHeaders.slice()]);
        return;
    }
    // If sheet has headers but different width, do nothing here —
    // copy routine will align rows to responsesHeaders length before appending.
}

/* ---------------- RemoveThe Responces  ---------------- */
function removeResponses() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const responsesSheet = ss.getSheetByName('Responses');
    if (!responsesSheet) throw new Error('Responses sheet not found.');

    const data = responsesSheet.getDataRange().getValues();
    if (data.length < 2) {
        Logger.log('No responses present.');
        return { removedCount: 0, removedIds: [] };
    }

    const headers = data[0].map(h => String(h || '').trim());
    const rows = data.slice(1);

    const rIdx = makeIndexFinder(headers);
    const RESP_ID_IDX = rIdx(['responseid', 'response id', 'response_id', 'response-id']);
    const RESP_TASKID_IDX = rIdx(['taskid', 'task id', 'task_id']);

    if (RESP_ID_IDX === -1) throw new Error('ResponseId column not found in Responses sheet.');
    if (RESP_TASKID_IDX === -1) throw new Error('TaskId column not found in Responses sheet.');

    // Get the TaskIds that have both disable flags true
    const disabledTaskIds = getDisabledTaskIds();
    const targetTaskSet = new Set(disabledTaskIds.map(id => String(id || '').trim()).filter(Boolean));
    if (targetTaskSet.size === 0) {
        Logger.log('No disabled TaskIds found (both flags true). Nothing to remove.');
        return { removedCount: 0, removedIds: [] };
    }

    // Partition rows into kept and removed
    const keptRows = [];
    const removedIds = [];
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const taskId = String(row[RESP_TASKID_IDX] || '').trim();
        const respId = String(row[RESP_ID_IDX] || '').trim();

        if (taskId && targetTaskSet.has(taskId)) {
            if (respId) removedIds.push(respId);
            // skip adding to keptRows (effectively removing it)
        } else {
            keptRows.push(row);
        }
    }

    if (removedIds.length === 0) {
        Logger.log('No responses matched disabled TaskIds (or they were already absent).');
        return { removedCount: 0, removedIds: [] };
    }

    // Prepare kept rows to exactly match header width
    const colCount = headers.length;
    const alignedKept = keptRows.map(r => {
        const copy = r.slice();
        while (copy.length < colCount) copy.push('');
        if (copy.length > colCount) copy.length = colCount;
        return copy;
    });

    // Overwrite sheet: clear contents then write header + kept rows (no gaps)
    responsesSheet.clearContents();
    responsesSheet.getRange(1, 1, 1, colCount).setValues([headers]);
    if (alignedKept.length > 0) {
        responsesSheet.getRange(2, 1, alignedKept.length, colCount).setValues(alignedKept);
    }

    Logger.log(`Removed ${removedIds.length} response(s) for disabled TaskIds. IDs: ${removedIds.join(', ')}`);
    return { removedCount: removedIds.length, removedIds };
}