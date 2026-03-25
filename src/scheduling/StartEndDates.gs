function findFirstAndLastAppearanceDates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formulaSheet = ss.getSheetByName('Formula Friendly Schedule');
  const jobTracking = ss.getSheetByName('Job Tracking');

  if (!formulaSheet || !jobTracking) {
    Logger.log('Missing sheet.');
    return;
  }

  const scheduleData = formulaSheet.getDataRange().getValues();
  const dateRow = scheduleData[2]; // row 3 (0-based index 2)
  const scheduleCols = scheduleData[0].length;

  // IDs from Job Tracking column B
  const ids = jobTracking.getRange('B2:B').getValues().map(r => r[0]);

  const output = []; // [firstDate, lastDate]

  ids.forEach(rawId => {
    const id = rawId ? rawId.toString().trim() : '';
    if (!id) {
      output.push(['', '']);
      return;
    }

    const idVariants = getIdVariants(id);

    let firstDate = null;
    let lastDate = null;

    for (let col = 0; col < scheduleCols; col++) {
      let foundInThisCol = false;

      for (let row = 1; row < Math.min(75, scheduleData.length); row++) {
        const cellVal = scheduleData[row][col];
        if (!cellVal) continue;

        const cellText = cellVal.toString().trim();
        if (idVariants.some(v => new RegExp(`\\b${escapeRegex_(v)}\\b`, 'i').test(cellText))) {
          foundInThisCol = true;
          break;
        }
      }

      if (foundInThisCol) {
        const date = dateRow[col];
        if (date instanceof Date) {
          if (!firstDate) firstDate = date;
          lastDate = date; // keep updating, so it ends as the last found
        }
      }
    }

    output.push([firstDate || '', lastDate || '']);
  });

  // Write to Job Tracking: C = first, D = last (starting row 2)
  jobTracking.getRange(2, 3, output.length, 2).setValues(output);
  Logger.log('First (C) and last (D) appearance dates written.');
}

/**
 * Escapes regex special chars for safe RegExp building.
 */
function escapeRegex_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
