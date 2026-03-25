function QuickMode() {
  // Step 4: Generate the manday map
  const mandayMap = generateMandayMap();
  SpreadsheetApp.getActive().toast('Step 4 Completed: Manday map generated.', 'Step 4 of 5', 3);

  // Step 5: Search for IDs and calculate mandays
  calculateMandaysForIDs(mandayMap);
  SpreadsheetApp.getActive().toast('Step 5 Completed: Mandays calculated for all Pipedrive IDs.', 'Step 5 of 5', 3);

  // Notify completion
  Logger.log('Manday calculation process completed.');
  SpreadsheetApp.getActive().toast('Manday calculation process completed.', 'Done', 5);
}

function MandayRecalculation() {
  // Step 1: Clear the 'Week Key' sheet
  clearWeekKeySheet();
  SpreadsheetApp.getActive().toast('Step 1 Completed: Week Key sheet cleared.', 'Step 1 of 5', 3);

  // Step 2: Copy only grey cells from 'Formula Friendly Schedule' to 'Week Key'
  copyGreyCellsToWeekKey();
  SpreadsheetApp.getActive().toast('Step 2 Completed: Grey cells copied to Week Key sheet.', 'Step 2 of 5', 3);

  // Step 3: Replace tech names with manday values in 'Week Key'
  replaceTechNamesWithMandayValues();
  SpreadsheetApp.getActive().toast('Step 3 Completed: Tech names replaced with manday values.', 'Step 3 of 5', 3);

  // Step 4: Generate the manday map
  const mandayMap = generateMandayMap();
  SpreadsheetApp.getActive().toast('Step 4 Completed: Manday map generated.', 'Step 4 of 5', 3);

  // Step 5: Search for IDs and calculate mandays
  calculateMandaysForIDs(mandayMap);
  SpreadsheetApp.getActive().toast('Step 5 Completed: Mandays calculated for all Pipedrive IDs.', 'Step 5 of 5', 3);

  // Notify completion
  Logger.log('Manday calculation process completed.');
  SpreadsheetApp.getActive().toast('Manday calculation process completed.', 'Done', 5);
}

function clearWeekKeySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const weekKeySheet = ss.getSheetByName('Week Key');
  if (!weekKeySheet) {
    Logger.log('Error: Week Key sheet is missing.');
    return;
  }
  weekKeySheet.clearContents();
  Logger.log('Week Key sheet has been cleared.');
}

function copyGreyCellsToWeekKey() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formulaSheet = ss.getSheetByName('Formula Friendly Schedule');
  const weekKeySheet = ss.getSheetByName('Week Key');
  if (!formulaSheet || !weekKeySheet) {
    Logger.log('Error: One or more required sheets are missing.');
    return;
  }

  const lastRow = formulaSheet.getLastRow();
  const lastCol = formulaSheet.getLastColumn();

  const dataRange = formulaSheet.getRange(1, 1, lastRow, lastCol);
  const dataValues = dataRange.getValues();
  const bgColors = dataRange.getBackgrounds();

  const greyCellValues = dataValues.map(row => row.map(() => ''));

  const knownGreyCell = formulaSheet.getRange('A4');
  const greyColorCode = knownGreyCell.getBackground();

  // Explicitly skip rows 2 and 3 (index 1 and 2)
  for (let i = 0; i < dataValues.length; i++) {
    if (i === 1 || i === 2) continue; // skip rows 2 & 3 outright
    for (let j = 0; j < dataValues[0].length; j++) {
      if (bgColors[i][j] === greyColorCode) {
        greyCellValues[i][j] = dataValues[i][j].toString();
      }
    }
  }

  weekKeySheet.getRange(1, 1, greyCellValues.length, lastCol).setValues(greyCellValues);
  Logger.log('Grey cells copied to Week Key sheet (rows 2 and 3 skipped).');
}

function replaceTechNamesWithMandayValues() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const weekKeySheet = ss.getSheetByName('Week Key');
  const techListSheet = ss.getSheetByName('TechList');
  const formulaSheet = ss.getSheetByName('Formula Friendly Schedule');

  if (!weekKeySheet || !techListSheet || !formulaSheet) {
    Logger.log('Error: One or more required sheets are missing.');
    return;
  }

  // ---- helper: "Week - 48 (2025)" -> 2025*100 + 48
  function weekOrdinal(label) {
    if (!label) return NaN;
    const s = String(label).trim();
    const m = s.match(/^Week\s*-\s*(\d{1,2})\s*\((\d{4})\)$/i);
    if (!m) return NaN;
    const w = parseInt(m[1], 10);
    const y = parseInt(m[2], 10);
    return (isNaN(w) || isNaN(y)) ? NaN : y * 100 + w;
  }

  // ======================
  // 1) TechList: build per-tech series from A4:A and B4:P
  // ======================

  // Use last row based on column A (names), not whole sheet.
  const lastTechRow = techListSheet.getRange('A:A').getLastRow();
  const techRowCount = Math.max(0, lastTechRow - 3); // rows 4+

  if (techRowCount === 0) {
    Logger.log('No tech rows found in TechList (A4:A is empty).');
  }

  const headerWeeksRaw = techListSheet.getRange('B2:P2').getValues()[0];
  const headerWeeks = headerWeeksRaw.map(v => v ? String(v).trim() : '');
  const headerKeys = headerWeeks.map(weekOrdinal);

  const techNames = techRowCount
    ? techListSheet.getRange(4, 1, techRowCount, 1).getValues().flat()
    : [];
  const techValues = techRowCount
    ? techListSheet.getRange(4, 2, techRowCount, 16).getValues() // B..P
    : [];

  const techSeries = {};
  for (let r = 0; r < techRowCount; r++) {
    const nameNorm = String(techNames[r] || '').trim().toLowerCase();
    if (!nameNorm) continue;

    const rowVals = techValues[r].slice();
    // left-fill: each new value applies at that week and all weeks after
    let lastSeen = null;
    for (let c = 0; c < rowVals.length; c++) {
      const v = rowVals[c];
      if (v === '' || v == null) {
        rowVals[c] = lastSeen;
      } else {
        lastSeen = v;
      }
    }
    techSeries[nameNorm] = rowVals;
  }

  function getTechValueForWeek(techNameNorm, weekLabel) {
    const series = techSeries[techNameNorm];
    if (!series) return 0.5;

    const targetKey = weekOrdinal(weekLabel);
    if (isNaN(targetKey)) return 0.5;

    // last header <= target
    let idx = -1;
    for (let i = 0; i < headerKeys.length; i++) {
      const k = headerKeys[i];
      if (!isNaN(k) && k <= targetKey) idx = i;
    }
    if (idx === -1) idx = 0;

    const num = Number(series[idx]);
    return isNaN(num) ? 0.5 : num;
  }

  // ======================
  // 2) Read Week Key + week labels with limited width
  // ======================
  const wkLastRow = weekKeySheet.getLastRow();
  if (wkLastRow < 2) {
    Logger.log('Week Key has no data to convert.');
    return;
  }

  const schedLastCol = formulaSheet.getLastColumn(); // authoritative width
  if (schedLastCol < 1) {
    Logger.log('Formula Friendly Schedule has no columns.');
    return;
  }

  // Week labels from Week Key row 1 (should be "Week - 48 (2025)" etc.)
  const headerRow = weekKeySheet.getRange(1, 1, 1, schedLastCol).getValues()[0];
  const weekLabels = headerRow.map(v => v ? String(v).trim() : '');

  // Tech names / values currently in Week Key (rows 2..end, cols 1..schedLastCol)
  const bodyValues = weekKeySheet.getRange(2, 1, wkLastRow - 1, schedLastCol).getValues();

  // Prepare output values: same size as existing Week Key (wkLastRow x schedLastCol)
  const outValues = new Array(wkLastRow);
  // row 1: keep week labels as-is
  outValues[0] = weekLabels.slice();

  // rows 2+: insert numeric values
  for (let r = 0; r < bodyValues.length; r++) {
    const rowOut = new Array(schedLastCol);
    const rowIn = bodyValues[r];
    for (let c = 0; c < schedLastCol; c++) {
      const raw = rowIn[c];

      // Empty stays empty
      if (raw === '' || raw == null) {
        rowOut[c] = '';
        continue;
      }

      // If already numeric (Week Key already converted), keep it. Don't re-map to 0.5.
      if (typeof raw === 'number') {
        rowOut[c] = raw;
        continue;
      }

      // Text: treat as tech name and map through TechList
      const techName = String(raw).trim();
      const norm = techName.toLowerCase();
      const weekLabel = weekLabels[c];
      const value = getTechValueForWeek(norm, weekLabel);

      rowOut[c] = value;

      if (value === 0.5 && !techSeries[norm]) {
        Logger.log(`Unrecognized tech "${techName}" at Week Key r${r + 2}, week "${weekLabel}". Assigned 0.5`);
      }
    }
    outValues[r + 1] = rowOut;
  }

  // ======================
  // 3) Write back once (no notes, no extra columns)
  // ======================
  const writeRange = weekKeySheet.getRange(1, 1, wkLastRow, schedLastCol);
  writeRange.clearContent(); // keep formats
  writeRange.setValues(outValues);

  Logger.log('Week Key updated: single column per week, numeric manday values only.');
}

function generateMandayMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const weekKeySheet = ss.getSheetByName('Week Key');
  if (!weekKeySheet) {
    Logger.log('Error: Week Key sheet is missing.');
    return {};
  }

  const lastRow = weekKeySheet.getLastRow();
  const lastCol = weekKeySheet.getLastColumn();

  const weekLabelsRaw = weekKeySheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const weekLabels = weekLabelsRaw.map(label => label ? label.toString().trim() : '');

  const mandayMap = {};
  for (let j = 0; j < weekLabels.length; j++) {
    const weekLabel = weekLabels[j];
    if (!weekLabel) continue;
    mandayMap[weekLabel] = [];

    for (let i = 1; i < Math.min(75, lastRow); i++) {
      let cellValue = weekKeySheet.getRange(i + 1, j + 1).getValue();
      if (cellValue === '' || cellValue == null) {
        cellValue = 0.5;
      } else {
        cellValue = Number(cellValue);
      }
      mandayMap[weekLabel].push(cellValue);
    }
    Logger.log(`Week Key "${weekLabel}": Values from rows 2-75: ${mandayMap[weekLabel].join(', ')}`);
  }

  Logger.log('Manday map generated successfully.');
  return mandayMap;
}

function calculateMandaysForIDs(mandayMap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formulaSheet = ss.getSheetByName('Formula Friendly Schedule');
  const jobTracking = ss.getSheetByName('Job Tracking');
  const techListSheet = ss.getSheetByName('TechList');

  if (!formulaSheet || !jobTracking || !techListSheet) {
    Logger.log('Error: Required sheet(s) missing.');
    return;
  }

  // Job Tracking columns:
  // A = Job Name
  // B = Pipedrive ID
  // F = Status
  // G = Start Date
  // H = Completed Date
  // P = Actual Mandays (output)
  const JT_COL_ID = 2;         // B
  const JT_COL_STATUS = 6;     // F
  const JT_COL_START = 7;      // G
  const JT_COL_COMPLETED = 8;  // H
  const JT_COL_MANDAYS = 16;   // O (output + existing mandays for skip)

  const jtLastRow = jobTracking.getLastRow();
  if (jtLastRow < 2) {
    Logger.log('Job Tracking has no data rows.');
    return;
  }

  // Read A:L so we have id/status/dates/mandays
  const jtData = jobTracking.getRange(2, 1, jtLastRow - 1, JT_COL_MANDAYS).getValues();
  const idInfo = [];

  jtData.forEach((row, idx) => {
    const id = row[JT_COL_ID - 1]; // B
    const status = row[JT_COL_STATUS - 1]; // F
    const startDateRaw = row[JT_COL_START - 1]; // G
    const completedRaw = row[JT_COL_COMPLETED - 1]; // H
    const existingMandays = row[JT_COL_MANDAYS - 1]; // L

    if (id && id.toString().trim() !== '') {
      idInfo.push({
        id: id.toString().trim(),
        startDateRaw,
        status,
        completedRaw,
        existingMandays,
        rowNumber: idx + 2, // because jtData starts at row 2
      });
    }
  });

  // "Today" from TechList!R1 (existing behavior)
  const today = techListSheet.getRange('R1').getValue();
  if (!(today instanceof Date)) {
    Logger.log('Error: Invalid date in TechList!R1.');
    return;
  }

  // Cutoff for future weeks (existing logic)
  const cutoffDate = new Date(today);
  cutoffDate.setMonth(cutoffDate.getMonth() + 2);

  // Two-week lookback for skipping completed jobs
  const twoWeeksAgo = new Date(today);
  twoWeeksAgo.setDate(twoWeeksAgo.getDate() - 14);

  const lastRow = formulaSheet.getLastRow();
  const lastCol = formulaSheet.getLastColumn();
  const scheduleData = formulaSheet.getRange(1, 1, lastRow, lastCol).getValues();
  const weekLabels = scheduleData[0].map(label => label ? label.toString().trim() : '');
  const dateRow = scheduleData[2];

  // Precompute cell A1 notations for logging
  const cellNotations = [];
  for (let row = 1; row < Math.min(75, scheduleData.length); row++) {
    cellNotations[row] = [];
    for (let col = 0; col < weekLabels.length; col++) {
      cellNotations[row][col] = formulaSheet.getRange(row + 1, col + 1).getA1Notation();
    }
  }

  const checkedCellsLog = [];

  idInfo.forEach(({ id, startDateRaw, status, completedRaw, existingMandays, rowNumber }) => {
    // Skip rule: Completed, completed > 2 weeks ago, and already has mandays > 0
    let skip = false;
    let completedDate = null;

    if (completedRaw instanceof Date && !isNaN(completedRaw)) {
      completedDate = completedRaw;
    } else if (completedRaw) {
      const tmp = new Date(completedRaw);
      if (!isNaN(tmp)) completedDate = tmp;
    }

    if (
      status === 'Completed' &&
      completedDate &&
      completedDate <= twoWeeksAgo &&
      Number(existingMandays) > 0
    ) {
      skip = true;
    }

    if (skip) {
      Logger.log(`Skipping ID ${id} at row ${rowNumber}: Completed, older than 2 weeks, mandays already set (${existingMandays}).`);
      return;
    }

    // Normal manday calculation for non-skipped rows
    if (!startDateRaw) {
      Logger.log(`Start date missing for ID ${id} at row ${rowNumber}`);
      return;
    }

    const startDate = new Date(startDateRaw);
    if (isNaN(startDate.getTime())) {
      Logger.log(`Invalid start date for ID ${id} at row ${rowNumber}`);
      return;
    }

    let totalMandays = 0;
    const idVariants = getIdVariants(id);

    for (let col = 0; col < weekLabels.length; col++) {
      const weekLabel = weekLabels[col];
      const weekDate = dateRow[col];

      if (!weekLabel || !weekDate || isNaN(new Date(weekDate))) continue;

      const currentWeekDate = new Date(weekDate);

      if (currentWeekDate >= startDate && currentWeekDate <= cutoffDate) {
        for (let row = 1; row < Math.min(75, scheduleData.length); row++) {
          const cellValue = scheduleData[row][col];
          if (cellValue) {
            const cellText = cellValue.toString().trim();
            const match = idVariants.some(variant =>
              new RegExp(`\\b${variant}\\b`, 'i').test(cellText)
            );

            if (match) {
              const mandayRowIndex = row - 1;
              const mandayValue = (mandayMap[weekLabel] && mandayMap[weekLabel][mandayRowIndex])
                ? mandayMap[weekLabel][mandayRowIndex]
                : 0.5;
              totalMandays += Number(mandayValue);
              checkedCellsLog.push([id, 'Formula Friendly Schedule', cellNotations[row][col], weekLabel, cellValue]);
            }
          }
        }
      }
    }

    // Write mandays to Job Tracking column L
    jobTracking.getRange(rowNumber, JT_COL_MANDAYS).setValue(totalMandays);
    Logger.log(`Mandays for ID ${id} (row ${rowNumber}): ${totalMandays}`);
    SpreadsheetApp.getActive().toast(`ID ${id} calculated: ${totalMandays} mandays`, 'Manday Calculation', 5);
  });

  logCheckedCells(checkedCellsLog);
}

function getIdVariants(id) {
  const idStr = id.toString().trim();
  return idStr.includes('/') ? idStr.split('/').map(s => s.trim()) : [idStr];
}

function logCheckedCells(logData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('CheckedLog');

  if (!logSheet) {
    logSheet = ss.insertSheet('CheckedLog');
  } else {
    logSheet.clearContents();
  }

  const headers = [['ID', 'Sheet', 'Cell', 'Week Label', 'Checked Value']];
  const outputData = headers.concat(logData);

  logSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);
  Logger.log('CheckedLog sheet updated.');
}
