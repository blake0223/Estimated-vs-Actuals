/**
 * Manday Breakdown Per Tech
 *
 * Runs in a SEPARATE spreadsheet from the master scheduling file.
 * Reads from the source spreadsheet:
 *   - "Formula Friendly Schedule" (schedule grid: row 1 = week labels, row 3 = dates, rows 4-75 = assignments)
 *   - "TechList"                  (tech manday values: A4:A = names, B2:P2 = week headers, B4:P = values)
 *   - "Job Tracking"             (B = Pipedrive ID, A = Job Name, F = Status, G = Start Date, H = Completed Date, P = Mandays)
 *
 * Writes to THIS spreadsheet:
 *   - "Mandays" tab: Pipedrive ID | Job Name | Tech | Mandays
 *
 * LOGIC:
 *   Same as MandayCalculatorPort, but instead of writing a single total per job,
 *   it attributes each matched cell's manday value to the tech whose name occupies
 *   that same [row, col] position in the grey-cell layer of the schedule.
 */

const MD_CONFIG = {
  SOURCE_SPREADSHEET_ID: '1MnkDZSNcNR4RFiPE3-QfSVtEvZ_M_qS_KdRHGiFyAvY',
  OUTPUT_SHEET_NAME: 'Mandays',
  MAX_SCHEDULE_ROW: 75,
};

/* ==============================================================
 * ENTRY POINT
 * ============================================================== */

function generateMandayBreakdownByTech() {
  const t0 = Date.now();
  console.log('generateMandayBreakdownByTech() start');

  const sourceSS = SpreadsheetApp.openById(MD_CONFIG.SOURCE_SPREADSHEET_ID);
  const formulaSheet = sourceSS.getSheetByName('Formula Friendly Schedule');
  const techListSheet = sourceSS.getSheetByName('TechList');
  const jobTracking = sourceSS.getSheetByName('Job Tracking');

  if (!formulaSheet || !techListSheet || !jobTracking) {
    throw new Error('Missing required sheet(s) in source spreadsheet.');
  }

  // ---- 1. Read the full schedule grid once ----
  const lastRow = formulaSheet.getLastRow();
  const lastCol = formulaSheet.getLastColumn();
  const scheduleData = formulaSheet.getRange(1, 1, lastRow, lastCol).getValues();
  const bgColors = formulaSheet.getRange(1, 1, lastRow, lastCol).getBackgrounds();

  const weekLabels = scheduleData[0].map(v => v ? String(v).trim() : '');
  const dateRow = scheduleData[2]; // row 3 (0-based index 2)

  // ---- 2. Determine grey color from A4 ----
  const greyColorCode = formulaSheet.getRange('A4').getBackground();

  // ---- 3. Build grey-cell tech-name grid (same dimensions as schedule) ----
  //   greyCellNames[row][col] = tech name string or '' if not grey
  const greyCellNames = [];
  for (let i = 0; i < scheduleData.length; i++) {
    greyCellNames[i] = [];
    for (let j = 0; j < scheduleData[0].length; j++) {
      if (i === 1 || i === 2) {
        // Skip rows 2 & 3 (structural rows)
        greyCellNames[i][j] = '';
      } else if (bgColors[i][j] === greyColorCode) {
        greyCellNames[i][j] = String(scheduleData[i][j] || '').trim();
      } else {
        greyCellNames[i][j] = '';
      }
    }
  }

  // ---- 4. Build TechList lookup (identical to MandayCalculatorPort) ----
  const techSeries = buildTechSeries_(techListSheet);
  const headerWeeksRaw = techListSheet.getRange('B2:P2').getValues()[0];
  const headerWeeks = headerWeeksRaw.map(v => v ? String(v).trim() : '');
  const headerKeys = headerWeeks.map(weekOrdinal_);

  // ---- 5. Determine which tech owns each schedule row ----
  //   Scan every grey cell in a row; the first one matching a known TechList name wins.
  const rowToTech = {};  // rowIndex -> tech name (original casing)
  for (let i = 0; i < greyCellNames.length; i++) {
    for (let j = 0; j < (greyCellNames[i] || []).length; j++) {
      const name = greyCellNames[i][j];
      if (!name) continue;
      if (techSeries[name.toLowerCase()]) {
        rowToTech[i] = name;
        break; // first known tech name in this row wins
      }
    }
  }

  console.log('Row-to-tech mapping: ' +
    Object.entries(rowToTech).map(([r, t]) => `r${Number(r)+1}=${t}`).join(', '));

  // ---- 6. Build in-memory manday map (identical to original MandayCalculatorPort) ----
  //   mandayMap[weekLabel][rowIndex] = numeric manday value
  //   Grey cells with known tech names get their TechList value; everything else = 0.5
  const mandayMap = {};
  for (let j = 0; j < weekLabels.length; j++) {
    const wl = weekLabels[j];
    if (!wl) continue;
    mandayMap[wl] = [];
    for (let i = 1; i < Math.min(MD_CONFIG.MAX_SCHEDULE_ROW, scheduleData.length); i++) {
      const techName = greyCellNames[i] && greyCellNames[i][j];
      if (techName && techSeries[techName.toLowerCase()]) {
        mandayMap[wl].push(getTechValueForWeek_(techName.toLowerCase(), wl, techSeries, headerKeys));
      } else {
        mandayMap[wl].push(0.5);
      }
    }
  }

  // ---- 7. Read Job Tracking ----
  const JT_COL_ID = 2;
  const JT_COL_STATUS = 6;
  const JT_COL_START = 7;
  const JT_COL_COMPLETED = 8;
  const JT_COL_MANDAYS = 16;

  const jtLastRow = jobTracking.getLastRow();
  if (jtLastRow < 2) {
    console.log('Job Tracking has no data rows.');
    return;
  }

  const jtData = jobTracking.getRange(2, 1, jtLastRow - 1, JT_COL_MANDAYS).getValues();

  const todayRaw = techListSheet.getRange('R1').getValue();
  const today = (todayRaw instanceof Date) ? todayRaw : new Date(todayRaw);
  if (isNaN(today.getTime())) {
    throw new Error('Invalid date in TechList!R1: ' + todayRaw);
  }

  const cutoffDate = new Date(today);
  cutoffDate.setMonth(cutoffDate.getMonth() + 2);

  const twoWeeksAgo = new Date(today);
  twoWeeksAgo.setDate(twoWeeksAgo.getDate() - 14);

  // ---- 8. Calculate mandays per job per tech ----
  //   results = [ { id, jobName, tech, mandays } ]
  const results = [];

  jtData.forEach((row, idx) => {
    const id = row[JT_COL_ID - 1];
    if (!id || String(id).trim() === '') return;

    const idStr = String(id).trim();
    const jobName = String(row[0] || '').trim();        // A
    const status = row[JT_COL_STATUS - 1];              // F
    const startDateRaw = row[JT_COL_START - 1];         // G
    const completedRaw = row[JT_COL_COMPLETED - 1];     // H
    const existingMandays = row[JT_COL_MANDAYS - 1];    // P
    const rowNumber = idx + 2;

    // ---- Skip rule (same as original) ----
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
      console.log(`Skipping ID ${idStr} at row ${rowNumber}: Completed >2wk ago, mandays=${existingMandays}`);
      return;
    }

    if (!startDateRaw) {
      console.log(`Start date missing for ID ${idStr} at row ${rowNumber}`);
      return;
    }

    const startDate = new Date(startDateRaw);
    if (isNaN(startDate.getTime())) {
      console.log(`Invalid start date for ID ${idStr} at row ${rowNumber}`);
      return;
    }

    // ---- Accumulate per-tech mandays for this job ----
    const techTotals = {};  // { techName: totalMandays }
    const idVariants = getIdVariants_(idStr);

    for (let col = 0; col < weekLabels.length; col++) {
      const weekLabel = weekLabels[col];
      const weekDate = dateRow[col];

      if (!weekLabel || !weekDate || isNaN(new Date(weekDate))) continue;

      const currentWeekDate = new Date(weekDate);
      if (currentWeekDate < startDate || currentWeekDate > cutoffDate) continue;

      for (let row = 1; row < Math.min(MD_CONFIG.MAX_SCHEDULE_ROW, scheduleData.length); row++) {
        const cellValue = scheduleData[row][col];
        if (!cellValue) continue;

        const cellText = String(cellValue).trim();
        const match = idVariants.some(variant =>
          new RegExp('\\b' + variant + '\\b', 'i').test(cellText)
        );

        if (!match) continue;

        // Manday value from the positional map (same as original)
        const mandayRowIndex = row - 1;
        const mandayValue = (mandayMap[weekLabel] && mandayMap[weekLabel][mandayRowIndex])
          ? mandayMap[weekLabel][mandayRowIndex]
          : 0.5;

        // Tech name from the row-level mapping
        const techName = rowToTech[row] || 'Unknown';

        if (!techTotals[techName]) techTotals[techName] = 0;
        techTotals[techName] += Number(mandayValue);
      }
    }

    // Push one result row per tech for this job
    const techs = Object.keys(techTotals);
    if (techs.length === 0) {
      results.push({ id: idStr, jobName, tech: '(none)', mandays: 0 });
    } else {
      for (const tech of techs) {
        results.push({ id: idStr, jobName, tech, mandays: techTotals[tech] });
      }
    }
  });

  // ---- 9. Pivot results: techs as rows, job names as columns ----
  //   Collect unique techs and job labels
  const techSet = new Set();
  const jobList = [];       // ordered list of { id, jobName, label }
  const jobLabelSet = new Set();

  for (const r of results) {
    techSet.add(r.tech);
    const label = r.jobName || r.id;
    if (!jobLabelSet.has(r.id)) {
      jobLabelSet.add(r.id);
      jobList.push({ id: r.id, jobName: r.jobName, label });
    }
  }

  const techs = Array.from(techSet).sort();
  // Build lookup: pivot[tech][id] = mandays
  const pivot = {};
  for (const r of results) {
    if (!pivot[r.tech]) pivot[r.tech] = {};
    pivot[r.tech][r.id] = (pivot[r.tech][r.id] || 0) + r.mandays;
  }

  // ---- 10. Write pivot table to "Mandays" tab ----
  const outputSS = SpreadsheetApp.getActiveSpreadsheet();
  let outputSheet = outputSS.getSheetByName(MD_CONFIG.OUTPUT_SHEET_NAME);
  if (!outputSheet) {
    outputSheet = outputSS.insertSheet(MD_CONFIG.OUTPUT_SHEET_NAME);
  } else {
    outputSheet.clearContents();
  }

  // Header row: "Tech" + each job name + "Total"
  const headerRow = ['Tech'];
  for (const job of jobList) headerRow.push(job.label);
  headerRow.push('Total');

  // Data rows: one per tech
  const dataRows = [];
  for (const tech of techs) {
    const row = [tech];
    let techTotal = 0;
    for (const job of jobList) {
      const val = (pivot[tech] && pivot[tech][job.id]) || 0;
      row.push(val);
      techTotal += val;
    }
    row.push(techTotal);
    dataRows.push(row);
  }

  // Totals row at the bottom
  const totalsRow = ['Total'];
  let grandTotal = 0;
  for (let c = 0; c < jobList.length; c++) {
    let colSum = 0;
    for (const dRow of dataRows) colSum += dRow[c + 1];
    totalsRow.push(colSum);
    grandTotal += colSum;
  }
  totalsRow.push(grandTotal);
  dataRows.push(totalsRow);

  const allRows = [headerRow, ...dataRows];
  const numCols = headerRow.length;

  outputSheet.getRange(2, 1, allRows.length, numCols).setValues(allRows);
  outputSheet.setFrozenRows(2);
  outputSheet.setFrozenColumns(1);

  const elapsed = Date.now() - t0;
  console.log(`generateMandayBreakdownByTech() complete. Techs=${techs.length} Jobs=${jobList.length} Elapsed=${elapsed}ms`);
  SpreadsheetApp.getActive().toast(
    `Done: ${techs.length} techs x ${jobList.length} jobs written to "${MD_CONFIG.OUTPUT_SHEET_NAME}".`,
    'Manday Breakdown',
    5
  );
}

/* ==============================================================
 * TECHLIST HELPERS (same logic as MandayCalculatorPort)
 * ============================================================== */

/**
 * "Week - 48 (2025)" -> 202548
 */
function weekOrdinal_(label) {
  if (!label) return NaN;
  const s = String(label).trim();
  const m = s.match(/^Week\s*-\s*(\d{1,2})\s*\((\d{4})\)$/i);
  if (!m) return NaN;
  const w = parseInt(m[1], 10);
  const y = parseInt(m[2], 10);
  return (isNaN(w) || isNaN(y)) ? NaN : y * 100 + w;
}

/**
 * Build { "tech name (lower)": [leftFilled values B..P] } from TechList.
 */
function buildTechSeries_(techListSheet) {
  const lastTechRow = techListSheet.getRange('A:A').getLastRow();
  const techRowCount = Math.max(0, lastTechRow - 3);

  const techNames = techRowCount
    ? techListSheet.getRange(4, 1, techRowCount, 1).getValues().flat()
    : [];
  const techValues = techRowCount
    ? techListSheet.getRange(4, 2, techRowCount, 16).getValues()
    : [];

  const techSeries = {};
  for (let r = 0; r < techRowCount; r++) {
    const nameNorm = String(techNames[r] || '').trim().toLowerCase();
    if (!nameNorm) continue;

    const rowVals = techValues[r].slice();
    // Left-fill: each new value applies at that week and all weeks after
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

  return techSeries;
}

/**
 * Look up a tech's manday value for a given week label.
 * Falls back to 0.5 if unknown.
 */
function getTechValueForWeek_(techNameNorm, weekLabel, techSeries, headerKeys) {
  const series = techSeries[techNameNorm];
  if (!series) return 0.5;

  const targetKey = weekOrdinal_(weekLabel);
  if (isNaN(targetKey)) return 0.5;

  // Last header <= target
  let idx = -1;
  for (let i = 0; i < headerKeys.length; i++) {
    const k = headerKeys[i];
    if (!isNaN(k) && k <= targetKey) idx = i;
  }
  if (idx === -1) idx = 0;

  const num = Number(series[idx]);
  return isNaN(num) ? 0.5 : num;
}

/**
 * Handles composite IDs like "12345/67890" -> ["12345", "67890"]
 */
function getIdVariants_(id) {
  const idStr = String(id).trim();
  return idStr.includes('/') ? idStr.split('/').map(s => s.trim()) : [idStr];
}
