/**
 * Material List Link Finder
 *
 * Reads:
 * - Sheet: "Job Tracking"
 * - ID column: B
 * - Company column: D  (used to pick root folder)
 *
 * Writes:
 * - Link column: Y
 *
 * Behavior:
 * - Scans upward starting from the LAST non-empty ID in column B
 * - Only processes rows where Y is blank OR "No list found" OR "No List Created"
 * - If runtime is near limit, saves progress and exits
 * - Next run resumes from the last row processed (continues scanning upward)
 */

const ML_CONFIG = {
  SHEET: 'Job Tracking',
  COL_ID: 2,        // B
  COL_COMPANY: 4,   // D
  COL_LINK: 25,     // Y
  START_ROW: 2,

  // Progress persistence
  PROP_KEY_ROW: 'ML_LAST_ROW',

  // Time budget
  MAX_RUNTIME_MS: 5.5 * 60 * 1000, // stop before Apps Script hard limit
};

function populateMaterialListLinks_JobTracking() {
  const t0 = Date.now();
  Logger.log('populateMaterialListLinks_JobTracking() start');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ML_CONFIG.SHEET);
  if (!sheet) {
    Logger.log(`ERROR: Missing "${ML_CONFIG.SHEET}" sheet`);
    return;
  }

  const lastRowById = findLastNonEmptyRowInColumn_(sheet, ML_CONFIG.COL_ID, ML_CONFIG.START_ROW);
  if (lastRowById < ML_CONFIG.START_ROW) {
    Logger.log('No IDs found in column B.');
    return;
  }

  // Resume row: if stored row is invalid/outside current range, reset to lastRowById
  const props = PropertiesService.getScriptProperties();
  let resumeRow = Number(props.getProperty(ML_CONFIG.PROP_KEY_ROW) || '');
  if (!resumeRow || resumeRow > lastRowById || resumeRow < ML_CONFIG.START_ROW) {
    resumeRow = lastRowById;
  }

  Logger.log(`Last non-empty ID row in B: ${lastRowById}`);
  Logger.log(`Starting scan at row: ${resumeRow} (upward)`);

  let processed = 0;
  let updated = 0;

  for (let r = resumeRow; r >= ML_CONFIG.START_ROW; r--) {
    // Time guard
    if (Date.now() - t0 > ML_CONFIG.MAX_RUNTIME_MS) {
      props.setProperty(ML_CONFIG.PROP_KEY_ROW, String(r));
      Logger.log(`TIME_GUARD: saving resume row ${r} and exiting`);
      Logger.log(`Processed=${processed} Updated=${updated}`);
      return;
    }

    const id = sheet.getRange(r, ML_CONFIG.COL_ID).getDisplayValue().toString().trim();
    if (!id) continue; // skip blanks in the ID column

    const linkVal = sheet.getRange(r, ML_CONFIG.COL_LINK).getDisplayValue().toString().trim();

    // Only process blanks or no-list markers
    if (!(linkVal === '' || /^(no\s+list\s+found|no\s+list\s+created)$/i.test(linkVal))) {
      continue;
    }

    const companyRaw = sheet.getRange(r, ML_CONFIG.COL_COMPANY).getDisplayValue().toString().trim();
    const folderName = getFolderNameForCompany_(companyRaw);

    processed++;

    if (!folderName) {
      Logger.log(`Row ${r}: unknown company "${companyRaw}" (ID=${id}), skipping`);
      continue;
    }

    Logger.log(`Row ${r}: searching ID="${id}", folder="${folderName}"`);

    let foundUrl = '';
    const folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
      const root = folders.next();
      foundUrl = searchFolderRecursively_(root, id, []); // keep keywords empty unless you want prioritization
    } else {
      Logger.log(`Row ${r}: root folder not found: "${folderName}"`);
    }

    sheet.getRange(r, ML_CONFIG.COL_LINK).setValue(foundUrl || 'No list found');
    updated++;
  }

  // Completed full scan; clear resume pointer
  props.deleteProperty(ML_CONFIG.PROP_KEY_ROW);
  Logger.log(`populateMaterialListLinks_JobTracking() complete. Processed=${processed} Updated=${updated}`);
}

/* =========================
 * Helpers
 * ========================= */

function findLastNonEmptyRowInColumn_(sheet, col, startRow) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return 0;

  // Read the whole column slice once
  const vals = sheet.getRange(startRow, col, lastRow - startRow + 1, 1).getDisplayValues();

  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0] || '').trim() !== '') return startRow + i;
  }
  return 0;
}

// Map D value to the correct root folder
function getFolderNameForCompany_(companyRaw) {
  if (!companyRaw) return '';
  const v = companyRaw.toLowerCase();

  if (v.indexOf('bell mech') !== -1) return 'Material Lists - Bell Mechanical';
  if (v.indexOf('pack timco') !== -1) return 'Material Lists - Pack Timco';
  if (v.indexOf('metro') !== -1) return 'Material Lists - Metro Aire';
  if (v.indexOf('mck') !== -1) return 'Material Lists - MCK Plumbing';

  return '';
}

// ─── HELPER: check if filename contains any keyword ───
function filenameMatchesKeywords_(fileNameLower, keywords) {
  if (!keywords || keywords.length === 0) return false;
  for (var i = 0; i < keywords.length; i++) {
    if (fileNameLower.indexOf(keywords[i]) !== -1) return true;
  }
  return false;
}

// ─── HELPER: process a list of Drive Files; return URL if code matched in C5 ───
function checkFilesForCode_(files, code) {
  while (files.length) {
    var file = files.shift();
    try {
      var ssFile = SpreadsheetApp.openById(file.getId());
      var matIn = ssFile.getSheetByName('Material Input');
      if (!matIn) {
        Logger.log('      ' + file.getName() + ': no "Material Input" sheet');
        continue;
      }
      var c5 = matIn.getRange('C5').getDisplayValue().toString().trim();
      if (c5 === code) {
        Logger.log('      MATCH in file: ' + file.getName());
        return file.getUrl();
      }
    } catch (e) {
      Logger.log('      Error opening "' + file.getName() + '": ' + e);
    }
  }
  return '';
}

// ─── HELPER: recursively search folders ───
function searchFolderRecursively_(folder, code, keywords) {
  var prefer = [];
  var others = [];
  var it = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (it.hasNext()) {
    var f = it.next();
    var nameLower = f.getName().toLowerCase();
    if (filenameMatchesKeywords_(nameLower, keywords)) {
      prefer.push(f);
    } else {
      others.push(f);
    }
  }
  if (prefer.length || others.length) {
    Logger.log('    Folder "' + folder.getName() + '": ' + prefer.length + ' prioritized, ' + others.length + ' other files');
  }

  var url = checkFilesForCode_(prefer, code);
  if (url) return url;

  url = checkFilesForCode_(others, code);
  if (url) return url;

  var subs = folder.getFolders();
  while (subs.hasNext()) {
    var sub = subs.next();
    Logger.log('    Descending into subfolder: ' + sub.getName());
    var found = searchFolderRecursively_(sub, code, keywords);
    if (found) return found;
  }

  return '';
}
