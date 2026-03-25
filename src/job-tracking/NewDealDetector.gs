/**
 * DEPOSIT RECEIVED (Gmail) -> Job Tracking (Google Sheet)
 *
 * FIX: write to the FIRST SAFE BLANK ROW (not "lastRow+1")
 * Safe row rule:
 * - If column A is blank OR column B is blank, that row is considered writable.
 *
 * Writes:
 * - Col A = Job Name
 * - Col B = Pipedrive ID
 *
 * Skip:
 * - If ID already exists anywhere in column B, skip it.
 */

const TARGET_SPREADSHEET_ID = "1MnkDZSNcNR4RFiPE3-QfSVtEvZ_M_qS_KdRHGiFyAvY";
const TARGET_SHEET_NAME = "Job Tracking";

const DAYS_LOOKBACK = 180;
const GMAIL_QUERY_BASE = 'from:me subject:"Deposit Received"';

const THREAD_PAGE_SIZE = 100;
const MAX_PAGES = 25;

function logNewDealsFromEmail() {
  logExec_("RUN_START", "Deposit Received -> Job Tracking");

  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const sh = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!sh) throw new Error(`Sheet not found: ${TARGET_SHEET_NAME}`);

  const maxRows = sh.getMaxRows();
  const dataAB = sh.getRange(2, 1, maxRows - 1, 2).getValues(); // A2:B

  // Existing IDs set + first safe writable row
  const existingIds = new Set();
  let firstWritableIndex = -1; // index into dataAB (0-based), row = index+2

  for (let i = 0; i < dataAB.length; i++) {
    const aVal = String(dataAB[i][0] ?? "").trim();
    const bVal = String(dataAB[i][1] ?? "").trim();

    if (bVal) existingIds.add(bVal);

    // SAFE ROW: if A blank OR B blank, it is writable
    if (firstWritableIndex === -1 && (!aVal || !bVal)) {
      firstWritableIndex = i;
    }
  }

  logExec_("EXISTING_IDS", `Count=${existingIds.size}`);
  logExec_("FIRST_WRITABLE_ROW", firstWritableIndex === -1 ? "NONE" : String(firstWritableIndex + 2));

  if (firstWritableIndex === -1) {
    throw new Error("No writable rows found (A blank OR B blank). Add more rows or clear gaps.");
  }

  const gmailQuery = `${GMAIL_QUERY_BASE} newer_than:${DAYS_LOOKBACK}d`;
  logExec_("GMAIL_QUERY", gmailQuery);

  let totalThreads = 0;
  let inspected = 0;
  let skippedNoId = 0;
  let skippedNoName = 0;
  let skippedDup = 0;

  const newRows = []; // [jobName, id]

  for (let page = 0; page < MAX_PAGES; page++) {
    const start = page * THREAD_PAGE_SIZE;
    const threads = GmailApp.search(gmailQuery, start, THREAD_PAGE_SIZE);

    logExec_("THREAD_PAGE", `page=${page + 1} fetched=${threads.length} start=${start}`);
    if (!threads.length) break;

    totalThreads += threads.length;

    for (const thread of threads) {
      for (const msg of thread.getMessages()) {
        inspected++;

        const body = msg.getPlainBody() || "";
        const jobNameRaw = extractDepositName_(body);
        const idStr = extractDepositPipedriveId_(body);

        if (!idStr) {
          skippedNoId++;
          continue;
        }

        const jobName = toProperCase_(jobNameRaw);
        if (!jobName) {
          skippedNoName++;
          continue;
        }

        if (existingIds.has(idStr)) {
          skippedDup++;
          continue;
        }

        existingIds.add(idStr);
        newRows.push([jobName, idStr]);
        logExec_("QUEUED", `${idStr} | ${jobName}`);
      }
    }
  }

  logExec_("SUMMARY", `Threads=${totalThreads} Messages=${inspected} New=${newRows.length} NoID=${skippedNoId} NoName=${skippedNoName} Dup=${skippedDup}`);

  if (!newRows.length) {
    logExec_("RUN_END", "Nothing to write");
    return;
  }

  // Write sequentially into the first writable row, filling downward.
  // (Assumes you want to fill the table top-down without skipping gaps.)
  const startRow = firstWritableIndex + 2;
  const endRow = startRow + newRows.length - 1;

  if (endRow > sh.getMaxRows()) {
    throw new Error(`Not enough rows to write. Need rows ${startRow}..${endRow}, maxRows=${sh.getMaxRows()}`);
  }

  sh.getRange(startRow, 1, newRows.length, 2).setValues(newRows);

  logExec_("WROTE", `Rows=${newRows.length} startRow=${startRow}`);
  logExec_("RUN_END", "Done");
}

/* =========================
 * Parsing helpers
 * ========================= */

function extractDepositName_(body) {
  if (!body) return "";
  const lines = body.split(/\r?\n/).map(l => l.trim());

  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];

    // Same-line: "Deposit has been received for: JOB NAME"
    let m = l.match(/^Deposit has been received for\s*:\s*(.+)$/i);
    if (m && m[1]) return m[1].trim();

    // Next-line:
    if (/^Deposit has been received for\s*:\s*$/i.test(l)) {
      for (let j = i + 1; j < lines.length; j++) {
        if (lines[j]) return lines[j];
      }
      return "";
    }
  }
  return "";
}

function extractDepositPipedriveId_(body) {
  if (!body) return "";

  // Inline match
  let m = body.match(/Pipedrive ID\s*:\s*\(?\s*(\d+)\s*\)?/i);
  if (m) return String(m[1]).trim();

  // Next-line match
  const lines = body.split(/\r?\n/).map(l => l.trim());
  for (let i = 0; i < lines.length; i++) {
    if (/^Pipedrive ID\s*:\s*$/i.test(lines[i])) {
      for (let j = i + 1; j < Math.min(lines.length, i + 5); j++) {
        const mm = lines[j].match(/\(?\s*(\d+)\s*\)?/);
        if (mm) return String(mm[1]).trim();
      }
    }
  }

  return "";
}

/* =========================
 * Formatting + logging
 * ========================= */

function toProperCase_(str) {
  if (!str) return "";
  return String(str)
    .trim()
    .toLowerCase()
    .split(/\s+/)
    .map(w => (w ? w[0].toUpperCase() + w.slice(1) : ""))
    .join(" ");
}

function logExec_(tag, msg) {
  const line = `${new Date().toISOString()} | ${tag} | ${msg}`;
  Logger.log(line);
  console.log(line);
}
