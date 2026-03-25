/**
 * HVACBizPro IMPORT + PARSE (Gmail -> Raw -> Jobs)
 *
 * RULES:
 * 1) NEVER overwrite existing rows in "HVACBizPro Jobs"
 *    - Existing parsed rows are identified by Message ID in column A
 *    - If Message ID already exists in Jobs column A, skip it
 *    - This preserves any manual edits in the Jobs tab
 *
 * 2) Raw import still appends only new emails into "HVACBizPro Raw Emails"
 *    - Message ID key in Raw column A
 *
 * 3) NO duplicates
 *    - Raw dedupes by Message ID in column A
 *    - Jobs dedupes by Message ID in column A
 *
 * SHEETS:
 * - Raw:  "HVACBizPro Raw Emails"
 *   A=Message ID, B=Gmail Date, C=From, D=Subject, E=Plain Body
 *
 * - Jobs: "HVACBizPro Jobs"
 *   A=Message ID (parse key)
 */

const HVACBIZPRO = {
  SPREADSHEET_ID: "",

  SHEET_RAW: "HVACBizPro Raw Emails",
  SHEET_JOBS: "HVACBizPro Jobs",

  GMAIL_QUERY: 'from:communications@hvacbizpro.com subject:"Financial Summary"',
  THREAD_PAGE_SIZE: 200,
  MAX_PAGES: 50,
  WRITE_BATCH_SIZE: 200
};

function hvacBizPro_ImportAndParse() {
  logExec_("RUN_START", "hvacBizPro_ImportAndParse");

  const ss = getSpreadsheet_();
  const shRaw = getOrCreateSheet_(ss, HVACBIZPRO.SHEET_RAW, rawHeaders_());
  const shJobs = getOrCreateSheet_(ss, HVACBIZPRO.SHEET_JOBS, jobHeaders_());

  // Ensure headers exist
  shRaw.getRange(1, 1, 1, rawHeaders_()[0].length).setValues(rawHeaders_());
  shRaw.setFrozenRows(1);

  shJobs.getRange(1, 1, 1, jobHeaders_()[0].length).setValues(jobHeaders_());
  shJobs.setFrozenRows(1);

  const imported = hvacBizPro_ImportRaw_(shRaw);
  logExec_("IMPORT_DONE", `newRaw=${imported}`);

  const appended = hvacBizPro_ParseAndAppendJobs_(shRaw, shJobs);
  logExec_("PARSE_DONE", `newJobs=${appended}`);

  logExec_("RUN_END", "hvacBizPro_ImportAndParse");
}

function hvacBizPro_ImportRaw_(shRaw) {
  const existingMsgIds = loadExistingIdsFromColumn_(shRaw, 1); // Raw col A
  logExec_("RAW_EXISTING_IDS", `count=${existingMsgIds.size}`);

  const buffer = [];
  let totalThreads = 0;
  let totalMessages = 0;
  let newRows = 0;

  for (let page = 0; page < HVACBIZPRO.MAX_PAGES; page++) {
    const start = page * HVACBIZPRO.THREAD_PAGE_SIZE;
    const threads = GmailApp.search(HVACBIZPRO.GMAIL_QUERY, start, HVACBIZPRO.THREAD_PAGE_SIZE);

    logExec_("THREAD_PAGE", `page=${page + 1} fetched=${threads.length} start=${start}`);
    if (!threads.length) break;

    totalThreads += threads.length;

    for (const thread of threads) {
      for (const msg of thread.getMessages()) {
        totalMessages++;

        const msgId = msg.getId();
        if (existingMsgIds.has(msgId)) continue;

        const date = msg.getDate();
        const from = msg.getFrom() || "";
        const subject = msg.getSubject() || "";
        const plainBody = normalizeBody_(msg.getPlainBody() || "");

        existingMsgIds.add(msgId);
        buffer.push([msgId, date, from, subject, plainBody]);
        newRows++;

        if (buffer.length >= HVACBIZPRO.WRITE_BATCH_SIZE) {
          appendRows_(shRaw, buffer);
          buffer.length = 0;
          logExec_("RAW_PROGRESS", `newRaw=${newRows} scannedMsgs=${totalMessages}`);
        }
      }
    }
  }

  if (buffer.length) appendRows_(shRaw, buffer);

  logExec_("RAW_SUMMARY", `threads=${totalThreads} scannedMsgs=${totalMessages} newRaw=${newRows}`);
  return newRows;
}

function hvacBizPro_ParseAndAppendJobs_(shRaw, shJobs) {
  const rawLastRow = shRaw.getLastRow();
  if (rawLastRow < 2) return 0;

  const rawValues = shRaw.getRange(2, 1, rawLastRow - 1, 5).getValues();
  const COL_MSG_ID = 0;
  const COL_DATE = 1;
  const COL_SUBJECT = 3;
  const COL_BODY = 4;

  // Existing parsed keys from Jobs column A
  const parsedMsgIds = loadExistingIdsFromColumn_(shJobs, 1);
  logExec_("JOBS_EXISTING_KEYS", `count=${parsedMsgIds.size}`);

  const rowsToAppend = [];
  let appended = 0;
  let skippedAlreadyParsed = 0;
  let skippedNotMatch = 0;

  for (const row of rawValues) {
    const msgId = String(row[COL_MSG_ID] ?? "").trim();
    if (!msgId) continue;

    // Never overwrite existing Jobs rows
    if (parsedMsgIds.has(msgId)) {
      skippedAlreadyParsed++;
      continue;
    }

    const gmailDate = row[COL_DATE];
    const subject = row[COL_SUBJECT] || "";
    const body = normalizeBody_(String(row[COL_BODY] || ""));

    if (!body || !body.toUpperCase().includes("PROPOSAL") || !body.toUpperCase().includes("CUSTOMER")) {
      skippedNotMatch++;
      continue;
    }

    const p = parseHvacBizProBody_(body);

    const out = fixRowToLength_([
      msgId,                    // A Message ID = parse key
      gmailDate,                // B Gmail Date
      subject,                  // C Subject

      p.proposal || "",         // D
      p.sales_rep || "",        // E
      p.opened || "",           // F
      p.closed || "",           // G
      p.install || "",          // H

      p.customer_name || "",    // I
      p.address1 || "",         // J
      p.city || "",             // K
      p.state || "",            // L
      p.zip || "",              // M
      p.tel || "",              // N
      p.email || "",            // O

      p.selling_price || "",    // P
      p.equipment || "",        // Q
      p.parts || "",            // R
      p.services || "",         // S
      p.labor || "",            // T
      p.tax || "",              // U
      p.labor_hours || "",      // V
      p.cogs || "",             // W
      p.gross_margin || "",     // X
      p.gross_profit || "",     // Y
      p.cost_of_financing || "",// Z
      p.payment_method || "",   // AA
      p.commissions || "",      // AB
      p.proposal_link || ""     // AC
    ], jobHeaders_()[0].length);

    rowsToAppend.push(out);
    parsedMsgIds.add(msgId);
    appended++;
  }

  if (rowsToAppend.length) {
    appendRows_(shJobs, rowsToAppend);
    logExec_("JOBS_APPENDED", `count=${rowsToAppend.length}`);
  }

  logExec_("JOBS_PARSE_SUMMARY", `appended=${appended} skippedParsed=${skippedAlreadyParsed} skippedNotMatch=${skippedNotMatch}`);
  return appended;
}

/* =========================
 * PARSER
 * ========================= */

function parseHvacBizProBody_(body) {
  const cleaned = body.replace(/\*/g, "");
  const lines = cleaned.split("\n").map(l => l.trim()).filter(Boolean);

  const r = {
    proposal: findValueAfterLabel_(lines, "PROPOSAL"),
    sales_rep: findValueAfterLabel_(lines, "SALES REP"),
    opened: findValueAfterLabel_(lines, "OPENED"),
    closed: findValueAfterLabel_(lines, "CLOSED"),
    install: findValueAfterLabel_(lines, "INSTALL"),

    customer_name: null,
    address1: null,
    city: null,
    state: null,
    zip: null,
    tel: null,
    email: null,

    selling_price: pickValueAfterPercentFromLines_(lines, "Selling Price"),
    equipment: pickValueAfterPercentFromLines_(lines, "Equipment"),
    parts: pickValueAfterPercentFromLines_(lines, "Parts"),
    services: pickValueAfterPercentFromLines_(lines, "Services"),
    labor: pickValueAfterPercentFromLines_(lines, "Labor"),
    tax: pickValueAfterPercentFromLines_(lines, "Tax"),
    labor_hours: pickLastNumberFromLines_(lines, "Labor Hours"),
    cogs: pickValueAfterPercentFromLines_(lines, "COGS"),
    gross_margin: pickValueAfterPercentFromLines_(lines, "Gross Margin"),
    gross_profit: pickLastNumberFromLines_(lines, "Gross Profit $"),
    cost_of_financing: pickLastNumberFromLines_(lines, "Cost of financing"),
    payment_method: null,
    commissions: pickLastNumberFromLines_(lines, "Commissions"),

    proposal_link: extractProposalLink_(cleaned)
  };

  const custIdx = findIndexStartsWith_(lines, "CUSTOMER");
  if (custIdx >= 0) {
    const rawName = lines[custIdx].replace(/^CUSTOMER\s*/i, "").trim();
    r.customer_name = toProperCase_(rawName);

    r.address1 = lines[custIdx + 1] || null;

    const cs = parseCityStateZipLoose_(lines[custIdx + 2] || "");
    r.city = cs.city;
    r.state = cs.state;
    r.zip = cs.zip;

    for (let i = custIdx + 1; i < Math.min(lines.length, custIdx + 15); i++) {
      const l = lines[i];
      if (/^Tel\s*:/i.test(l)) r.tel = l.replace(/^Tel\s*:\s*/i, "").trim();
      if (/\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/i.test(l)) r.email = l.trim();
    }
  }

  const payLine = findLineStartsWith_(lines, "Payment Method");
  if (payLine) {
    const m = payLine.match(/Pay by:\s*(.+)$/i);
    r.payment_method = m ? m[1].trim() : payLine.replace(/^Payment Method/i, "").trim();
  }

  return r;
}

/* =========================
 * HEADERS
 * ========================= */

function rawHeaders_() {
  return [[
    "Message ID",
    "Gmail Date",
    "From",
    "Subject",
    "Plain Body"
  ]];
}

function jobHeaders_() {
  return [[
    "Message ID","Gmail Date","Subject",
    "Proposal","Sales Rep","Opened","Closed","Install",
    "Customer Name","Address Line 1","City","State","ZIP","Telephone","Email",
    "Selling Price","Equipment","Parts","Services","Labor","Tax","Labor Hours",
    "COGS","Gross Margin","Gross Profit $","Cost of financing","Payment Method",
    "Commissions","Proposal Link"
  ]];
}

/* =========================
 * HELPERS
 * ========================= */

function toProperCase_(str) {
  if (!str) return str;
  return String(str)
    .toLowerCase()
    .split(/\s+/)
    .map(w => w ? w[0].toUpperCase() + w.slice(1) : "")
    .join(" ");
}

function fixRowToLength_(row, len) {
  while (row.length < len) row.push("");
  return row.slice(0, len);
}

function findValueAfterLabel_(l, k) {
  return (findLineStartsWith_(l, k) || "").replace(new RegExp("^" + k + "\\s*", "i"), "").trim() || null;
}

function pickValueAfterPercentFromLines_(l, k) {
  const x = findLineStartsWith_(l, k);
  if (!x) return null;
  return (x.match(/%\s*([\d,]+)/) || [])[1] || (x.match(/([\d,]+)\s*$/) || [])[1] || null;
}

function pickLastNumberFromLines_(l, k) {
  const x = findLineStartsWith_(l, k);
  return x ? (x.match(/([\d,]+)\s*$/) || [])[1] : null;
}

function findLineStartsWith_(l, k) {
  k = k.toLowerCase();
  return l.find(x => x.toLowerCase().startsWith(k)) || null;
}

function findIndexStartsWith_(l, k) {
  k = k.toLowerCase();
  return l.findIndex(x => x.toLowerCase().startsWith(k));
}

function parseCityStateZipLoose_(l) {
  const m = l.match(/^(.+?),\s*([A-Z]{2})\s+(\d{5})$/i);
  return { city: m ? m[1] : null, state: m ? m[2] : null, zip: m ? m[3] : null };
}

function extractProposalLink_(b) {
  const m = b.match(/https?:\/\/hvacbizpro\.com\/proposals\/hardcopy\/\S+/i);
  return m ? m[0] : null;
}

function normalizeBody_(plain) {
  let t = String(plain || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  t = t.replace(/\t+/g, " ");
  t = t.replace(/[ ]{2,}/g, " ");
  return t.trim();
}

function logExec_(tag, msg) {
  const line = `${new Date().toISOString()} | ${tag} | ${msg}`;
  Logger.log(line);
  console.log(line);
}

function getSpreadsheet_() {
  return HVACBIZPRO.SPREADSHEET_ID
    ? SpreadsheetApp.openById(HVACBIZPRO.SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet_(ss, name, headers2d) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  const headerLen = headers2d[0].length;
  sh.getRange(1, 1, 1, headerLen).setValues(headers2d);
  sh.setFrozenRows(1);

  return sh;
}

function appendRows_(sheet, rows2d) {
  if (!rows2d || rows2d.length === 0) return;
  sheet.getRange(sheet.getLastRow() + 1, 1, rows2d.length, rows2d[0].length).setValues(rows2d);
}

function loadExistingIdsFromColumn_(sheet, colIndex) {
  const lastRow = sheet.getLastRow();
  const set = new Set();
  if (lastRow < 2) return set;

  const vals = sheet.getRange(2, colIndex, lastRow - 1, 1).getValues();
  for (const [v] of vals) {
    const s = String(v ?? "").trim();
    if (s) set.add(s);
  }
  return set;
}