/**
 * BizPro Line Item Breakdown
 *
 * Parses the "HVACBizPro Raw Emails" sheet and extracts itemized line items
 * from Financial Summary emails into a new "Line Item Breakdown" tab.
 *
 * Each email's body has a table with four sections:
 *   - Units Added       → category "Equipment"
 *   - Special Parts     → category "Materials"
 *   - Common Parts      → category "Materials"
 *   - Services Added    → category "Services"
 *
 * Each line item has the pattern: DESCRIPTION @ $AMOUNT
 *
 * Output columns:
 *   A = Message ID | B = Gmail Date | C = Customer Name | D = Proposal
 *   E = Category (Equipment / Materials / Services) | F = Description | G = Amount
 */

const LIB_CONFIG = {
  SHEET_RAW: 'HVACBizPro Raw Emails',
  SHEET_JOBS: 'HVACBizPro Jobs',
  SHEET_OUTPUT: 'Line Item Breakdown',
};

function generateLineItemBreakdown() {
  const t0 = Date.now();
  console.log('generateLineItemBreakdown() start');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(LIB_CONFIG.SHEET_RAW);
  if (!rawSheet) throw new Error('Missing sheet: ' + LIB_CONFIG.SHEET_RAW);

  const jobsSheet = ss.getSheetByName(LIB_CONFIG.SHEET_JOBS);

  const rawLastRow = rawSheet.getLastRow();
  if (rawLastRow < 2) {
    console.log('No raw email rows.');
    return;
  }

  // Raw: A=Message ID, B=Gmail Date, C=From, D=Subject, E=Plain Body
  const rawData = rawSheet.getRange(2, 1, rawLastRow - 1, 5).getValues();

  // Jobs: A=Message ID, D=Proposal, I=Customer Name (for lookup)
  let jobLookup = {};
  if (jobsSheet) {
    const jobsLastRow = jobsSheet.getLastRow();
    if (jobsLastRow >= 2) {
      const jobsData = jobsSheet.getRange(2, 1, jobsLastRow - 1, 9).getValues();
      for (const row of jobsData) {
        const msgId = String(row[0] || '').trim();
        if (msgId) {
          jobLookup[msgId] = {
            proposal: String(row[3] || '').trim(),
            customerName: String(row[8] || '').trim(),
          };
        }
      }
    }
  }

  const results = [];

  for (const row of rawData) {
    const msgId = String(row[0] || '').trim();
    const gmailDate = row[1];
    const subject = String(row[3] || '').trim();
    const body = normalizeBody_(String(row[4] || ''));

    if (!body || !body.toUpperCase().includes('PROPOSAL')) continue;

    const job = jobLookup[msgId] || {};
    const customerName = job.customerName || extractCustomerFromSubject_(subject);
    const proposal = job.proposal || '';

    const items = parseLineItems_(body);

    for (const item of items) {
      results.push([
        msgId,
        gmailDate,
        customerName,
        proposal,
        item.category,
        item.description,
        item.amount,
      ]);
    }
  }

  // Write output
  let outputSheet = ss.getSheetByName(LIB_CONFIG.SHEET_OUTPUT);
  if (!outputSheet) {
    outputSheet = ss.insertSheet(LIB_CONFIG.SHEET_OUTPUT);
  } else {
    outputSheet.clearContents();
  }

  const headers = [['Message ID', 'Gmail Date', 'Customer Name', 'Proposal',
                     'Category', 'Description', 'Amount']];
  outputSheet.getRange(1, 1, 1, 7).setValues(headers);
  outputSheet.setFrozenRows(1);

  if (results.length > 0) {
    outputSheet.getRange(2, 1, results.length, 7).setValues(results);
  }

  const elapsed = Date.now() - t0;
  console.log(`generateLineItemBreakdown() complete. Items=${results.length} Elapsed=${elapsed}ms`);
  SpreadsheetApp.getActive().toast(
    `Done: ${results.length} line items written to "${LIB_CONFIG.SHEET_OUTPUT}".`,
    'Line Item Breakdown',
    5
  );
}

/* ======================== PARSER ======================== */

/**
 * Parses the plain text body and returns an array of
 * { category: "Equipment"|"Materials"|"Services", description: String, amount: Number }
 *
 * The email body has section headers that appear as a group:
 *   Units Added / Special Parts / Common Parts / Services Added
 * Then line items follow in that same order, separated by "None" for empty sections.
 *
 * Strategy: find the section header block, then walk forward through lines,
 * tracking which section we're in based on item flow.
 */
function parseLineItems_(body) {
  const lines = body.split('\n').map(l => l.trim()).filter(Boolean);
  const items = [];

  // Find the header block: "Units Added" must appear
  const unitsIdx = findIndexStartsWith_(lines, 'Units Added');
  if (unitsIdx < 0) return items;

  // Find where item lines start (after the header block)
  // Headers are: Units Added, Special Parts, Common Parts, Services Added
  // They appear on consecutive lines, then items follow
  let itemStart = unitsIdx + 1;
  const headerLabels = ['special parts', 'common parts', 'services added'];
  for (let i = unitsIdx + 1; i < Math.min(unitsIdx + 5, lines.length); i++) {
    const lower = lines[i].toLowerCase();
    if (headerLabels.some(h => lower.startsWith(h))) {
      itemStart = i + 1;
    }
  }

  // Walk lines from itemStart until we hit the financial summary section
  // ("Category Sold" or "Selling Price" or "%" signals the end)
  const stopRe = /^(category sold|selling price|%\s*$|\$\s*value)/i;

  // We need to figure out which section each item belongs to.
  // The HTML table renders columns in order: Units Added | Special Parts | Common Parts | Services Added
  // In plain text, items from each column appear in order.
  // We detect section boundaries by content patterns:
  //   - Equipment items: model numbers with (Catalogue/Dist Part Number)
  //   - Materials items: description with "KIT", "MATERIAL", "LINESET", etc.
  //   - Services items: action descriptions (REMOVE, CONNECT, PROVIDE, FURNISH, INSTALL, etc.)
  //   - "None" marks an empty section

  // Simpler approach: gather all priced lines, then categorize each one
  for (let i = itemStart; i < lines.length; i++) {
    const line = lines[i];
    if (stopRe.test(line)) break;
    if (/^none$/i.test(line)) continue;

    // Extract items with @ $amount pattern (can be multiple per line separated by commas + newlines)
    const re = /([^,@]+?)(?:\*\s*\d+\s*)?@\s*\$?\s*([0-9][0-9,]*(?:\.\d{1,2})?)/g;
    let m;
    while ((m = re.exec(line)) !== null) {
      const desc = m[1].replace(/^\s*[☐☑✓]?\s*/, '').trim();
      const amount = Number(m[2].replace(/,/g, ''));
      if (!isFinite(amount)) continue;

      const category = categorizeLineItem_(desc);
      items.push({ category, description: desc, amount });
    }
  }

  return items;
}

/**
 * Categorize a line item description as Equipment, Materials, or Services.
 */
function categorizeLineItem_(desc) {
  const upper = desc.toUpperCase();

  // Equipment: model numbers — typically short alphanumeric codes with digits,
  // and/or contain "(Catalogue/Dist Part Number"
  if (/CATALOGUE|DIST PART NUMBER/i.test(upper)) return 'Equipment';
  // Model number heuristic: starts with letters+digits, short, no common service verbs
  if (/^[A-Z]{2,}[0-9]/.test(upper) && upper.length < 60) return 'Equipment';

  // Materials: kits, linesets, physical supplies
  if (/\bKIT\b|\bMATERIAL|\bLINESET|\bPIPING\b|\bWIRE\b|\bCONDENSATE PUMP\b|\bCONDENSER STAND\b|\bCONDENSER PAD\b|\bFOAM BLOCK|\bLINEHIDE\b|\bDRAIN\b.*\bPIPE\b|\bTHERMOSTAT\b|\bPAN\b|\bRISED PAD\b|\bDISCONNECT\b|\bWHIP\b/i.test(upper)) {
    return 'Materials';
  }

  // Services: action verbs
  if (/^(REMOVE|RE-CONNECT|RECONNECT|CONNECT|PROVIDE|FURNISH|INSTALL|START AND TEST|DECOMMISSION|PERMIT|POWER WIRING|ELECTRICAL(?! WHIP))/i.test(upper)) {
    return 'Services';
  }
  if (/\bINSTALLATION\b|\bREBATE\b|\bFINANCING\b|\bPROMOTION|\bFLOOR PROTECTION\b|\bSAFETY SWITCH\b/i.test(upper)) {
    return 'Services';
  }

  // Fallback: if it has @ $0.00 it's probably a service note
  return 'Services';
}

/* ======================== HELPERS ======================== */

function extractCustomerFromSubject_(subject) {
  // Subject format: "Financial Summary : LASTNAME"
  const m = subject.match(/Financial Summary\s*:\s*(.+)$/i);
  return m ? toProperCase_(m[1].trim()) : '';
}

function normalizeBody_(plain) {
  let t = String(plain || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  t = t.replace(/\t+/g, ' ');
  t = t.replace(/[ ]{2,}/g, ' ');
  return t.trim();
}

function toProperCase_(str) {
  if (!str) return str;
  return String(str)
    .toLowerCase()
    .split(/\s+/)
    .map(w => w ? w[0].toUpperCase() + w.slice(1) : '')
    .join(' ');
}

function findIndexStartsWith_(l, k) {
  k = k.toLowerCase();
  return l.findIndex(x => x.toLowerCase().startsWith(k));
}
