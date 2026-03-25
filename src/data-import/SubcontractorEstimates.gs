/**
 * Subcontractor Estimates (PMT IMPORT OVERRIDE FIRST, RAW EMAIL BACKUP) + INSTALL ADDRESS -> COL B
 *
 * TARGET SHEET: "Subcontractor Estimates"
 * A Job Name | B Install Address | C Electrical ($) | D Drilling ($) | E Permitting ($) | F Labor ($) | G Snippets
 *
 * OVERRIDE RULE (FIRST):
 * 1) Look for exact matching Job Name in column A of tab "PMT Import"
 * 2) If found AND AE (Drilling) is NON-ZERO:
 *    - Set Drilling ($) = value from column AE of PMT Import (same row)
 *    - Set Electrical/Permitting/Labor = 0
 *    - Set B = "-"
 *    - Set Snippets = "PMT Import AE"
 *    - Next job
 * 3) If found but AE is zero/blank: IGNORE PMT and parse emails normally.
 *
 * EMAIL RULE (BACKUP):
 * - Use sheet "HVACBizPro Raw Emails" (D subject, E body)
 * - Pick ONE best matching email (no summing across many emails)
 * - Parse priced items "@ $amount" and categorize by canonical similarity
 * - Extract Install Address from the chosen body and write to column B
 *
 * ADDRESS RULE (YOUR CHANGE):
 * - Find "Install Address"
 * - Take the text AFTER the LAST COMMA within that Install Address block
 *   Example: "123 Main St, Town, NY 12345" -> "NY 12345"
 *
 * LOGGING:
 * - Execution log only (console.log)
 */

function updateSubcontractorEstimatesFromRawEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const SC_EST_TARGET_SHEET_NAME = "Subcontractor Estimates";
  const SC_EST_RAW_SHEET_NAME = "HVACBizPro Raw Emails";
  const SC_EST_PMT_SHEET_NAME = "PMT Import";

  const targetSheet = ss.getSheetByName(SC_EST_TARGET_SHEET_NAME);
  if (!targetSheet) throw new Error("Missing sheet: " + SC_EST_TARGET_SHEET_NAME);

  const rawSheet = ss.getSheetByName(SC_EST_RAW_SHEET_NAME);
  if (!rawSheet) throw new Error("Missing sheet: " + SC_EST_RAW_SHEET_NAME);

  const pmtSheet = ss.getSheetByName(SC_EST_PMT_SHEET_NAME);
  if (!pmtSheet) throw new Error("Missing sheet: " + SC_EST_PMT_SHEET_NAME);

  SC_EST_ensureHeaders_(targetSheet);

  const targetLastRow = SC_EST_findLastNonEmptyRowInColumn_(targetSheet, 1, 2);
  if (targetLastRow < 2) {
    console.log("No data rows (no Job Names in column A).");
    return;
  }

  const pmtLastRow = pmtSheet.getLastRow();
  const pmtMap = SC_EST_buildPmtJobToAeMap_(pmtSheet, pmtLastRow);

  const rawLastRow = rawSheet.getLastRow();
  const rawValues = (rawLastRow >= 2)
    ? rawSheet.getRange(2, 4, rawLastRow - 1, 2).getValues()
    : [];

  console.log(
    `RUN_START updateSubcontractorEstimatesFromRawEmails | TargetRows=2-${targetLastRow} | ` +
    `PMT_Entries=${Object.keys(pmtMap).length} | RawEmails=${rawValues.length}`
  );

  for (let r = 2; r <= targetLastRow; r++) {
    const jobName = String(targetSheet.getRange(r, 1).getDisplayValue() || "").trim();

    if (!jobName) {
      targetSheet.getRange(r, 2, 1, 6).setValues([["-", "-", "-", "-", "-", "No Subcontractor"]]);
      console.log(`Row ${r} | Job="" | No Subcontractor`);
      continue;
    }

    const pmtKey = SC_EST_normKey_(jobName);
    if (Object.prototype.hasOwnProperty.call(pmtMap, pmtKey)) {
      const drillingVal = pmtMap[pmtKey];
      if (drillingVal !== 0) {
        targetSheet.getRange(r, 2, 1, 6).setValues([["-", 0, drillingVal, 0, 0, "PMT Import AE"]]);
        console.log(`Row ${r} | Job="${jobName}" | Mode=PMT_IMPORT | Drill=${drillingVal}`);
        continue;
      } else {
        console.log(`Row ${r} | Job="${jobName}" | PMT found but AE=0 -> ignoring PMT, using EMAIL parse`);
      }
    }

    const chosen = SC_EST_pickAndParseBestEmailForQuery_(jobName, rawValues, {
      requireLastNameInSubject: true,
      minTokenHits: 2,
      maxLineItem: 50000
    });

    const installAddr = chosen.installAddress || "-";

    const rowValues = SC_EST_hasAnyAmounts_(chosen.agg)
      ? [installAddr, chosen.agg.elec, chosen.agg.drill, chosen.agg.permit, chosen.agg.labor, SC_EST_dedupeSnippets_(chosen.agg.snips).join(" | ")]
      : [installAddr, "-", "-", "-", "-", "No Subcontractor"];

    targetSheet.getRange(r, 2, 1, 6).setValues([rowValues]);

    console.log(
      `Row ${r} | Job="${jobName}" | Mode=EMAIL_job | Matches=${chosen.matches} | BestScore=${chosen.bestScore} | ` +
      `AddrFound=${installAddr !== "-" ? "Y" : "N"} | Addr="${installAddr}" | ` +
      `Elec=${rowValues[1]} Drill=${rowValues[2]} Permit=${rowValues[3]} Labor=${rowValues[4]}`
    );
  }

  console.log("RUN_END updateSubcontractorEstimatesFromRawEmails");
}

/* ======================== PMT IMPORT MAP ======================== */

function SC_EST_buildPmtJobToAeMap_(pmtSheet, pmtLastRow) {
  const map = Object.create(null);
  if (pmtLastRow < 2) return map;

  const numCols = 31; // A..AE
  const vals = pmtSheet.getRange(2, 1, pmtLastRow - 1, numCols).getDisplayValues();

  for (let i = 0; i < vals.length; i++) {
    const job = String(vals[i][0] || "").trim();      // A
    const aeStr = String(vals[i][30] || "").trim();   // AE
    if (!job) continue;

    const key = SC_EST_normKey_(job);
    const aeNum = SC_EST_parseMoney_(aeStr);

    if (!Object.prototype.hasOwnProperty.call(map, key)) map[key] = aeNum;
    else if (aeNum !== 0) map[key] = aeNum;
  }

  return map;
}

function SC_EST_parseMoney_(s) {
  const x = String(s || "").replace(/[^0-9.\-]/g, "");
  const n = Number(x);
  return isFinite(n) ? n : 0;
}

function SC_EST_normKey_(s) {
  return String(s || "").toLowerCase().replace(/\s+/g, " ").trim();
}

/* ======================== PICK ONE BEST EMAIL (SCORING) ======================== */

function SC_EST_pickAndParseBestEmailForQuery_(querySource, rawValues, opts) {
  const options = opts || {};
  const minTokenHits = Number(options.minTokenHits || 2);
  const requireLastNameInSubject = !!options.requireLastNameInSubject;
  const maxLineItem = Number(options.maxLineItem || 50000);

  const queryTokens = SC_EST_tokenize_(querySource).slice(0, 6);
  const lastNameToken = SC_EST_extractLastNameToken_(querySource);

  let matches = 0;
  let bestIdx = -1;
  let bestScore = -1;

  for (let i = 0; i < rawValues.length; i++) {
    const subj = String(rawValues[i][0] || "");
    const body = String(rawValues[i][1] || "");
    if (!subj && !body) continue;

    const subjLower = subj.toLowerCase();
    const hayLower = (subj + " " + body).toLowerCase();

    if (subjLower.indexOf("financial summary :") === -1) continue;
    if (hayLower.indexOf("@") === -1) continue;
    if (hayLower.indexOf("$") === -1) continue;

    if (requireLastNameInSubject && lastNameToken) {
      if (subjLower.indexOf(lastNameToken) === -1) continue;
    }

    let hits = 0;
    let subjHits = 0;
    for (const t of queryTokens) {
      if (hayLower.indexOf(t) !== -1) hits++;
      if (subjLower.indexOf(t) !== -1) subjHits++;
    }
    if (hits < minTokenHits) continue;

    matches++;

    const lnBonus = (lastNameToken && subjLower.indexOf(lastNameToken) !== -1) ? 3 : 0;
    const score = (subjHits * 5) + (hits * 2) + lnBonus;

    if (score > bestScore) {
      bestScore = score;
      bestIdx = i;
    }
  }

  const agg = SC_EST_initAgg_();
  if (bestIdx === -1) return { agg, matches, bestScore: 0, installAddress: "" };

  const bestSubj = String(rawValues[bestIdx][0] || "");
  const bestBody = String(rawValues[bestIdx][1] || "");

  const installAddress = SC_EST_extractInstallAddressAfterLastComma_(bestBody);

  const parsed = SC_EST_categorizeAndSumByExactLines_(bestBody, maxLineItem);
  agg.elec = parsed.electricalTotal;
  agg.drill = parsed.drillingTotal;
  agg.permit = parsed.permittingTotal;
  agg.labor = parsed.laborTotal;
  agg.snips = parsed.snippets;

  console.log(
    `BestEmail | Query="${querySource}" | LastName="${lastNameToken || ""}" | ` +
    `BestScore=${bestScore} | Subject="${bestSubj}" | InstallAddress="${installAddress || ""}"`
  );

  return { agg, matches, bestScore, installAddress };
}

function SC_EST_extractInstallAddressAfterLastComma_(body) {
  const full = SC_EST_extractInstallAddressBlock_(body);
  if (!full) return "";

  // take text after the last comma
  const lastComma = full.lastIndexOf(",");
  if (lastComma === -1) return full.trim();

  return full.slice(lastComma + 1).trim();
}

/**
 * Extracts the "Install Address" block as a single line (may join up to 3 lines).
 */
function SC_EST_extractInstallAddressBlock_(body) {
  const text = String(body || "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n");

  const lower = text.toLowerCase();
  const idx = lower.indexOf("install address");
  if (idx === -1) return "";

  let after = text.slice(idx + "install address".length);
  after = after.replace(/^[ \t]*[:\-–—]?[ \t]*/g, "");

  const lines = after.split("\n").map(l => l.trim());
  const stopLabelRe = /^(billing address|mailing address|install date|customer|phone|email|job name|project|scope|financial summary|total|notes)\b/i;

  const out = [];
  for (let i = 0; i < lines.length && out.length < 3; i++) {
    const l = lines[i];
    if (!l) continue;

    if (stopLabelRe.test(l)) break;
    if (/^[A-Za-z][A-Za-z \t]{2,30}:\s*/.test(l)) break;

    out.push(l);
  }

  if (!out.length) {
    const m = text.match(/Install Address\s*[:\-–—]\s*([^\n\r]+)/i);
    if (m && m[1]) return String(m[1]).trim();
    return "";
  }

  return out.join(", ").replace(/\s+/g, " ").trim();
}

function SC_EST_extractLastNameToken_(jobName) {
  const toks = SC_EST_tokenize_(jobName);
  if (!toks.length) return "";

  const stop = new Set(["install", "job", "project", "street", "avenue", "road", "drive", "lane", "court", "place", "unit"]);
  for (const t of toks) {
    if (t.length < 3) continue;
    if (stop.has(t)) continue;
    if (/^\d+$/.test(t)) continue;
    return t;
  }
  return toks[0] || "";
}

/* ======================== EXACT LINE MATCH (SLIGHT VARIATION) ======================== */

function SC_EST_categorizeAndSumByExactLines_(text, maxLineItem) {
  const normalized = String(text || "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/[ \t]+/g, " ")
    .trim();

  const items = SC_EST_extractPricedItems_(normalized, maxLineItem);

  let electricalTotal = 0;
  let drillingTotal = 0;
  let permittingTotal = 0;
  let laborTotal = 0;

  const snippets = [];

  for (const it of items) {
    const cat = SC_EST_categoryFromCanonicalSimilarity_(it.itemText);
    if (!cat) continue;

    if (cat === "electrical") electricalTotal += it.amount;
    if (cat === "drilling") drillingTotal += it.amount;
    if (cat === "permitting") permittingTotal += it.amount;
    if (cat === "labor") laborTotal += it.amount;

    snippets.push(`${cat.toUpperCase()}: ${it.rawSnippet}`);
  }

  return { electricalTotal, drillingTotal, permittingTotal, laborTotal, snippets };
}

function SC_EST_extractPricedItems_(body, maxLineItem) {
  const out = [];
  const cap = isFinite(maxLineItem) ? maxLineItem : 50000;
  const re = /@\s*\$?\s*([0-9][0-9,]*(?:\.[0-9]{1,2})?)/g;

  let m;
  while ((m = re.exec(body)) !== null) {
    const amount = Number(String(m[1]).replace(/,/g, ""));
    if (!isFinite(amount)) continue;
    if (amount > cap) continue;

    const atIndex = m.index;
    const startWindow = Math.max(0, atIndex - 260);
    let before = body.slice(startWindow, atIndex).trim();

    const lastComma = before.lastIndexOf(",");
    if (lastComma >= 0 && lastComma > before.length - 220) {
      before = before.slice(lastComma + 1).trim();
    }

    const rawSnippet = `${before} @ $${m[1]}`.replace(/\s+/g, " ").trim();
    out.push({ itemText: before, amount, rawSnippet });
  }

  return out;
}

function SC_EST_categoryFromCanonicalSimilarity_(segment) {
  const segNorm = SC_EST_normalizeLine_(segment);
  const CANON = SC_EST_getCanonicalLines_();

  const TH = { electrical: 0.88, drilling: 0.88, permitting: 0.80, labor: 0.90 };
  const MUST = {
    electrical: ["ELECTRICAL", "PERMIT"],
    drilling: ["GEOTHERMAL", "SUBCONTRACTOR"],
    permitting: ["PERMIT", "PROCESS", "HOMEOWNER"],
    labor: []
  };

  if (SC_EST_passesMustWords_(segNorm, MUST.electrical) && SC_EST_isSimilarEnough_(segNorm, SC_EST_normalizeLine_(CANON.electrical), TH.electrical)) {
    return "electrical";
  }
  if (SC_EST_passesMustWords_(segNorm, MUST.drilling) && SC_EST_isSimilarEnough_(segNorm, SC_EST_normalizeLine_(CANON.drilling), TH.drilling)) {
    return "drilling";
  }
  if (SC_EST_passesMustWords_(segNorm, MUST.permitting) && SC_EST_isSimilarEnough_(segNorm, SC_EST_normalizeLine_(CANON.permitting), TH.permitting)) {
    return "permitting";
  }

  for (const l of CANON.laborLines) {
    const ln = SC_EST_normalizeLine_(l);
    if (SC_EST_isSimilarEnough_(segNorm, ln, TH.labor)) return "labor";
  }
  return null;
}

function SC_EST_normalizeLine_(s) {
  return String(s || "")
    .toUpperCase()
    .replace(/[“”]/g, '"')
    .replace(/[’]/g, "'")
    .replace(/\s+/g, " ")
    .trim();
}

function SC_EST_passesMustWords_(segNorm, mustWords) {
  if (!mustWords || !mustWords.length) return true;
  let hits = 0;
  for (const w of mustWords) if (segNorm.indexOf(w) !== -1) hits++;
  const need = Math.min(2, mustWords.length);
  return hits >= need;
}

function SC_EST_isSimilarEnough_(aNorm, bNorm, threshold) {
  if (!aNorm || !bNorm) return false;
  if (aNorm.indexOf(bNorm) !== -1 || bNorm.indexOf(aNorm) !== -1) return true;
  const sim = SC_EST_normalizedSimilarity_(aNorm, bNorm);
  return sim >= threshold;
}

function SC_EST_normalizedSimilarity_(a, b) {
  const dist = SC_EST_levenshtein_(a, b);
  const maxLen = Math.max(a.length, b.length) || 1;
  return 1 - (dist / maxLen);
}

function SC_EST_levenshtein_(a, b) {
  const m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;

  const dp = new Array(n + 1);
  for (let j = 0; j <= n; j++) dp[j] = j;

  for (let i = 1; i <= m; i++) {
    let prev = dp[0];
    dp[0] = i;
    for (let j = 1; j <= n; j++) {
      const tmp = dp[j];
      const cost = a.charAt(i - 1) === b.charAt(j - 1) ? 0 : 1;
      dp[j] = Math.min(dp[j] + 1, dp[j - 1] + 1, prev + cost);
      prev = tmp;
    }
  }
  return dp[n];
}

function SC_EST_getCanonicalLines_() {
  return {
    electrical: "ELECTRICAL PERMIT AND THIRD PARTY INSPECTION INCLUDED (GFI NOT INCLUDED)",
    drilling: "PROVIDE GEOTHERMAL WELLS WITH A LICENSED GEOTHERMAL SUBCONTRACTOR",
    permitting:
      "PERMIT PROCESS: WE WILL APPLY FOR THE PERMIT AND PASS THE COST TO THE HOMEOWNER ONCE RECEIVED. " +
      "ELECTRICAL/PLUMBING PERMITS, ENGINEERING STAMPS/DRAWINGS, AND DUCT INTEGRITY TESTS ARE NOT INCLUDED IF REQUIRED BY THE BUILDING DEPARTMENT.",
    laborLines: [
      "DECOMMISSIONING OF EXISTING HEATING SYSTEM FOR REBATE MAXIMIZATION - ADDITIONAL COST",
      "REMOVE AND DISPOSE OF PROPERLY EXISTING UNIT",
      "RE-CONNECT EXISTING DUCTWORK",
      "RE-CONNECT TO EXISTING PIPING, DUCTWORK AND ELECTRICAL",
      "RE-CONNECT EXISTING REFRIGERANT PIPING AND FLUSH OUT",
      "RE-CONNECT EXISTING OIL PIPING",
      "RE-CONNECT EXISTING FLUE PIPING",
      "RE-CONNECT EXISTING GAS PIPING",
      "RE-CONNECT EXISTING DRAIN PIPING",
      "START AND TEST COMPLETED INSTALLATION"
    ]
  };
}

/* ======================== HEADERS ======================== */

function SC_EST_ensureHeaders_(sheet) {
  const headers = ["Job Name", "Install Address", "Electrical ($)", "Drilling ($)", "Permitting ($)", "Labor ($)", "Snippets"];
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];

  let mismatch = current.length !== headers.length;
  if (!mismatch) {
    for (let i = 0; i < headers.length; i++) {
      if (current[i] !== headers[i]) { mismatch = true; break; }
    }
  }
  if (mismatch) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

/* ======================== SNIPPET UTILS ======================== */

function SC_EST_dedupeSnippets_(snips) {
  const seen = new Set();
  const out = [];
  for (const s of snips || []) {
    const k = String(s).toLowerCase();
    if (seen.has(k)) continue;
    seen.add(k);
    out.push(s);
    if (out.length >= 12) break;
  }
  return out;
}

/* ======================== TOKENS / AGG / SHEET UTILS ======================== */

function SC_EST_tokenize_(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/[^\w\s]/g, " ")
    .split(/\s+/)
    .filter(w => w && w.length >= 3);
}

function SC_EST_initAgg_() {
  return { elec: 0, drill: 0, permit: 0, labor: 0, snips: [] };
}

function SC_EST_hasAnyAmounts_(agg) {
  return (agg.elec > 0) || (agg.drill > 0) || (agg.permit > 0) || (agg.labor > 0);
}

function SC_EST_findLastNonEmptyRowInColumn_(sheet, col, startRow) {
  const max = sheet.getMaxRows();
  const vals = sheet.getRange(startRow, col, max - startRow + 1, 1).getDisplayValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0] || "").trim() !== "") return startRow + i;
  }
  return 0;
}
