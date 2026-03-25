/**
 * Fast refresh using one batch write via Sheets Advanced Service.
 * Row 1 preserved. Writes A2:I in a single call. Clears tail implicitly.
 * DEST: "Actual Costs" tab in spreadsheet 1MnkDZSNcNR4RFiPE3-QfSVtEvZ_M_qS_KdRHGiFyAvY
 */
function refreshMaterialCostImport() {
  const DEST_FILE_ID = '1MnkDZSNcNR4RFiPE3-QfSVtEvZ_M_qS_KdRHGiFyAvY';
  const DEST_SHEET_NAME = 'Material Cost Import';

  const SOURCE_IDS = [
    '1-HFrkzF3Mqf0WxZhoSXEArc_3ZSzWT6NB-XBjD0bl9Q',
    '1OSdrTRwBv4ZRWbueTM7lPguID8d9OehFJkQEimNlstU',
    '1_frrVtIqkfg5hVtglhHTNK8yIO7TnUO4liC8uW0pQAE',
  ];
  const KEEP_SET = new Set(['Equipment', 'Stock Equipment', 'Special Order']);

  const t0 = Date.now();

  // ---- Destination spreadsheet/sheet (NOT the active spreadsheet)
  const destSS = SpreadsheetApp.openById(DEST_FILE_ID);
  const dst =
    destSS.getSheetByName(DEST_SHEET_NAME) || destSS.insertSheet(DEST_SHEET_NAME);

  function safe(v) {
    return v == null ? '' : v;
  }

  function fmtMDY(d) { // MM/DD/YYYY
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    const y = d.getFullYear();
    return `${m}/${day}/${y}`;
  }

  function normalizeDatesInMatrix(rows) {
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      for (let c = 0; c < row.length; c++) {
        const v = row[c];
        if (v instanceof Date && !isNaN(v)) row[c] = fmtMDY(v);
      }
    }
    return rows;
  }

  function retry_(label, fn, tries = 4, baseMs = 300) {
    let last;
    for (let i = 0; i < tries; i++) {
      try {
        const t = Date.now();
        const r = fn();
        console.log(`[time] ${label} ${Date.now() - t}ms try${i + 1}/${tries}`);
        return r;
      } catch (e) {
        last = e;
        const wait = Math.min(baseMs * Math.pow(2, i), 5000);
        console.log(`[retry] ${label} ${i + 1}/${tries} wait ${wait}ms`, String(e));
        Utilities.sleep(wait);
      }
    }
    throw last;
  }

  // ---- Build columns in memory (no writes yet)
  const A = [], B = [], C = [], D = [], E = [], F = [], G = [], H = [], I = [];

  SOURCE_IDS.forEach((id, idx) => {
    const srcLabel = `src#${idx + 1}`;
    const src = retry_(`${srcLabel}: openById`, () => SpreadsheetApp.openById(id));

    // Released: B..Q → B,E,H,Q
    retry_(`${srcLabel}: Released B..Q`, () => {
      const sh = src.getSheetByName('Released'); if (!sh) return;
      const last = sh.getLastRow(); if (last < 3) return;
      const block = sh.getRange(3, 2, last - 2, 16).getValues();
      block.forEach(r => {
        I.push(safe(r[0]));   // B
        B.push(safe(r[3]));   // E
        A.push(safe(r[6]));   // H
        C.push(safe(r[15]));  // Q
      });
    });

    // Received: B..U → B,E,H,T,U
    retry_(`${srcLabel}: Received B..U`, () => {
      const sh = src.getSheetByName('Received'); if (!sh) return;
      const last = sh.getLastRow(); if (last < 3) return;
      const block = sh.getRange(3, 2, last - 2, 20).getValues();
      block.forEach(r => {
        const b = safe(r[0]);            // B
        const cat = String(safe(r[3]));  // E
        const h = safe(r[6]);            // H
        const t = safe(r[18]);           // T
        const u = safe(r[19]);           // U
        F.push(h);
        G.push(t);
        if (KEEP_SET.has(cat)) {
          D.push(h);
          E.push(u);
          H.push(b);
        } else {
          D.push('');
          E.push('');
          H.push('');
        }
      });
    });
  });

  // ---- Assemble one rectangular matrix A2:I and write once
  const prevLast = dst.getLastRow(); // includes header
  const newHeights = [A.length, B.length, C.length, D.length, E.length, F.length, G.length, H.length, I.length];
  const newMax = Math.max(...newHeights);
  const prevHeight = Math.max(0, prevLast - 1);
  const height = Math.max(newMax, prevHeight);

  const rows2D = new Array(height);
  for (let r = 0; r < height; r++) {
    rows2D[r] = [
      A[r] ?? '', B[r] ?? '', C[r] ?? '',
      D[r] ?? '', E[r] ?? '', F[r] ?? '',
      G[r] ?? '', H[r] ?? '', I[r] ?? ''
    ];
  }

  // Force MM/DD/YYYY for any Date cells
  normalizeDatesInMatrix(rows2D);

  // Single call via Advanced Service. Overwrites A2:I(1+height) in one request.
  const range = `'${DEST_SHEET_NAME}'!A2:I${height + 1}`;
  retry_('values.update A2:I', () =>
    Sheets.Spreadsheets.Values.update(
      { range, majorDimension: 'ROWS', values: rows2D },
      DEST_FILE_ID,
      range,
      { valueInputOption: 'RAW' }
    ),
    3, 400
  );

  const totalMillis = Date.now() - t0;
  console.log(`[summary] totalMillis: ${totalMillis}, rows: ${height}`);
}
