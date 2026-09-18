// ================================================================
//  BUSINESS CARD SCANNER — Google Apps Script Backend
//  Works as BOTH a Standalone script AND a Container-bound script.
// ================================================================

// ── CONFIGURATION ────────────────────────────────────────────────
// The Spreadsheet ID (used when running as a standalone script).
// When running as a container-bound script (opened from the Sheet),
// getSpreadsheet_() will automatically use the active spreadsheet instead.
const SHEET_ID        = '1NNCfLyqguar-IHcXBRplQMQ4ui0iGGTcu-oF8OTqVb0';

// Name of the output tab that will be created automatically
const OUT_TAB         = 'Scanned Cards';

// Gemini model: 'gemini-1.5-pro' (most accurate) or 'gemini-2.0-flash' (faster)
const GEMINI_MDL      = 'gemini-2.0-flash';

// GID of the Form Response sheet tab (from your sheet URL: gid=1981435111)
const FORM_SHEET_GID  = 1981435111;

// Which column (1-based) in the Form Response sheet has the image/file URL?
// Set to 0 to auto-detect (recommended). Set to a fixed number if auto-detect fails.
const IMG_COL         = 0;

// ── WEB APP ENTRY POINT ──────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Business Card Scanner')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const to = data.to;
    const subject = data.subject || 'Follow-Up';
    const body = data.body || '';
    const footer = data.footer || '';
    const fullBody = footer ? (body + '\n\n--\n' + footer) : body;
    
    GmailApp.sendEmail(to, subject, fullBody);
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'Email sent directly via GmailApp successfully!'
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ================================================================
//  FUNCTIONS CALLED FROM THE FRONTEND (google.script.run)
// ================================================================

/**
 * Scan a base64 image with Gemini Vision and return extracted card data.
 * Also checks for duplicates in the Output Sheet.
 */
function scanCard(base64Image, mimeType) {
  try {
    const cardData  = callGeminiVision_(base64Image, mimeType || 'image/jpeg');
    const dupResult = checkDuplicate_(cardData.phone || [], cardData.email || []);
    return { success: true, data: cardData, duplicate: dupResult };
  } catch (err) {
    Logger.log('scanCard error: ' + err.message);
    return { success: false, error: err.message };
  }
}

/**
 * Save (or update) card data to the Output Sheet.
 * Pass updateRowIndex to overwrite an existing row (for duplicate edits).
 */
function saveCard(cardData, updateRowIndex) {
  try {
    const result = writeToOutputSheet_(cardData, updateRowIndex || null);
    return { success: true, updated: result.updated, rowIndex: result.rowIndex };
  } catch (err) {
    Logger.log('saveCard error: ' + err.message);
    return { success: false, error: err.message };
  }
}

/**
 * Fetch the last N rows from the Output Sheet for the "Recently Scanned" table.
 */
function getRecentCards(limit) {
  try {
    const ss    = getSpreadsheet_();
    const sheet = ss.getSheetByName(OUT_TAB);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, rows: [] };

    const n        = Math.min(limit || 15, sheet.getLastRow() - 1);
    const startRow = sheet.getLastRow() - n + 1;
    const data     = sheet.getRange(startRow, 1, n, 15).getValues();

    const rows = data.reverse().map((r, i) => ({
      rowIndex:    sheet.getLastRow() - i,
      timestamp:   r[0] ? Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), 'dd MMM, HH:mm') : '',
      name:        r[1]  || '',
      designation: r[2]  || '',
      company:     r[3]  || '',
      phone:       r[4]  || '',
      email:       r[6]  || '',
      source:      r[12] || ''
    }));

    return { success: true, rows };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

/**
 * Manually process all unprocessed rows in the Form Response Sheet.
 * Skips any row whose image URL is already present in the Output Sheet.
 */
function processFormSheetNow() {
  try {
    const ss    = getSpreadsheet_();
    const forms = getFormResponseSheet_(ss);
    if (!forms) {
      return {
        success: false,
        error:   'Form Response sheet (GID ' + FORM_SHEET_GID + ') not found. ' +
                 'Make sure FORM_SHEET_GID matches the gid= value in your sheet URL.'
      };
    }
    const out = getOrCreateOutputSheet_(ss);

    if (forms.getLastRow() < 2) return { success: true, processed: 0 };

    const imgCol = IMG_COL > 0 ? IMG_COL : detectImageColumn_(forms);
    if (!imgCol) {
      return {
        success: false,
        error:   'Could not find an image/file-upload column in the Form Response sheet. ' +
                 'Run debugFormSheet() in the Apps Script editor to see your column headers.'
      };
    }

    // Build a set of already-processed image URLs (column N in output sheet)
    const processedUrls = new Set();
    if (out.getLastRow() > 1) {
      out.getRange(2, 14, out.getLastRow() - 1, 1).getValues()
         .forEach(r => { if (r[0]) processedUrls.add(r[0].toString().trim()); });
    }

    let processed = 0;
    const errors  = [];

    for (let row = 2; row <= forms.getLastRow(); row++) {
      const url = forms.getRange(row, imgCol).getValue().toString().trim();
      if (!url || processedUrls.has(url)) continue;

      try {
        const fileId   = extractDriveFileId_(url);
        const file     = DriveApp.getFileById(fileId);
        const blob     = file.getBlob();
        const b64      = Utilities.base64Encode(blob.getBytes());
        const mime     = blob.getContentType() || 'image/jpeg';
        const cardData = callGeminiVision_(b64, mime);
        cardData.source   = 'Google Form';
        cardData.imageUrl = url;
        writeToOutputSheet_(cardData, null);
        processed++;
        Utilities.sleep(700); // stay within Gemini rate limits
      } catch (rowErr) {
        const msg = `Row ${row}: ${rowErr.message}`;
        Logger.log('processFormSheetNow error — ' + msg);
        errors.push(msg);
        forms.getRange(row, imgCol).setNote('⚠️ Scan error: ' + rowErr.message);
      }
    }

    return { success: true, processed, errors };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ================================================================
//  GOOGLE FORM TRIGGER  (fires automatically on form submit)
// ================================================================

/**
 * Installed trigger — fires when the linked Form is submitted.
 * Run setupTrigger() ONCE from the Apps Script editor to install this.
 */
function onFormSubmit(e) {
  try {
    const row   = e.range.getRow();
    const sheet = e.range.getSheet();

    const imgCol = IMG_COL > 0 ? IMG_COL : detectImageColumn_(sheet);
    if (!imgCol) {
      Logger.log('onFormSubmit: no image column found in sheet "' + sheet.getName() + '"');
      return;
    }

    const imageUrl = sheet.getRange(row, imgCol).getValue().toString().trim();
    Logger.log('onFormSubmit row ' + row + ' — image URL: ' + imageUrl);

    if (!imageUrl) {
      Logger.log('onFormSubmit: empty image URL at row ' + row + ', col ' + imgCol);
      return;
    }

    const fileId   = extractDriveFileId_(imageUrl);
    const file     = DriveApp.getFileById(fileId);
    const blob     = file.getBlob();
    const b64      = Utilities.base64Encode(blob.getBytes());
    const mime     = blob.getContentType() || 'image/jpeg';

    const cardData    = callGeminiVision_(b64, mime);
    cardData.source   = 'Google Form';
    cardData.imageUrl = imageUrl;

    writeToOutputSheet_(cardData, null);
    Logger.log('onFormSubmit: ✅ processed row ' + row + ' — ' + (cardData.name || 'unknown name'));
  } catch (err) {
    Logger.log('onFormSubmit ERROR: ' + err.message + ' | Stack: ' + err.stack);
    try { e.range.getSheet().getRange(e.range.getRow(), 1).setNote('⚠️ Scan error: ' + err.message); } catch(_) {}
  }
}

// ================================================================
//  GEMINI VISION API
// ================================================================

function callGeminiVision_(base64, mimeType) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error(
      'GEMINI_API_KEY not set. Go to Apps Script → Project Settings → Script Properties → Add property → Key: GEMINI_API_KEY'
    );
  }

  const endpoint =
    `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_MDL}:generateContent?key=${apiKey}`;

  const prompt =
    `You are a business card OCR expert. Carefully extract every piece of visible text ` +
    `from this business card image and organize it into a structured JSON object.\n\n` +
    `Return ONLY a raw JSON object — no markdown, no code fences, no explanation.\n\n` +
    `Required structure:\n` +
    `{\n` +
    `  "name": "Full name of the person on the card",\n` +
    `  "designation": "Job title, role, or designation",\n` +
    `  "company": "Company or organization name",\n` +
    `  "phone": ["ALL phone numbers — mobile, office, fax, etc."],\n` +
    `  "email": ["ALL email addresses"],\n` +
    `  "website": "Website URL",\n` +
    `  "address": "Complete address as a single string",\n` +
    `  "linkedin": "LinkedIn URL or handle (if present)",\n` +
    `  "other": "Any other text — taglines, social handles, certifications, etc."\n` +
    `}\n\n` +
    `Rules:\n` +
    `1. Return ONLY the JSON — no markdown, no extra prose\n` +
    `2. Use null for any field not found on the card\n` +
    `3. phone and email MUST be arrays even for a single value\n` +
    `4. Preserve phone number formatting exactly as shown\n` +
    `5. Include ALL visible text — miss nothing, even small print\n` +
    `6. Do NOT guess or hallucinate data not visible on the card\n` +
    `7. For dense cards, read every line and corner carefully\n` +
    `8. Include company name even if it only appears in a logo`;

  const payload = {
    contents: [{
      parts: [
        { inlineData: { mimeType, data: base64 } },
        { text: prompt }
      ]
    }],
    generationConfig: {
      temperature:      0.05,
      maxOutputTokens:  2048,
      responseMimeType: 'application/json'
    }
  };

  const resp = UrlFetchApp.fetch(endpoint, {
    method:             'post',
    contentType:        'application/json',
    payload:            JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  if (code !== 200) {
    const snippet = resp.getContentText().slice(0, 500);
    throw new Error(`Gemini API error (HTTP ${code}): ${snippet}`);
  }

  const body = JSON.parse(resp.getContentText());
  const text = body?.candidates?.[0]?.content?.parts?.[0]?.text;
  if (!text) throw new Error('Gemini returned an empty response. Check your API key and try again.');

  return parseCardData_(text);
}

// ================================================================
//  SHEET OPERATIONS
// ================================================================

function writeToOutputSheet_(data, updateRowIndex) {
  const ss    = getSpreadsheet_();
  const sheet = getOrCreateOutputSheet_(ss);

  const row = [
    new Date(),
    data.name        || '',
    data.designation || '',
    data.company     || '',
    (data.phone  && data.phone[0])  || '',
    (data.phone  || []).join(' | '),
    (data.email  && data.email[0])  || '',
    (data.email  || []).join(' | '),
    data.website  || '',
    data.address  || '',
    data.linkedin || '',
    data.other    || '',
    data.source   || 'Webcam',
    data.imageUrl || '',
    ''   // Notes — left blank for manual use
  ];

  if (updateRowIndex) {
    sheet.getRange(updateRowIndex, 1, 1, row.length).setValues([row]);
    return { rowIndex: updateRowIndex, updated: true };
  }

  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1, 1, row.length).setValues([row]);

  // Light alternating row shading
  if (newRow % 2 === 0) {
    sheet.getRange(newRow, 1, 1, row.length).setBackground('#f8f9fa');
  }

  return { rowIndex: newRow, updated: false };
}

function checkDuplicate_(phones, emails) {
  const ss    = getSpreadsheet_();
  const sheet = ss.getSheetByName(OUT_TAB);
  if (!sheet || sheet.getLastRow() < 2) return { found: false };

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 15).getValues();

  for (let i = 0; i < rows.length; i++) {
    const existingPhones = (rows[i][5] || '').split(' | ');
    const existingEmails = (rows[i][7] || '').split(' | ');

    const phoneHit = phones.some(p =>
      p && existingPhones.some(ep => {
        const a = p.replace(/\D/g, ''), b = ep.replace(/\D/g, '');
        // Match last 8 digits to handle country-code variations
        return a.length > 5 && b.length > 5 && a.slice(-8) === b.slice(-8);
      })
    );

    const emailHit = emails.some(em =>
      em && existingEmails.some(ee =>
        ee.toLowerCase().trim() === em.toLowerCase().trim()
      )
    );

    if (phoneHit || emailHit) {
      return {
        found:    true,
        rowIndex: i + 2,
        existing: {
          name:    rows[i][1],
          company: rows[i][3],
          phone:   rows[i][4],
          email:   rows[i][6]
        }
      };
    }
  }

  return { found: false };
}

function getOrCreateOutputSheet_(ss) {
  let sheet = ss.getSheetByName(OUT_TAB);
  if (sheet) return sheet;

  sheet = ss.insertSheet(OUT_TAB);

  const headers = [
    'Timestamp', 'Name', 'Designation', 'Company',
    'Phone (Primary)', 'Phone (All)', 'Email (Primary)', 'Email (All)',
    'Website', 'Address', 'LinkedIn', 'Other Info',
    'Source', 'Image URL', 'Notes'
  ];

  const hdr = sheet.getRange(1, 1, 1, headers.length);
  hdr.setValues([headers]);
  hdr.setFontWeight('bold');
  hdr.setBackground('#1a73e8');
  hdr.setFontColor('#ffffff');
  hdr.setFontSize(11);
  sheet.setFrozenRows(1);
  sheet.setRowHeight(1, 32);

  const widths = [150, 170, 160, 170, 140, 220, 180, 220, 180, 270, 200, 200, 100, 250, 160];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));

  Logger.log('✅ Created output sheet: "' + OUT_TAB + '"');
  return sheet;
}

// ================================================================
//  HELPERS
// ================================================================

/**
 * Smart spreadsheet getter.
 * Uses the active spreadsheet if this is a container-bound script,
 * otherwise opens by SHEET_ID (standalone mode).
 */
function getSpreadsheet_() {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (_) {}
  return SpreadsheetApp.openById(SHEET_ID);
}

/**
 * Finds the Form Response sheet tab by its GID.
 * Also falls back to searching by tab name if GID not found.
 */
function getFormResponseSheet_(ss) {
  // Primary: find by GID
  const byGid = ss.getSheets().find(s => s.getSheetId() === FORM_SHEET_GID);
  if (byGid) return byGid;

  // Fallback: find by name containing "response" or "form"
  const byName = ss.getSheets().find(s => {
    const n = s.getName().toLowerCase();
    return n.includes('response') || n.includes('form responses');
  });
  if (byName) {
    Logger.log('getFormResponseSheet_: GID not found, using tab "' + byName.getName() + '" as fallback');
    return byName;
  }

  return null;
}

function extractDriveFileId_(url) {
  const patterns = [
    /\/file\/d\/([a-zA-Z0-9_-]{20,})/,
    /[?&]id=([a-zA-Z0-9_-]{20,})/,
    /\/open\?id=([a-zA-Z0-9_-]{20,})/,
    /\/d\/([a-zA-Z0-9_-]{20,})\//
  ];

  for (const pat of patterns) {
    const m = url.match(pat);
    if (m) return m[1];
  }

  // Bare file ID
  if (/^[a-zA-Z0-9_-]{25,}$/.test(url.trim())) return url.trim();

  throw new Error('Cannot extract Drive file ID from URL: ' + url);
}

/**
 * Auto-detects which column in the Form Response sheet contains image/Drive URLs.
 * First checks headers for keywords, then scans data rows for Drive URLs.
 */
function detectImageColumn_(sheet) {
  if (sheet.getLastColumn() < 1) return null;
  const lastCol = sheet.getLastColumn();

  // 1. Check header row for file/image/upload keywords
  if (sheet.getLastRow() >= 1) {
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const kw = ['file', 'image', 'upload', 'photo', 'card', 'picture', 'pic', 'scan'];
    for (let c = 0; c < headers.length; c++) {
      const h = headers[c].toString().toLowerCase();
      if (kw.some(k => h.includes(k))) return c + 1;
    }
  }

  // 2. Scan up to 10 data rows for Drive URLs
  const scanRows = Math.min(sheet.getLastRow(), 11);
  for (let r = 2; r <= scanRows; r++) {
    const vals = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
    for (let c = 0; c < vals.length; c++) {
      const v = vals[c].toString();
      if (v.includes('drive.google.com') || /\/file\/d\//.test(v)) return c + 1;
    }
  }

  return null;
}

function parseCardData_(text) {
  let t = text.trim()
    .replace(/^```json\s*/i, '').replace(/\s*```$/i, '')
    .replace(/^```\s*/i,     '').replace(/\s*```$/i, '')
    .trim();

  const data = JSON.parse(t); // throws on bad JSON

  if (data.phone !== null && !Array.isArray(data.phone)) data.phone = [data.phone];
  if (data.email !== null && !Array.isArray(data.email)) data.email = [data.email];
  if (!data.phone) data.phone = [];
  if (!data.email) data.email = [];

  return data;
}

// ================================================================
//  DEBUG & DIAGNOSTIC FUNCTIONS
//  Run these from the Apps Script editor to diagnose issues.
// ================================================================

/**
 * Run this to see what the script can detect about your sheet setup.
 * Check the Execution log after running.
 */
function debugFormSheet() {
  Logger.log('=== DEBUG FORM SHEET ===');

  const ss = getSpreadsheet_();
  Logger.log('Spreadsheet: ' + ss.getName() + ' | ID: ' + ss.getId());

  const allSheets = ss.getSheets();
  Logger.log('All tabs (' + allSheets.length + '):');
  allSheets.forEach(s => Logger.log('  - "' + s.getName() + '" | GID: ' + s.getSheetId() + ' | Rows: ' + s.getLastRow()));

  const formSheet = getFormResponseSheet_(ss);
  if (!formSheet) {
    Logger.log('❌ Form Response sheet NOT FOUND. FORM_SHEET_GID = ' + FORM_SHEET_GID);
    Logger.log('   → Update FORM_SHEET_GID in Code.gs to match the correct tab GID above.');
    return;
  }

  Logger.log('\nForm Response sheet: "' + formSheet.getName() + '" | GID: ' + formSheet.getSheetId());
  Logger.log('Rows: ' + formSheet.getLastRow() + ' | Columns: ' + formSheet.getLastColumn());

  if (formSheet.getLastRow() >= 1) {
    const headers = formSheet.getRange(1, 1, 1, formSheet.getLastColumn()).getValues()[0];
    Logger.log('\nColumn headers:');
    headers.forEach((h, i) => Logger.log('  Col ' + (i+1) + ': ' + h));
  }

  const imgCol = IMG_COL > 0 ? IMG_COL : detectImageColumn_(formSheet);
  Logger.log('\nDetected image column: ' + (imgCol ? 'Column ' + imgCol : '❌ NOT DETECTED'));

  if (formSheet.getLastRow() > 1) {
    Logger.log('\nLast form row data:');
    const lastRow = formSheet.getRange(formSheet.getLastRow(), 1, 1, formSheet.getLastColumn()).getValues()[0];
    lastRow.forEach((v, i) => Logger.log('  Col ' + (i+1) + ': ' + v));

    if (imgCol) {
      const url = lastRow[imgCol - 1];
      Logger.log('\nImage URL in last row: ' + url);
      if (url) {
        try {
          const fileId = extractDriveFileId_(url.toString());
          Logger.log('Extracted Drive file ID: ' + fileId);
          const file = DriveApp.getFileById(fileId);
          Logger.log('✅ Drive file accessible: ' + file.getName() + ' (' + file.getMimeType() + ')');
        } catch (e) {
          Logger.log('❌ Drive file error: ' + e.message);
        }
      }
    }
  }

  const triggers = ScriptApp.getProjectTriggers();
  Logger.log('\nInstalled triggers (' + triggers.length + '):');
  triggers.forEach(t => Logger.log('  - ' + t.getHandlerFunction() + ' | type: ' + t.getEventType()));
  if (!triggers.some(t => t.getHandlerFunction() === 'onFormSubmit')) {
    Logger.log('  ⚠️  onFormSubmit trigger NOT found — run setupTrigger() to install it.');
  }

  Logger.log('\n=== END DEBUG ===');
}

/**
 * Manually process the LAST row in the Form Response sheet.
 * Use this to test if everything works without waiting for a new form submission.
 */
function testProcessLastFormRow() {
  Logger.log('=== TEST: Process Last Form Row ===');

  const ss    = getSpreadsheet_();
  const forms = getFormResponseSheet_(ss);

  if (!forms) {
    Logger.log('❌ Form sheet not found. Check FORM_SHEET_GID = ' + FORM_SHEET_GID);
    return;
  }
  if (forms.getLastRow() < 2) {
    Logger.log('❌ No data rows in form sheet yet. Submit the form first.');
    return;
  }

  const imgCol = IMG_COL > 0 ? IMG_COL : detectImageColumn_(forms);
  if (!imgCol) {
    Logger.log('❌ Image column not detected. Run debugFormSheet() for details.');
    return;
  }

  const lastRow = forms.getLastRow();
  const url     = forms.getRange(lastRow, imgCol).getValue().toString().trim();
  Logger.log('Row: ' + lastRow + ' | Image URL: ' + url);

  if (!url) {
    Logger.log('❌ Image URL is empty in last row, column ' + imgCol);
    return;
  }

  try {
    const fileId   = extractDriveFileId_(url);
    Logger.log('Drive file ID: ' + fileId);

    const file = DriveApp.getFileById(fileId);
    Logger.log('File: ' + file.getName() + ' | Type: ' + file.getMimeType() + ' | Size: ' + file.getSize() + ' bytes');

    const blob     = file.getBlob();
    const b64      = Utilities.base64Encode(blob.getBytes());
    const mime     = blob.getContentType() || 'image/jpeg';

    Logger.log('Calling Gemini Vision API…');
    const cardData = callGeminiVision_(b64, mime);
    Logger.log('Extracted: ' + JSON.stringify(cardData, null, 2));

    cardData.source   = 'Google Form (Manual Test)';
    cardData.imageUrl = url;

    const result = writeToOutputSheet_(cardData, null);
    Logger.log('✅ Written to "' + OUT_TAB + '" row ' + result.rowIndex);
    Logger.log('=== DONE ===');
  } catch (err) {
    Logger.log('❌ ERROR: ' + err.message);
    Logger.log('Stack: ' + err.stack);
  }
}

// ================================================================
//  ONE-TIME SETUP FUNCTIONS — Run manually from the editor
// ================================================================

/**
 * Installs the onFormSubmit trigger on the spreadsheet.
 * Run this ONCE. Safe to re-run (removes old trigger first).
 */
function setupTrigger() {
  // Remove any existing onFormSubmit triggers to avoid duplicates
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'onFormSubmit')
    .forEach(t => ScriptApp.deleteTrigger(t));

  const ss = getSpreadsheet_();
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  Logger.log('✅ onFormSubmit trigger installed on: ' + ss.getName());
}

/**
 * Quick test to confirm the Gemini API key is valid.
 */
function testGeminiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) {
    Logger.log('❌ GEMINI_API_KEY not found in Script Properties.');
    Logger.log('   Go to: Apps Script → Project Settings (⚙️) → Script Properties → Add property');
    return;
  }
  Logger.log('Key found (first 8 chars): ' + key.slice(0, 8) + '…');

  const url  = `https://generativelanguage.googleapis.com/v1beta/models?key=${key}`;
  const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  Logger.log(resp.getResponseCode() === 200
    ? '✅ Gemini API key is valid and working!'
    : '❌ Gemini API error: ' + resp.getContentText().slice(0, 300));
}
