// Client side
// Reverts any cell edit on non-Sign-In sheets.
// Sign-In is exempt — user must be able to type credentials.
// Note: onEdit fires for manual edits only, not button-triggered scripts.
//       Button blocking is handled by assertSystemAccess() inside the library.{
  /*
  function onEdit(e) {
  if (isInternalExecution_()) return;
  // your existing logic below
  if (e.range.getSheet().getName() === "Sign-in") return;
}*/

function validateLicense() { LicenseLib.validateLicense(); }
function lockSystem()      { LicenseLib.lockSystem(); }

// ============================================================
// EXECUTION CONTROL LAYER (CRITICAL FIX)
// ============================================================

// Prevent onEdit recursion
function isInternalExecution_() {
  return PropertiesService.getScriptProperties().getProperty("INTERNAL_RUN") === "true";
}

function setInternalExecution_(val) {
  PropertiesService.getScriptProperties().setProperty("INTERNAL_RUN", val ? "true" : "false");
}

// Cache Spreadsheet instance (huge performance win)
let _ss = null;
function getSS_() {
  if (!_ss) _ss = SpreadsheetApp.getActiveSpreadsheet();
  return _ss;
}

// Cache Spreadsheet ID
let _sheetId = null;
function getSheetId_() {
  if (!_sheetId) _sheetId = getSS_().getId();
  return _sheetId;
}

// litter ("Litter Management Form")
function searchLitter() {
  setInternalExecution_(true);
  try {
  LicenseLib.searchLitter();
 } finally {
    setInternalExecution_(false);
  }
}

function saveLitterData() {
  setInternalExecution_(true);
  try {
  LicenseLib.saveLitterData();
 } finally {
    setInternalExecution_(false);
  }
}

function modifyLitter() {
  setInternalExecution_(true);
  try {
  LicenseLib.modifyLitter();
 } finally {
    setInternalExecution_(false);
  }
}

function clearLitterManagementForm() {
  setInternalExecution_(true);
  try {
  LicenseLib.clearLitterManagementForm();
 } finally {
    setInternalExecution_(false);
  }
}

function deleteLitterBasedOnID() {
  setInternalExecution_(true);
  try {
  LicenseLib.deleteLitterBasedOnID();
 } finally {
    setInternalExecution_(false);
  }
}

// sows (Sow Management Form")
function searchSows() {
  setInternalExecution_(true);
  try {
    LicenseLib.searchSows();
  } finally {
    setInternalExecution_(false);
  }
}

function saveSow() {
  setInternalExecution_(true);
  try {
  LicenseLib.saveSow();
 } finally {
    setInternalExecution_(false);
  }
}

function modifySow() {
  setInternalExecution_(true);
  try {
  LicenseLib.modifySow();
 } finally {
    setInternalExecution_(false);
  }
}

function clearSowManagementForm() {
  setInternalExecution_(true);
  try {
  LicenseLib.clearSowManagementForm();
 } finally {
    setInternalExecution_(false);
  }
}

function deleteSowManagement() {
  setInternalExecution_(true);
  try {
  LicenseLib.deleteSowManagement();
 } finally {
    setInternalExecution_(false);
  }
}

function transferDataUpdate() {
  setInternalExecution_(true);
  try {
  LicenseLib.transferDataUpdate();
 } finally {
    setInternalExecution_(false);
  }
}

// pdf reporting
function generateSingleSowReport(){
  setInternalExecution_(true);
  try {
   LicenseLib.generateSingleSowReport();
 } finally {
    setInternalExecution_(false);
  }
}

function generateReporting() {
  setInternalExecution_(true);
  try {
  LicenseLib.generateReporting();
 } finally {
    setInternalExecution_(false);
  }
}

function generateMonthlyReporting() {
  setInternalExecution_(true);
  try {
  LicenseLib.generateMonthlyReporting();
 } finally {
    setInternalExecution_(false);
  }
}
function updateKPIs(){
   setInternalExecution_(true);
  try {
 LicenseLib.updateKPIs();
  } finally {
    setInternalExecution_(false);
  }
}

function clearBenchmark() {
   setInternalExecution_(true);
  try {
  LicenseLib.clearBenchmark();
 } finally {
    setInternalExecution_(false);
  }
}

// CODE LIBRARY (NOT cleints side!)
// ============================================================
// LICENSE & SECURITY SYSTEM — LIBRARY v3
// Pig Herd Management System
// ============================================================
//
// ARCHITECTURE:
// Every protected function calls assertSystemAccess() which hits
// Firebase LIVE on every call. There is no cached license state.
// A revoked license is caught the moment the next button is pressed.
//
// UserProperties stores ONLY the email + licenseKey so Firebase
// can be queried. It never stores "is valid = true/false" because
// that is what Firebase is for.
//
// onEdit in the user-side script only reverts edits on non-Sign-In
// sheets. Sign-In is always fully editable so the user can type
// their credentials.
//
// CACHING:
// To prevent Firebase bandwidth quota errors on the Spark plan,
// a successful license check is cached in UserProperties for
// CACHE_DURATION_MS (1 hour). The cache stores only the result
// of the last valid check — it never stores "valid=true" directly.
// Firebase is always queried on: first use, cache expiry, and
// any time validateLicense() or lockSystem() is called.
//
// ============================================================

const FIREBASE_PROJECT_ID = "pig-herd-management-admin";
const FIREBASE_API_KEY    = "AIzaSyCjtlbmjEF8gLp9HhMura1QnodKIGHgU1k";
const FIREBASE_BASE_URL   = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents`;

const SIGNIN_SHEET_NAME = "Sign-in";
const EMAIL_CELL        = "D5";
const LICENSE_KEY_CELL  = "D8";
const EXPIRY_DATE_CELL  = "D11";
const STATUS_CELL       = "D14";

// Property keys — store credentials only, never license state
const PROP_EMAIL       = "LICENSED_EMAIL";
const PROP_LICENSE_KEY = "LICENSED_KEY";

// Cache keys and duration
const PROP_CACHE_RESULT    = "LICENSE_CACHE_RESULT";   // JSON of last valid result
const PROP_CACHE_TIMESTAMP = "LICENSE_CACHE_TIMESTAMP"; // ms timestamp of last Firebase hit
const CACHE_DURATION_MS    = 60 * 60 * 1000;           // 1 hour

function getProps() {
  return PropertiesService.getUserProperties();
}

// ── CACHE HELPERS ─────────────────────────────────────────────

// Returns a cached result if it is still fresh, otherwise null.
function getCachedResult() {
  const props     = getProps();
  const timestamp = parseInt(props.getProperty(PROP_CACHE_TIMESTAMP) || "0", 10);
  const cached    = props.getProperty(PROP_CACHE_RESULT);

  if (!cached || !timestamp) return null;

  const age = Date.now() - timestamp;
  if (age > CACHE_DURATION_MS) return null; // cache expired

  try {
    return JSON.parse(cached);
  } catch (e) {
    return null;
  }
}

// Stores a successful Firebase result in the cache.
function setCachedResult(result) {
  const props = getProps();
  props.setProperty(PROP_CACHE_RESULT,    JSON.stringify(result));
  props.setProperty(PROP_CACHE_TIMESTAMP, String(Date.now()));
}

// Clears the cache (called on lock/validate so next check always hits Firebase).
function clearCache() {
  const props = getProps();
  props.deleteProperty(PROP_CACHE_RESULT);
  props.deleteProperty(PROP_CACHE_TIMESTAMP);
}

// ── LIVE FIREBASE CHECK ───────────────────────────────────────
// Called by assertSystemAccess() on every button press.
// Returns { valid, benchmarkEnabled, reason, expirationDate }
// Uses a 1-hour cache to stay within Firebase Spark plan quotas.
function checkLicenseNow() {
  const props      = getProps();
  const email      = props.getProperty(PROP_EMAIL)       || "";
  const licenseKey = props.getProperty(PROP_LICENSE_KEY) || "";

  if (!email || !licenseKey) {
    return { valid: false, reason: "No credentials found. Please validate your license first." };
  }

  // ── Check cache first ────────────────────────────────────────
  const cached = getCachedResult();
  if (cached) {
    Logger.log("License check served from cache (age < 1 hour).");
    return cached;
  }

  // ── Cache miss — query Firebase ──────────────────────────────
  Logger.log("Cache miss or expired — querying Firebase.");
  const result = queryFirebase(email, licenseKey, SpreadsheetApp.getActiveSpreadsheet().getId());

  // Only cache valid results — invalid results (revoked, expired) always re-check
  if (result.valid) setCachedResult(result);

  return result;
}

// Raw Firebase query — returns { valid, benchmarkEnabled, reason, expirationDate }
function queryFirebase(email, licenseKey, sheetId) {
  try {
    const query = {
      structuredQuery: {
        from: [{ collectionId: "licenses" }],
        where: {
          compositeFilter: {
            op: "AND",
            filters: [
              { fieldFilter: { field: { fieldPath: "email"      }, op: "EQUAL", value: { stringValue: email      } } },
              { fieldFilter: { field: { fieldPath: "licenseKey" }, op: "EQUAL", value: { stringValue: licenseKey } } }
            ]
          }
        },
        limit: 1
      }
    };

    const response = UrlFetchApp.fetch(
      `${FIREBASE_BASE_URL}:runQuery?key=${FIREBASE_API_KEY}`,
      { method: "post", contentType: "application/json", payload: JSON.stringify(query), muteHttpExceptions: true }
    );

    const data = JSON.parse(response.getContentText());
    Logger.log("Firebase response: " + JSON.stringify(data));

    if (!data || !data[0] || !data[0].document) {
      return { valid: false, reason: "No matching license found." };
    }

    const doc    = data[0].document.fields;
    const status = doc.status ? doc.status.stringValue.trim().toLowerCase() : "";
    Logger.log("Status: '" + status + "'");

    if (status !== "active") {
      return { valid: false, reason: "Your license has been revoked or is inactive. Please contact info@training4farmers.com" };
    }

    const expirationDate = doc.expirationDate ? doc.expirationDate.stringValue.trim() : "";
    Logger.log("Expiry: '" + expirationDate + "'");

    if (isExpired(expirationDate)) {
      return { valid: false, reason: "Your license expired on " + expirationDate + ". Please contact info@training4farmers.com" };
    }

    // Check spreadsheet binding — prevent one license being used on multiple sheets
    const storedSheetId = doc.url ? extractSheetIdFromUrl(doc.url.stringValue) : null;
    Logger.log("Stored sheet: " + storedSheetId + " | Current: " + sheetId);

    if (storedSheetId && storedSheetId !== sheetId) {
      return { valid: false, reason: "This license is registered to a different spreadsheet. Please contact info@training4farmers.com" };
    }

    // First use — bind this spreadsheet to the license in Firebase
    if (!storedSheetId) bindSheetIdToLicense(data[0].document.name, sheetId);

    const benchmarkEnabled = doc.benchmark ? doc.benchmark.stringValue.trim().toLowerCase() === "yes" : false;
    Logger.log("Benchmark: " + benchmarkEnabled);

    return { valid: true, benchmarkEnabled, expirationDate };

  } catch (e) {
    Logger.log("queryFirebase error: " + e.message);
    return { valid: false, reason: "Network error: " + e.message };
  }
}

// ── ACCESS GUARDS ─────────────────────────────────────────────
// Every button-triggered function in the library starts with one of these.
// They use the 1-hour cache — only the first call per hour hits Firebase.

function assertSystemAccess() {
  const result = checkLicenseNow();

  if (!result.valid) {
    SpreadsheetApp.getUi().alert(
      "Access Denied",
      result.reason + "\n\nContact info@training4farmers.com",
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    // 🔥 Only lock if not already locked
    if (!isSystemLocked_()) {
      lockAllSheets();
    }

    throw new Error("ACCESS_DENIED");
  }

  // 🔥 DO NOT unlock every time
  // unlocking is expensive and unnecessary
  //unlockAllSheets();
}

function assertBenchmarkAccess() {
  const result = checkLicenseNow();

  if (!result.valid) {
    SpreadsheetApp.getUi().alert(
      "Access Denied",
      result.reason + "\n\nContact info@training4farmers.com",
      SpreadsheetApp.getUi().ButtonSet.OK
    );

    if (!isSystemLocked_()) {
      lockAllSheets();
    }

    throw new Error("ACCESS_DENIED");
  }

  if (!result.benchmarkEnabled) {
    SpreadsheetApp.getUi().alert(
      "Premium Feature",
      "Benchmarking is not included in your current plan.\n\nTo upgrade contact info@training4farmers.com",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw new Error("BENCHMARK_NOT_ENABLED");
  }

  // ❌ DO NOT unlock here
}

// ── VALIDATE LICENSE (user-facing) ────────────────────────────
// Always hits Firebase directly — bypasses cache intentionally.
// Saves credentials and refreshes the cache on success.
function validateLicense() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const signInSheet = ss.getSheetByName(SIGNIN_SHEET_NAME);
  const ui          = SpreadsheetApp.getUi();

  if (!signInSheet) {
    ui.alert("Error", "Sign-In sheet not found.", ui.ButtonSet.OK);
    return;
  }

  const email      = signInSheet.getRange(EMAIL_CELL).getValue().toString().trim().toLowerCase();
  const licenseKey = signInSheet.getRange(LICENSE_KEY_CELL).getValue().toString().trim();

  if (!email || !licenseKey) {
    getProps().deleteAllProperties();
    clearCache();
    lockAllSheets();
    writeStatus("Locked");
    ui.alert(
      "Missing Credentials",
      "Please enter your email in 'Email' (D5) and your license key in 'License Key' (D8) before validating.",
      ui.ButtonSet.OK
    );
    return;
  }

  // Always query Firebase directly on explicit validate — ignore cache
  const result = queryFirebase(email, licenseKey, ss.getId());

  if (result.valid) {
    const props = getProps();
    props.setProperty(PROP_EMAIL,       email);
    props.setProperty(PROP_LICENSE_KEY, licenseKey);

    // Refresh cache with the fresh result
    setCachedResult(result);

    signInSheet.getRange(EXPIRY_DATE_CELL).setValue(result.expirationDate);
    writeStatus("Validated");
    unlockAllSheets();

    ui.alert(
      "License Validated",
      "Your license has been validated successfully.\nExpiry: " + result.expirationDate +
      (result.benchmarkEnabled ? "\nBenchmarking: Enabled" : "\nBenchmarking: Not included in your plan."),
      ui.ButtonSet.OK
    );

  } else {
    getProps().deleteAllProperties();
    clearCache();
    writeStatus("Invalid");
    lockAllSheets();

    ui.alert(
      "Validation Failed",
      result.reason,
      ui.ButtonSet.OK
    );
  }
}

// ── ON OPEN ───────────────────────────────────────────────────
// Uses the cache — avoids a Firebase hit every time the sheet opens.
function onOpen() {
  const cached = getCachedResult();

  if (cached && cached.valid) {
    writeStatus("Validated — Expiry: " + cached.expirationDate);
    return;
  }

  const result = checkLicenseNow();

  if (result.valid) {
    writeStatus("Validated — Expiry: " + result.expirationDate);
    unlockAllSheets();
  } else {
    writeStatus("Locked");
    lockAllSheets();
    showSignInPrompt();
  }
}

// ── LOCK SYSTEM (manual) ──────────────────────────────────────
// Clears stored credentials, cache, and locks all sheets.
function lockSystem() {
  getProps().deleteAllProperties();
  clearCache();
  clearBenchmarkData();
  lockAllSheets();
  writeStatus("Locked");
}

// ── SHEET PROTECTION ─────────────────────────────────────────
function lockAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const me = Session.getEffectiveUser().getEmail();

  ss.getSheets().forEach(sheet => {
    if (sheet.getName() === SIGNIN_SHEET_NAME) return;

    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    const protection = sheet.protect().setDescription("SYSTEM_LOCK");
    protection.addEditor(me);
    protection.getEditors().forEach(e => {
      if (e.getEmail() !== me) protection.removeEditor(e);
    });
    protection.setWarningOnly(false);
  });

  const signInSheet = ss.getSheetByName(SIGNIN_SHEET_NAME);
  if (signInSheet) {
    signInSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  }

  PropertiesService.getScriptProperties().setProperty("SYSTEM_LOCKED", "true");
}

function unlockAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ss.getSheets().forEach(sheet => {
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => {
      if (p.getDescription() === "SYSTEM_LOCK") p.remove();
    });
  });

  PropertiesService.getScriptProperties().setProperty("SYSTEM_LOCKED", "false");
}

// ── BIND SHEET TO LICENSE ─────────────────────────────────────
function bindSheetIdToLicense(documentName, sheetId) {
  try {
    UrlFetchApp.fetch(
      `https://firestore.googleapis.com/v1/${documentName}?updateMask.fieldPaths=url&key=${FIREBASE_API_KEY}`,
      {
        method: "patch",
        contentType: "application/json",
        payload: JSON.stringify({ fields: { url: { stringValue: `https://docs.google.com/spreadsheets/d/${sheetId}` } } }),
        muteHttpExceptions: true
      }
    );
  } catch (e) {
    Logger.log("bindSheetIdToLicense error: " + e.message);
  }
}

// ── INTERNAL HELPERS ──────────────────────────────────────────
function isSystemLocked_() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty("SYSTEM_LOCKED") === "true";
} 

function clearBenchmarkData() {
  const sh = SpreadsheetApp.getActive().getSheetByName("Benchmark");
  if (!sh) return;

  // Only clear dynamic/output columns
  sh.getRangeList([
    "C6:C14", // Your Current Value
    "E6:E14", // Status
    "F6:F14", // Recommendation
    "G6:G14", // Best Sows
    "H6:H14"  // Worst Sows
  ]).clearContent();
}

function writeStatus(message) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SIGNIN_SHEET_NAME);
  if (sh) sh.getRange(STATUS_CELL).setValue(message);
}

function extractSheetIdFromUrl(url) {
  if (!url) return null;
  const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

// Handles dd/mm/yyyy, yyyy/mm/dd, yyyy-mm-dd and JS-parseable strings
function isExpired(dateStr) {
  if (!dateStr) return true;

  const dmyMatch = dateStr.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (dmyMatch) {
    return new Date() > new Date(parseInt(dmyMatch[3]), parseInt(dmyMatch[2]) - 1, parseInt(dmyMatch[1]));
  }

  const ymdMatch = dateStr.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (ymdMatch) {
    return new Date() > new Date(parseInt(ymdMatch[1]), parseInt(ymdMatch[2]) - 1, parseInt(ymdMatch[3]));
  }

  const fallback = new Date(dateStr);
  if (!isNaN(fallback)) return new Date() > fallback;

  Logger.log("isExpired: unrecognised format: '" + dateStr + "'");
  return true;
}

function showSignInPrompt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SIGNIN_SHEET_NAME);
  if (sh) ss.setActiveSheet(sh);
  SpreadsheetApp.getUi().alert(
    "System Locked",
    "Please enter your email (D5) and license key (D8), then use License > Validate License.\n\nFor assistance contact info@training4farmers.com",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
// Setting our variables. 

var strDuplicate = "NO";
var ssdata = SpreadsheetApp.getActiveSpreadsheet();
var ss = SpreadsheetApp.getActiveSpreadsheet();
//let isActive = false;
let _ss = null;

function getSS_() {
  if (!_ss) {
    _ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  return _ss;
}
//----------------------------LITTER------------------------------------

// Search Litter
function searchLitter() {
  assertSystemAccess();
  //getSS_()
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  var litterFormSheet = ss.getSheetByName('Litter Management Form');
  var litterDatabaseSheet = ss.getSheetByName('Litter Management Database');
  var ui = SpreadsheetApp.getUi();

  // Value to search for (from D4)
  var nameToSearch = litterFormSheet.getRange('D4').getValue();

  var lastRow = litterDatabaseSheet.getLastRow();
  var namesColumn = litterDatabaseSheet.getRange('E1:E' + lastRow).getValues();
  var foundRowIndex = -1;

  // Check for matching Litter ID in column E
  for (var i = 0; i < namesColumn.length; i++) {
    if (String(namesColumn[i][0]).trim() === String(nameToSearch).trim()) {
      foundRowIndex = i + 1;
      break;
    }
  }

  if (foundRowIndex !== -1) {
    var rowData = litterDatabaseSheet.getRange(foundRowIndex, 2, 1, litterDatabaseSheet.getLastColumn() - 1).getValues()[0];

    // LEFT COLUMN (D column)
    litterFormSheet.getRange('D7').setValue(rowData[0]);    // B → Sow Number
    litterFormSheet.getRange('D9').setValue(rowData[1]);    // C → Parity
    litterFormSheet.getRange('D11').setValue(rowData[2]);   // D → Boar Number
    litterFormSheet.getRange('D13').setValue(rowData[3]);   // E → Litter ID
    litterFormSheet.getRange('D15').setValue(rowData[4]);   // F → D.O.B
    litterFormSheet.getRange('D17').setValue(rowData[5]);   // G → Born Alive
    litterFormSheet.getRange('D19').setValue(rowData[6]);   // H → Born Dead
    litterFormSheet.getRange('D21').setValue(rowData[7]);   // I → Mummified
    litterFormSheet.getRange('D23').setValue(rowData[8]);   // J → Overlays & Other causes

    // RIGHT COLUMN (G column)
    litterFormSheet.getRange('G7').setValue(rowData[9]);    // K → Total Born
    litterFormSheet.getRange('G9').setValue(rowData[10]);   // L → Total Birth Weight
    litterFormSheet.getRange('G11').setValue(rowData[11]);  // M → Date Weaned
    litterFormSheet.getRange('G13').setValue(rowData[12]);  // N → Number Weaned
    litterFormSheet.getRange('G15').setValue(rowData[13]);  // O → Date Killed
    litterFormSheet.getRange('G17').setValue(rowData[14]);  // P → Number Sent To Abattoir
    litterFormSheet.getRange('G19').setValue(rowData[15]);  // Q → Total Carcass Weight
    litterFormSheet.getRange('G21').setValue(rowData[16]);  // R → Age at Slaughter
    litterFormSheet.getRange('G23').setValue(rowData[17]);  // S → Number Sold Live
    
    // Far right
    litterFormSheet.getRange('J7').setValue(rowData[18]);   // T → Total Weight of Stock Sold Live

    ui.alert('Match Found', 'Matching data has been populated in the Results panel.', ui.ButtonSet.OK);
  } else {
    ui.alert('No matching data found', 'No matching data found in the database.', ui.ButtonSet.OK);
  }
}

// Save Litter
function saveLitterData() {
  assertSystemAccess();
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName('Litter Management Form');
  var targetSheet = ss.getSheetByName('Litter Management Database');

  // ── Collect form values ──────────────────────────────────────
  // Col K (Total Born) and Col N (Number Weaned) are FORMULA columns
  // in the database — we never write to them; the sheet computes them.
  var group1 = [                                          // → DB cols B–J (sheet cols 2–10)
    sourceSheet.getRange('D7').getValue(),   // B → Sow Number
    sourceSheet.getRange('D9').getValue(),   // C → Parity
    sourceSheet.getRange('D11').getValue(),  // D → Boar Number
    sourceSheet.getRange('D13').getValue(),  // E → Litter ID
    sourceSheet.getRange('D15').getValue(),  // F → D.O.B
    sourceSheet.getRange('D17').getValue(),  // G → Born Alive
    sourceSheet.getRange('D19').getValue(),  // H → Born Dead
    sourceSheet.getRange('D21').getValue(),  // I → Mummified
    sourceSheet.getRange('D23').getValue()   // J → Overlays & Other causes
    // ── Col K (Total Born) is a formula — SKIPPED ──
  ];
  var group2 = [                                          // → DB cols L–M (sheet cols 12–13)
    sourceSheet.getRange('G9').getValue(),   // L → Total Birth Weight
    sourceSheet.getRange('G11').getValue()   // M → Date Weaned
    // ── Col N (Number Weaned) is a formula — SKIPPED ──
  ];
  var group3 = [                                          // → DB cols O–T (sheet cols 15–20)
    sourceSheet.getRange('G15').getValue(),  // O → Date Killed
    sourceSheet.getRange('G17').getValue(),  // P → Number Sent To Abattoir
    sourceSheet.getRange('G19').getValue(),  // Q → Total Carcass Weight
    sourceSheet.getRange('G21').getValue(),  // R → Age at Slaughter
    sourceSheet.getRange('G23').getValue(),  // S → Number Sold Live
    sourceSheet.getRange('J7').getValue()    // T → Total Weight of Stock Sold Live
  ];

  var identifier = sourceSheet.getRange('D13').getValue(); // Litter ID

  // ── Duplicate check (col E = Litter ID) ─────────────────────
  var columnEValues = targetSheet.getRange('E1:E' + targetSheet.getLastRow()).getValues().flat();
  if (columnEValues.includes(identifier)) {
    var ui = SpreadsheetApp.getUi();
    ui.alert('Error', 'Duplicate Litter ID found. Data not transferred.', ui.ButtonSet.OK);
    return;
  }

  // ── Write to next available row, skipping formula columns ───
  var nextRow = targetSheet.getLastRow() + 1;
  targetSheet.getRange(nextRow, 2,  1, group1.length).setValues([group1]); // B–J
  targetSheet.getRange(nextRow, 12, 1, group2.length).setValues([group2]); // L–M
  targetSheet.getRange(nextRow, 15, 1, group3.length).setValues([group3]); // O–T

  SpreadsheetApp.getUi().alert('Success', 'Data saved successfully.', SpreadsheetApp.getUi().ButtonSet.OK);
}
