/// <reference types="google-apps-script" />

/** Optional hook: return active spreadsheet ID or null. */
function getSheetId(): string | null {
  try { return SpreadsheetApp.getActiveSpreadsheet()?.getId() || null; }
  catch { return null; }
}

/** Resolve Spreadsheet ID from hook or Script Properties. */
function resolveSpreadsheetId(): string {
  try {
    if (typeof getSheetId === "function") {
      const id = getSheetId();
      if (id) return id;
    }
  } catch {}

  const sp = PropertiesService.getScriptProperties();
  const id = sp.getProperty("SPREADSHEET_ID") || sp.getProperty("DATA_SPREADSHEET_ID");
  if (!id) {
    throw new Error("Missing Spreadsheet ID. Provide getSheetId() or set Script Property SPREADSHEET_ID.");
  }
  return id;
}

/** Get an existing sheet or create it (does not touch headers). */
function getSheetByNameOrCreate(
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  sheetName: string
): GoogleAppsScript.Spreadsheet.Sheet {
  return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
}

/** Serialize writes to avoid collisions. */
function withSheetLock<T>(fn: () => T): T {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30_000);
  try { return fn(); }
  finally { lock.releaseLock(); }
}
