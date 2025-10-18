/**
 * 01_access.ts — Sheet access & locking (GAS globals)
 * - getSheetId (optional hook)
 * - resolveSpreadsheetId
 * - getOrCreateSheetByName
 * - getSheetByNameOrCreate
 * - withSheetLock
 */

/** Optional hook: return a specific spreadsheet ID, or null to use Script Properties. */
function getSheetId(): string | null {
  return SpreadsheetApp.getActiveSpreadsheet()?.getId() || null;
}

/** Resolve the spreadsheet ID: hook → Script Properties (SPREADSHEET_ID / DATA_SPREADSHEET_ID). */
function resolveSpreadsheetId(): string {
  try {
    if (typeof getSheetId === "function") {
      const id = getSheetId();
      if (id) return id;
    }
  } catch {}

  const sp = PropertiesService.getScriptProperties();
  const id = sp.getProperty("SPREADSHEET_ID") || sp.getProperty("DATA_SPREADSHEET_ID");
  if (!id) throw new Error("Missing Spreadsheet ID. Provide getSheetId() or set Script Property SPREADSHEET_ID.");
  return id;
}

/**
 * Create or reuse a sheet by name, and ensure the header row equals `headers`.
 * If headers differ, clears the sheet and writes the new header row.
 */
function getOrCreateSheetByName(sheetName: string, headers: string[]) {
  const ss = SpreadsheetApp.openById(resolveSpreadsheetId());
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0] || [];
  const sameHeader = current.join(" | ") === headers.join(" | ");
  if (!sameHeader) {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (typeof formatSheet === "function") {
      try { formatSheet(sheet, headers); } catch {}
    }
  }
  return sheet;
}

/** Return an existing sheet or create it if missing. (Does NOT touch headers.) */
function getSheetByNameOrCreate(
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  sheetName: string
): GoogleAppsScript.Spreadsheet.Sheet {
  const found = ss.getSheetByName(sheetName);
  return found || ss.insertSheet(sheetName);
}

/** Serialize write operations to avoid concurrent collisions. */
function withSheetLock<T>(fn: () => T): T {
  const lock = LockService.getScriptLock();
  lock.waitLock(30_000); // up to 30s
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}