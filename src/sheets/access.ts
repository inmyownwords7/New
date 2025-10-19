/// <reference types="google-apps-script" />

/**
 * 01_access.ts — Sheet access & locking (GAS globals)
 * ---------------------------------------------------
 * Exposes:
 *  - getSheetId (optional hook)
 *  - resolveSpreadsheetId
 *  - getSheetByNameOrCreate
 *  - getOrCreateSheetByName (alias used by other modules)
 *  - withSheetLock
 *  - getOrCreateSheetAndWipeAll
 *  - getOrCreateSheetFixHeaders
 */

/* -------------------------------------------------------------------------- */
/*                             Spreadsheet resolve                             */
/* -------------------------------------------------------------------------- */

/**
 * Optional hook: return a specific spreadsheet ID, or null to use Script Properties.
 * Override this if you want to target a fixed spreadsheet.
 */
function getSheetId(): string | null {
  try { return SpreadsheetApp.getActiveSpreadsheet()?.getId() || null; }
  catch { return null; }
}

/**
 * Resolve the spreadsheet ID: hook → Script Properties (SPREADSHEET_ID / DATA_SPREADSHEET_ID).
 * @throws if an ID cannot be resolved.
 */
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

/* -------------------------------------------------------------------------- */
/*                              Sheet creation/get                             */
/* -------------------------------------------------------------------------- */

/**
 * Return an existing sheet or create it if missing. (Does NOT touch headers.)
 *
 * @param ss        Target spreadsheet.
 * @param sheetName Tab name to get or create.
 */
function getSheetByNameOrCreate(
    ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
    sheetName: string
): GoogleAppsScript.Spreadsheet.Sheet {
  return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
}

/**
 * Alias used by other modules: get or create a sheet by name in the *resolved* spreadsheet.
 * Optionally writes headers if provided (row 1).
 *
 * @param name    Tab name.
 * @param headers Optional header labels to set on row 1.
 */
function getOrCreateSheetByName(
    name: string,
    headers?: string[]
): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.openById(resolveSpreadsheetId());
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (headers && headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    try { if (typeof formatSheet === "function") formatSheet(sheet, headers); } catch {}
  }
  return sheet;
}

/* -------------------------------------------------------------------------- */
/*                                   Locking                                   */
/* -------------------------------------------------------------------------- */

/**
 * Serialize write operations to avoid concurrent collisions.
 * Uses a document-scoped lock so unrelated scripts are not blocked.
 *
 * @param fn Work to run under lock.
 */
function withSheetLock<T>(fn: () => T): T {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30_000); // up to 30s
  try { return fn(); }
  finally { lock.releaseLock(); }
}

/* -------------------------------------------------------------------------- */
/*                               Higher-level ops                              */
/* -------------------------------------------------------------------------- */

/**
 * Get (or create) a sheet by name and completely clear its contents.
 * Useful when you want a fresh export or full resync.
 *
 * - Clears values, notes, and formats
 * - Removes filters and banding if present
 * - Optionally writes headers (row 1) and formats
 *
 * @param spreadsheetId Spreadsheet ID to open.
 * @param sheetName     Tab name to reset.
 * @param headers       Optional header labels to set on row 1.
 */
function getOrCreateSheetAndWipeAll(
    spreadsheetId: string,
    sheetName: string,
    headers?: string[]
): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sheet = ss.getSheetByName(sheetName);

  if (sheet) {
    sheet.clearContents();
    sheet.clearNotes();
    sheet.clearFormats();
    try { const f = sheet.getFilter(); if (f) f.remove(); } catch {}
    try { sheet.getBandings().forEach(b => { try { b.remove(); } catch {} }); } catch {}
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  if (headers && headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    try { if (typeof formatSheet === "function") formatSheet(sheet, headers); } catch {}
  }

  return sheet;
}

/**
 * Get or create a sheet and ensure its header row matches `headers`.
 * Only modifies row 1 — data rows remain untouched.
 *
 * @param spreadsheetId Spreadsheet ID to open.
 * @param sheetName     Tab name to ensure.
 * @param headers       Header labels desired on row 1.
 * @returns             { sheet, headerChanged, created }
 */
function getOrCreateSheetFixHeaders(
    spreadsheetId: string,
    sheetName: string,
    headers: string[]
): { sheet: GoogleAppsScript.Spreadsheet.Sheet; headerChanged: boolean; created: boolean } {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sheet = ss.getSheetByName(sheetName);
  const created = !sheet;

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  const lastCol = Math.max(sheet.getLastColumn(), headers.length);
  const existing = (lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : []) as string[];
  const currentSlice = existing.slice(0, headers.length);
  const headerChanged =
      currentSlice.length !== headers.length ||
      !currentSlice.every((v, i) => v === headers[i]);

  if (headerChanged) {
    // Clear only row 1 (headers) and re-write
    sheet.getRange(1, 1, 1, lastCol).clearContent().clearNote();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    try { if (typeof formatSheet === "function") formatSheet(sheet, headers); } catch {}
  }

  return { sheet, headerChanged, created };
}