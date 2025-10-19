/// <reference types="google-apps-script" />

/** Resolve spreadsheet, get/create a sheet by name, optionally set headers. */
function getOrCreateSheetByName(
  name: string,
  headers?: string[]
): GoogleAppsScript.Spreadsheet.Sheet {
  const ss = SpreadsheetApp.openById(resolveSpreadsheetId());
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (headers?.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    try { formatSheet(sheet, headers); } catch {}
  }
  return sheet;
}

/** Reset a sheet completely; optionally write headers and format. */
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

  if (headers?.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    try { formatSheet(sheet, headers); } catch {}
  }
  return sheet;
}

/** Ensure headers match on an existing/created sheet. */
function getOrCreateSheetFixHeaders(
  spreadsheetId: string,
  sheetName: string,
  headers: string[]
): { sheet: GoogleAppsScript.Spreadsheet.Sheet; headerChanged: boolean; created: boolean } {
  const ss = SpreadsheetApp.openById(spreadsheetId);
  let sheet = ss.getSheetByName(sheetName);
  const created = !sheet;

  if (!sheet) sheet = ss.insertSheet(sheetName);

  const lastCol = Math.max(sheet.getLastColumn(), headers.length);
  const existing = (lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : []) as string[];
  const currentSlice = existing.slice(0, headers.length);
  const headerChanged =
    currentSlice.length !== headers.length ||
    !currentSlice.every((v, i) => v === headers[i]);

  if (headerChanged) {
    sheet.getRange(1, 1, 1, lastCol).clearContent().clearNote();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    try { formatSheet(sheet, headers); } catch {}
  }

  return { sheet, headerChanged, created };
}
