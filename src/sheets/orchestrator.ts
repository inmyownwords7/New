/// <reference types="google-apps-script" />

/**
 * sheets/orchestrator.ts
 * ------------------------------------------------
 * Core helpers for working with Google Sheets in GAS (globals).
 * Safe to include alongside your existing files.
 *
 * Exposes:
 *  - getSheetId
 *  - resolveSpreadsheetId
 *  - getSheetByNameOrCreate
 *  - withSheetLock
 *  - ensureHeaders
 *  - formatSheet
 */

/* -------------------------------------------------------------------------- */
/*                             Access & locking                                */
/* -------------------------------------------------------------------------- */

/**
 * Optional hook: use the active spreadsheet if present; else null.
 * You can replace this to route to a specific spreadsheet.
 */
function getSheetId(): string | null {
    try {
        return SpreadsheetApp.getActiveSpreadsheet()?.getId() || null;
    } catch {
        return null;
    }
}

/**
 * Resolve a spreadsheet ID via the hook or Script Properties.
 * Checks SPREADSHEET_ID, then DATA_SPREADSHEET_ID.
 *
 * @throws when no spreadsheet id can be resolved.
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

/**
 * Return an existing sheet or create it if missing. (Does NOT touch headers.)
 *
 * @param ss        Target spreadsheet.
 * @param sheetName Tab name.
 */
function getSheetByNameOrCreate(
    ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
    sheetName: string
): GoogleAppsScript.Spreadsheet.Sheet {
    const found = ss.getSheetByName(sheetName);
    return found || ss.insertSheet(sheetName);
}

/**
 * Serialize write operations to avoid concurrent collisions.
 * Uses a document-scoped lock so unrelated scripts aren't blocked.
 *
 * @param fn Work to run under lock.
 */
function withSheetLock<T>(fn: () => T): T {
    const lock = LockService.getDocumentLock();
    lock.waitLock(30_000);
    try {
        return fn();
    } finally {
        lock.releaseLock();
    }
}

/* -------------------------------------------------------------------------- */
/*                             Headers & formatting                            */
/* -------------------------------------------------------------------------- */

/**
 * Ensure header row (row 1) equals `headers` (idempotent, labels only).
 * Does not clear or alter data rows.
 *
 * @param sheet   Target sheet.
 * @param headers Desired header labels for row 1.
 * @returns       Whether the header row was changed.
 */
function ensureHeaders(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    headers: string[]
): { changed: boolean } {
    if (!headers?.length) return { changed: false };

    const width = headers.length;
    const lastCol = Math.max(sheet.getLastColumn(), width);
    const current = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
    const currentSlice = current.slice(0, width).map(v => String(v ?? ""));

    const same =
        currentSlice.length === width &&
        currentSlice.every((v, i) => v === headers[i]);

    if (!same) {
        // Clear only headers row (avoid touching data below)
        if (lastCol > 0) sheet.getRange(1, 1, 1, lastCol).clearContent().clearNote();
        sheet.getRange(1, 1, 1, width).setValues([headers]);

        // Optional polish
        try { formatSheet(sheet, headers); } catch {}
    }

    return { changed: !same };
}

/**
 * Subtle sheet formatting after headers are set.
 * - Bold, left-aligned headers
 * - Freeze header row
 * - Auto-resize columns to fit header text
 * - Optional alternating banding + filter
 *
 * Safe no-op if anything fails.
 *
 * @param sheet   Target sheet.
 * @param headers Header labels you just set (row 1).
 */
function formatSheet(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    headers: string[]
): void {
    if (!sheet || !headers?.length) return;

    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    try {
        headerRange
            .setFontWeight("bold")
            .setWrap(false)
            .setHorizontalAlignment("left")
            .setVerticalAlignment("middle");

        // Light header background
        headerRange.setBackground("#f3f4f6"); // Tailwind-ish gray-100

        // Freeze header row
        sheet.setFrozenRows(1);

        // Auto-resize current header columns (avoid resizing beyond headers.length)
        try { sheet.autoResizeColumns(1, headers.length); } catch {}

        // Optional: add a basic filter if none exists and there are rows
        try {
            if (!sheet.getFilter() && sheet.getLastRow() >= 1 && sheet.getLastColumn() >= 1) {
                const filterRange = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), Math.max(sheet.getLastColumn(), headers.length));
                filterRange.createFilter();
            }
        } catch {}

        // Optional: alternating banding on data region (not on the header row)
        try {
            // Remove existing bandings to avoid stacking
            const bandings = sheet.getBandings();
            bandings.forEach(b => { try { b.remove(); } catch {} });

            if (sheet.getLastRow() > 1) {
                sheet.getRange(2, 1, sheet.getLastRow() - 1, Math.max(sheet.getLastColumn(), headers.length))
                    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
            }
        } catch {}
    } catch {
        // swallow formatting errors; they're cosmetic
    }
}