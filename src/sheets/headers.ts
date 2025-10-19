/// <reference types="google-apps-script" />

/**
 * 02_headers.ts — Header row management (GAS globals)
 * ---------------------------------------------------
 * Exposes:
 *  - ensureHeaders
 *  - formatSheet
 *  - rebuildHeaderMetadataFromNotes
 *  - setHeaderCellWithId (stores Notion propId in header NOTE + sheet-level JSON map)
 *  - findColumnByPropId
 *  - ensureAliasHeadersExact* helpers
 *
 * Depends on globals defined elsewhere:
 *  - const META_KEY_COLMAP: string
 *  - function getSheetMeta(sheet: Sheet, key: string): string|null
 *  - function setSheetMeta(sheet: Sheet, key: string, value: string): void
 *  - function decodeId(id: unknown): string
 *  - function resolveSpreadsheetId(): string
 *  - types: Sheet, DevMeta
 */

/* -------------------------------------------------------------------------- */
/*                               Header writers                                */
/* -------------------------------------------------------------------------- */

/**
 * Ensure header row (row 1) equals `headers` (idempotent, labels only).
 * Does not touch data rows. Returns whether row 1 was changed.
 *
 * @param sheet   Target sheet.
 * @param headers Desired header labels for row 1.
 * @returns       { changed: boolean }
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

    // Cosmetic polish (safe to fail)
    try { formatSheet(sheet, headers); } catch {}
  }

  return { changed: !same };
}

/**
 * Subtle, idempotent formatting for a tab after headers are set.
 * - Bold, left-aligned headers w/ light background
 * - Freeze header row
 * - Auto-resize columns to fit header text
 * - Optional filter + alternating banding (data area only)
 *
 * @param sheet   Target sheet.
 * @param headers Header labels that now exist on row 1.
 */
function formatSheet(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    headers: string[]
): void {
  if (!sheet || !headers?.length) return;

  const headerRange = sheet.getRange(1, 1, 1, headers.length);

  try {
    // Header styles
    headerRange
        .setFontWeight("bold")
        .setWrap(false)
        .setHorizontalAlignment("left")
        .setVerticalAlignment("middle")
        .setBackground("#eef3f8");
    try { headerRange.setFontSize(11); } catch {}

    // Freeze header row
    try { sheet.setFrozenRows(1); } catch {}

    // Auto-resize columns used by headers
    try { sheet.autoResizeColumns(1, headers.length); } catch {}

    // Add/refresh filter spanning header + existing data
    try {
      const range = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), Math.max(sheet.getLastColumn(), headers.length));
      const existing = sheet.getFilter();
      if (existing) {
        const fr = existing.getRange();
        const same =
            fr.getRow() === range.getRow() &&
            fr.getColumn() === range.getColumn() &&
            fr.getNumRows() === range.getNumRows() &&
            fr.getNumColumns() === range.getNumColumns();
        if (!same) existing.remove();
      }
      if (!sheet.getFilter()) range.createFilter();
    } catch {}

    // Alternating banding on data rows (2..lastRow), not header
    try {
      sheet.getBandings().forEach(b => { try { b.remove(); } catch {} });
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const dataRange = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), headers.length));
        const banding = dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
        banding
            .setHeaderRowColor(null)
            .setFirstRowColor("#ffffff")
            .setSecondRowColor("#f7f9fb")
            .setFooterRowColor(null);
      }
    } catch {}

    // Comfortable row height
    try { sheet.setRowHeights(1, Math.max(sheet.getLastRow(), 1), 22); } catch {}
  } catch {
    // Cosmetic only—ignore errors
  }
}

/* -------------------------------------------------------------------------- */
/*                        Metadata rebuild & header IDs                        */
/* -------------------------------------------------------------------------- */

/**
 * Rebuild cell-level developer metadata on header cells from the header NOTE,
 * useful after copying/sync operations. Scans from startCol to last column.
 *
 * - Reads header NOTE (decoded propId expected)
 * - Ensures Range-level DeveloperMetadata "notionPropId" matches the note
 *
 * @param sheetName Tab name to process.
 * @param startCol  1-based column to start scanning (default: 2).
 * @returns         Number of header cells fixed.
 */
function rebuildHeaderMetadataFromNotes(sheetName: string, startCol = 2): number {
  const ss = SpreadsheetApp.openById(resolveSpreadsheetId());
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;

  const lastCol = sheet.getLastColumn();
  let fixed = 0;

  for (let col = startCol; col <= lastCol; col++) {
    const cell = sheet.getRange(1, col);
    const note = cell.getNote();
    if (!note) continue;

    const pretty = decodeId(note);
    const metas = cell.getDeveloperMetadata();
    const has = metas && metas.some(m => m.getKey() === "notionPropId" && decodeId(m.getValue()) === pretty);
    if (!has) {
      // Remove stale metadata entries of the same key
      metas.forEach(m => { if (m.getKey() === "notionPropId") m.remove(); });
      cell.addDeveloperMetadata("notionPropId", pretty);
      fixed++;
    }
  }
  Logger.log(`rebuildHeaderMetadataFromNotes: fixed ${fixed} cells`);
  return fixed;
}

/**
 * Write a header cell: set visible label, store decoded propId in NOTE,
 * and update the sheet-level column→propId JSON map in DeveloperMetadata.
 *
 * @param sheet   Target sheet.
 * @param row     1-based row (usually 1).
 * @param col     1-based column.
 * @param label   Visible column header label.
 * @param propId  Notion property id (raw or percent-encoded).
 */
function setHeaderCellWithId(
    sheet: Sheet,
    row: number,
    col: number,
    label: string,
    propId: string
): void {
  const cell = sheet.getRange(row, col);
  const pretty = decodeId(propId);

  // 1) Visible label + note for humans
  cell.setValue(label);
  cell.setNote(pretty);

  // 2) Sheet-level metadata map: { [col]: decodedPropId }
  let map: Record<string, string> = {};
  const raw = getSheetMeta(sheet, META_KEY_COLMAP);
  if (raw) {
    try { map = JSON.parse(raw) as Record<string, string>; } catch { map = {}; }
  }
  map[String(col)] = pretty;
  setSheetMeta(sheet, META_KEY_COLMAP, JSON.stringify(map));

  // 3) Optional: add Range-level metadata for the cell itself (useful for filters/queries)
  try {
    const metas = cell.getDeveloperMetadata();
    metas.forEach(m => { if (m.getKey() === "notionPropId") m.remove(); });
    cell.addDeveloperMetadata("notionPropId", pretty);
  } catch {}
}

/**
 * Find a column by Notion propId:
 * 1) Use the sheet-level JSON map (META_KEY_COLMAP)
 * 2) Fallback to scanning header NOTE values
 *
 * @param sheet    Target sheet.
 * @param propId   Notion property id (raw/encoded ok).
 * @param startCol 1-based start column (default: 2; leaves A for keys).
 * @param width    Optional width to constrain the scan window (startCol..startCol+width-1).
 * @returns        1-based column index or null if not found.
 */
function findColumnByPropId(
    sheet: Sheet,
    propId: string,
    startCol = 2,
    width?: number
): number | null {
  const want = decodeId(propId);
  const lastCol = width ? startCol + width - 1 : sheet.getLastColumn();

  // 1) Check sheet-level metadata map first
  const raw = getSheetMeta(sheet, META_KEY_COLMAP);
  if (raw) {
    try {
      const map = JSON.parse(raw) as Record<string, string>;
      for (const k in map) {
        const col = Number(k);
        if (col >= startCol && col <= lastCol && decodeId(map[k]) === want) {
          return col;
        }
      }
    } catch { /* ignore parse error, fallback below */ }
  }

  // 2) Fallback: match header NOTE (decoded)
  for (let col = startCol; col <= lastCol; col++) {
    const note = decodeId(sheet.getRange(1, col).getNote() || "");
    if (note === want) return col;
  }

  return null;
}

/* -------------------------------------------------------------------------- */
/*                         Exact alias header management                       */
/* -------------------------------------------------------------------------- */

/**
 * Ensure the header row exactly matches the given label+propId pairs
 * (writes labels, stores decoded propId in header note, and updates sheet-level
 * col→propId map). Idempotent: only rewrites when something differs.
 *
 * @param sheet Target sheet.
 * @param pairs Array of { label, propId } in desired order, starting at column 1.
 * @returns     { changed: boolean } whether labels were rewritten.
 */
function ensureAliasHeadersExact(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    pairs: Array<{ label: string; propId: string }>
): { changed: boolean } {
  if (!sheet || !pairs?.length) return { changed: false };

  const width = pairs.length;
  const existing = sheet.getLastColumn() > 0
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      : [];

  // Compare labels only for the quick equality check
  const desiredLabels = pairs.map(p => p.label);
  const sameLength = existing.length === width;
  const sameLabels = sameLength && desiredLabels.every((lbl, i) => existing[i] === lbl);

  if (!sameLabels) {
    // Clear old header if wider than new
    if (existing.length && existing.length > width) {
      sheet.getRange(1, 1, 1, existing.length).clearContent().clearNote();
    }
    sheet.getRange(1, 1, 1, width).setValues([desiredLabels]);
  }

  // Ensure each header cell has correct note + metadata
  for (let i = 0; i < width; i++) {
    const { label, propId } = pairs[i];
    const cell = sheet.getRange(1, i + 1);
    const note = cell.getNote() || "";
    const decodedNote = decodeId(note);
    const decodedWant = decodeId(propId);

    const needsLabel = cell.getValue() !== label;
    const needsNote = decodedNote !== decodedWant;

    if (needsLabel || needsNote) {
      setHeaderCellWithId(sheet as Sheet, 1, i + 1, label, propId);
    }
  }

  // Cosmetic formatting
  try { formatSheet(sheet, desiredLabels); } catch {}

  return { changed: !sameLabels };
}

/**
 * Convenience overload: ensure exact alias headers from a Map<label, propId>.
 *
 * @param sheet        Target sheet.
 * @param labelToPropId Map of header label -> Notion propId.
 * @returns            { changed: boolean }
 */
function ensureAliasHeadersExactFromMap(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    labelToPropId: Map<string, string>
): { changed: boolean } {
  const pairs = Array.from(labelToPropId.entries()).map(([label, propId]) => ({ label, propId }));
  return ensureAliasHeadersExact(sheet, pairs);
}

/**
 * Convenience overload: ensure exact alias headers from parallel arrays.
 *
 * @param sheet   Target sheet.
 * @param labels  Header labels.
 * @param propIds Corresponding Notion propIds.
 * @returns       { changed: boolean }
 */
function ensureAliasHeadersExactFromArrays(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    labels: string[],
    propIds: string[]
): { changed: boolean } {
  const pairs = labels.map((label, i) => ({ label, propId: propIds[i] || "" }));
  return ensureAliasHeadersExact(sheet, pairs);
}

/**
 * Ensure exact alias headers starting at a given column (1-based).
 * Writes labels, notes, and updates the sheet-level col→propId map for that window.
 *
 * @param sheet    Target sheet.
 * @param pairs    Array of { label, propId } for the segment.
 * @param startCol 1-based starting column (default: 1).
 * @returns        { changed: boolean }
 */
function ensureAliasHeadersExactAt(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    pairs: Array<{ label: string; propId: string }>,
    startCol: number = 1
): { changed: boolean } {
  if (!sheet || !pairs?.length) return { changed: false };
  const width = pairs.length;

  // Current labels in the target window
  const headerRange = sheet.getRange(1, startCol, 1, width);
  const current = (width > 0 ? headerRange.getValues()[0] : []) as string[];

  const desiredLabels = pairs.map(p => p.label);
  const sameLength = current.length === width;
  const sameLabels = sameLength && desiredLabels.every((lbl, i) => current[i] === lbl);

  if (!sameLabels) {
    // Clear the existing header cells in our window (notes too)
    sheet.getRange(1, startCol, 1, width).clearContent().clearNote();
  }

  // Write each header cell + note + metadata
  for (let i = 0; i < width; i++) {
    const { label, propId } = pairs[i];
    setHeaderCellWithId(sheet as Sheet, 1, startCol + i, label, propId);
  }

  // Formatting pass using the full current header width
  try {
    const hdrVals = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), startCol + width - 1)).getValues()[0];
    formatSheet(sheet, hdrVals);
  } catch {}

  return { changed: !sameLabels };
}