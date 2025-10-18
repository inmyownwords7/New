/**
 * 02_headers.ts — Header row management (GAS globals)
 * - ensureHeaders
 * - formatSheet
 * - rebuildHeaderMetadataFromNotes
 * - setHeaderCellWithId (stores Notion propId in header NOTE + sheet-level JSON map)
 * - findColumnByPropId
 */

/** JSON key used on the Sheet's developer metadata to store { [colNumber]: decodedPropId } */

/** internal: read a dev-metadata value by key from a sheet */
function getSheetMeta(sheet: Sheet, key: string): string | null {
  const metas = sheet.getDeveloperMetadata() as DevMeta[];
  const hit = metas && metas.find(m => m.getKey() === key);
  return hit ? hit.getValue() : null;
}

/** internal: upsert a dev-metadata key/value on a sheet (remove old first) */
function setSheetMeta(sheet: Sheet, key: string, value: string): void {
  const metas = sheet.getDeveloperMetadata() as DevMeta[];
  metas.forEach(m => { if (m.getKey() === key) m.remove(); });
  sheet.addDeveloperMetadata(key, value);
}

/**
 * Ensure header row equals `headers` (idempotent).
 * Rewrites row 1 if different/missing.
 */
function ensureHeaders(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  headers: string[]
): { changed: boolean } {
  if (!headers?.length) return { changed: false };

  const lastCol = sheet.getLastColumn();
  const existing = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    : [];

  const sameLength = existing.length === headers.length;
  const same = sameLength && existing.every((v, i) => v === headers[i]);

  if (!same) {
    // Clear existing header row if wider than new headers
    if (existing.length && existing.length > headers.length) {
      sheet.getRange(1, 1, 1, existing.length).clearContent();
    }
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return { changed: !same };
}

/**
 * Subtle, idempotent formatting for a tab after headers are set.
 */
function formatSheet(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  headers: string[]
): void {
  if (!sheet || !headers || !headers.length) return;

  // 1) Ensure header row exists (row 1)
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  const current = headerRange.getValues()[0] || [];
  const sameHeader = current.join(" | ") === headers.join(" | ");
  if (!sameHeader) headerRange.setValues([headers]);

  // 2) Freeze header row
  try { sheet.setFrozenRows(1); } catch {}

  // 3) Header styles
  headerRange
    .setFontWeight("bold")
    .setWrap(false)
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle")
    .setBackground("#eef3f8");
  try { headerRange.setFontSize(11); } catch {}

  // 4) Auto-resize columns
  try { sheet.autoResizeColumns(1, headers.length); } catch {}

  // 5) Filter over header + data
  try {
    const range = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), headers.length);
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

  // 6) Alternating row colors for data area only
  try {
    sheet.getBandings().forEach(b => b.remove());
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const dataRange = sheet.getRange(2, 1, lastRow - 1, headers.length);
      const banding = dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
      banding
        .setHeaderRowColor(null)
        .setFirstRowColor("#ffffff")
        .setSecondRowColor("#f7f9fb")
        .setFooterRowColor(null);
    }
  } catch {}

  // 7) Comfortable row height
  try { sheet.setRowHeights(1, Math.max(sheet.getLastRow(), 1), 22); } catch {}
}

/**
 * Rebuild `notionPropId` cell metadata from header NOTES (useful after copies).
 * Only affects cells from `startCol` to the last column.
 */
function rebuildHeaderMetadataFromNotes(sheetName: string, startCol = 2) {
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
    const has = metas && metas.some(m => decodeId(m.getValue()) === pretty && m.getKey() === "notionPropId");
    if (!has) {
      metas.forEach(m => { if (m.getKey() === "notionPropId") m.remove(); });
      cell.addDeveloperMetadata("notionPropId", pretty);
      fixed++;
    }
  }
  Logger.log(`rebuildHeaderMetadataFromNotes: fixed ${fixed} cells`);
  return fixed;
}

/** Header cell writer: label visible, ID in note, and sheet-level JSON map for stability */
function setHeaderCellWithId(
  sheet: Sheet,
  row: number,
  col: number,
  label: string,
  propId: string
) {
  const cell = sheet.getRange(row, col);
  const pretty = decodeId(propId);

  // 1) visible label + note for humans
  cell.setValue(label);
  cell.setNote(pretty);

  // 2) sheet-level metadata map (since Apps Script won't allow cell-range metadata)
  let map: Record<string, string> = {};
  const raw = getSheetMeta(sheet, META_KEY_COLMAP);
  if (raw) {
    try { map = JSON.parse(raw); } catch { map = {}; }
  }
  map[String(col)] = pretty;
  setSheetMeta(sheet, META_KEY_COLMAP, JSON.stringify(map));
}

/** Find column by propId using sheet-level map first, then fallback to header note */
function findColumnByPropId(
  sheet: Sheet,
  propId: string,
  startCol = 2,
  width?: number
): number | null {
  const want = decodeId(propId);
  const lastCol = width ? startCol + width - 1 : sheet.getLastColumn();

  // 1) check sheet-level metadata map
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
    } catch { /* ignore */ }
  }

  // 2) fallback to the header note on row 1 (still visible & robust when copied)
  for (let col = startCol; col <= lastCol; col++) {
    const note = decodeId(sheet.getRange(1, col).getNote() || "");
    if (note === want) return col;
  }
  return null;
}