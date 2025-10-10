function getSheetId(): (string | null) {
    return SpreadsheetApp.getActiveSpreadsheet()?.getId() || null;
}

function formatSheet(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  headers: string[]
): void {
  if (!sheet || !headers || !headers.length) return;

  // 1) Ensure header row exists (row 1)
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  const current = headerRange.getValues()[0] || [];
  const sameHeader = current.join(" | ") === headers.join(" | ");
  if (!sameHeader) {
    headerRange.setValues([headers]);
  }

  // 2) Freeze header row
  try { sheet.setFrozenRows(1); } catch {}

  // 3) Header styles
  headerRange
    .setFontWeight("bold")
    .setWrap(false)
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle")
    .setBackground("#eef3f8"); // subtle header background

  // 4) Auto-resize columns to fit content
  try { sheet.autoResizeColumns(1, headers.length); } catch {}

  // 5) Set a filter on the header (recreate if needed)
  try {
    const range = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), headers.length);
    // Remove existing filter if it doesn't match the range
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
    if (!sheet.getFilter()) {
      range.createFilter();
    }
  } catch {
    // Some sheets (very small/empty) can fail filter creation; ignore
  }

  // 6) Alternating row colors (banded rows) for data area only
  try {
    // Clear existing banding
    const bandings = sheet.getBandings();
    bandings.forEach(b => b.remove());

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const dataRange = sheet.getRange(2, 1, lastRow - 1, headers.length);
      sheet
          // Apply row banding with a light theme, then tweak colors
    const banding = dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
    banding
      .setFirstRowColor("#ffffff")
      .setSecondRowColor("#f7f9fb")
      .setFooterRowColor(null); // no footer band
    // No need to setHeaderRowColor since header isn't in dataRange
  }
  } catch {}

  // 7) Reasonable row height for readability
  try { sheet.setRowHeightsForced(1, Math.max(sheet.getLastRow(), 1), 22); } catch {}

  // 8) Ensure header font size a touch larger
  try { headerRange.setFontSize(11); } catch {}
}

function ensureHeadersByPropId(
  sheetName: string,
  specs: Array<{ label: string; propId: string }>,
  startCol = 2 // B
) {
  const ss = SpreadsheetApp.openById(resolveSpreadsheetId());
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  for (let i = 0; i < specs.length; i++) {
    setHeaderCellWithId(sheet, 1, startCol + i, specs[i].label, specs[i].propId);
  }

  try { sheet.setFrozenRows(1); sheet.autoResizeColumns(startCol, specs.length); } catch {}
  Logger.log(`✅ ensureHeadersByPropId: wrote ${specs.length} headers at row 1 from col ${startCol}`);
  return { count: specs.length, startCol };
}

function extractCellValue(prop: any): string {
  if (!prop) return "";
  switch (prop.type) {
    case "title":
      return (prop.title || []).map((t: any) => t?.plain_text || "").join("");
    case "rich_text":
      return (prop.rich_text || []).map((t: any) => t?.plain_text || "").join("");
    case "email":         return prop.email || "";
    case "phone_number":  return prop.phone_number || "";
    case "url":           return prop.url || "";
    case "date":          return prop.date?.start || "";
    case "status":        return prop.status?.name || "";
    case "select":        return prop.select?.name || "";
    case "multi_select":  return (prop.multi_select || []).map((o: any) => o?.name || "").join(", ");
    case "people":        return (prop.people || []).map((p: any) => p?.name || p?.person?.email || "").join(", ");
    case "number":        return (prop.number ?? "").toString();
    case "checkbox":      return prop.checkbox ? "TRUE" : "FALSE";
    default:              return "";
  }
}



/** Sheets helpers */
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
      // remove old key, add back
      metas.forEach(m => { if (m.getKey() === "notionPropId") m.remove(); });
      cell.addDeveloperMetadata("notionPropId", pretty);
      fixed++;
    }
  }
  Logger.log(`rebuildHeaderMetadataFromNotes: fixed ${fixed} cells`);
  return fixed;
}

function writeContiguousRow(sheetName: string, row: number, startCol: number, values: string[]) {
  const ss = SpreadsheetApp.openById(resolveSpreadsheetId());
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  if (!values.length) return;
  sheet.getRange(row, startCol, 1, values.length).setValues([values]);
}

function writePropIdsToSheet(obj: { properties?: any }, sheetName = "Notion Property IDs") {
  const headers = ["Property Name", "Property ID (raw)", "Property ID (pretty)", "Type"];
  const sheet = getOrCreateSheetByName(sheetName, headers);

  // Clear rows below header
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();

  // Build & de-dupe
  const seen = new Set<string>();
  const rows: any[] = [];
  for (const row of notionPropIdsToRows(obj)) {
    const raw = row[1] as string;
    if (raw && !seen.has(raw)) {
      seen.add(raw);
      rows.push(row);
    }
  }

  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  try { sheet.autoResizeColumns(1, headers.length); } catch {}
  if (typeof formatSheet === "function") {
    try { formatSheet(sheet, headers); } catch {}
  }
  return rows.length;
}

/** Convenience runners */
function writePagePropIdsToSheet(pageIdOrUrl: string, sheetName = "Notion Page Property IDs") {
  const page = notionGetPage(pageIdOrUrl);
  const n = writePropIdsToSheet(page, sheetName);
  Logger.log(`Wrote ${n} unique properties to "${sheetName}".`);
  return n;
}

function writeDataSourcePropIdsToSheet(dsIdOrUrl: string, sheetName = "Notion Data Source Property IDs") {
  const ds = notionGetDataSource(dsIdOrUrl);
  const n = writePropIdsToSheet(ds, sheetName);
  Logger.log(`Wrote ${n} unique properties to "${sheetName}".`);
  return n;
}


/** Build ID→name map from a page/database object */
function buildIdToNameMap(obj: { properties?: any }): Map<string, string> {
  const m = new Map<string, string>();
  const props = obj?.properties || {};
  for (const [name, prop] of Object.entries<any>(props)) {
    const raw = String(prop?.id || "");
    m.set(raw, name);
    m.set(decodeId(raw), name); // allow pretty form too
  }
  return m;
}

/** Cache helpers (Script Properties; tiny + durable) */
function saveIdNameMap(key: string, map: Map<string, string>) {
  const arr = Array.from(map.entries());
  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(arr));
}
function loadIdNameMap(key: string): Map<string, string> {
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) return new Map();
  try { return new Map(JSON.parse(raw) as [string, string][]); } catch { return new Map(); }
}

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