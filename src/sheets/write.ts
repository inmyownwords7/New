/*************************************************
 * sheets/03_writes.ts
 * ------------------------------------------------
 * Handles all row-level write operations to Google Sheets:
 * - Writing or appending rows
 * - Clearing data areas
 * - Upserting data by key
 * - Writing Notion property IDs
 * - Flattening Notion properties into plain text
 *************************************************/

/// <reference types="google-apps-script" />

/**
 * Clear all data below the header row (keeps row 1 intact).
 * Optionally, extend this to also clear notes or formats if needed.
 *
 * @param sheet The target Google Sheet.
 */
function clearDataBelowHeader(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol === 0) return;
  sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
}

/**
 * Append rows to a sheet in batches to avoid quota limits and improve speed.
 *
 * @param sheet      Target sheet.
 * @param rows       2D array of rows to append.
 * @param batchSize  Maximum rows per batch (default 500).
 */
function appendRowsBatched(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    rows: any[][],
    batchSize: number = 500
): void {
  if (!rows?.length) return;

  let writeRow = Math.max(sheet.getLastRow() + 1, 2); // preserve header row
  for (let i = 0; i < rows.length; i += batchSize) {
    const chunk = rows.slice(i, i + batchSize);
    const width = chunk[0]?.length || 0;
    if (!width) continue;
    sheet.getRange(writeRow, 1, chunk.length, width).setValues(chunk);
    writeRow += chunk.length;
  }
}

/**
 * Upsert (update or insert) rows into a sheet by matching a key column.
 * - If the key already exists, that row is updated.
 * - If the key is missing, the row is appended.
 *
 * @param sheet     Target sheet.
 * @param keyHeader Header label to use as the unique key.
 * @param headers   List of all column headers.
 * @param rows      Rows to upsert.
 * @returns         Counts of inserted and updated rows.
 */
function upsertRowsByKey(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    keyHeader: string,
    headers: string[],
    rows: any[][]
): { inserted: number; updated: number } {
  let inserted = 0, updated = 0;
  if (!rows?.length) return { inserted, updated };

  const keyIdx0 = headers.indexOf(keyHeader);
  if (keyIdx0 === -1) throw new Error(`Key header "${keyHeader}" not found.`);

  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), headers.length);
  const existing = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

  // Build a map: key → sheet row number
  const keyToRow = new Map<string, number>();
  for (let i = 0; i < existing.length; i++) {
    const k = String(existing[i][keyIdx0] ?? "").trim();
    if (k) keyToRow.set(k, 2 + i);
  }

  // Normalize row width to match headers
  const normalizeRow = (r: any[]) => {
    const out = r.slice(0, headers.length);
    while (out.length < headers.length) out.push("");
    return out;
  };

  const inserts: any[][] = [];
  for (const r of rows) {
    const row = normalizeRow(r);
    const key = String(row[keyIdx0] ?? "").trim();
    if (!key) continue;

    const at = keyToRow.get(key);
    if (at) {
      sheet.getRange(at, 1, 1, headers.length).setValues([row]);
      updated++;
    } else {
      inserts.push(row);
    }
  }

  if (inserts.length) {
    appendRowsBatched(sheet, inserts, 500);
    inserted += inserts.length;
  }

  return { inserted, updated };
}

/**
 * Write contiguous values starting from a specific row and column.
 *
 * @param sheetName Name of the sheet.
 * @param row       Row index to start writing (1-based).
 * @param startCol  Column index to start writing (1-based).
 * @param values    Values to write as a single row.
 */
function writeContiguousRow(
    sheetName: string,
    row: number,
    startCol: number,
    values: string[]
): void {
  const ss = SpreadsheetApp.openById(resolveSpreadsheetId());
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  if (!values.length) return;
  sheet.getRange(row, startCol, 1, values.length).setValues([values]);
}

/**
 * Write a Notion object’s property IDs to a dedicated "Property IDs" sheet.
 *
 * @param obj       A Notion page or data source object.
 * @param sheetName Optional name of the sheet (default "Notion Property IDs").
 * @returns         Number of unique properties written.
 */
function writePropIdsToSheet(
    obj: { properties?: any },
    sheetName = "Notion Property IDs"
): number {
  const headers = ["Property Name", "Property ID (raw)", "Property ID (pretty)", "Type"];
  const sheet = getOrCreateSheetByName(sheetName, headers);

  // Clear existing property rows
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();

  const seen = new Set<string>();
  const rows: any[] = [];

  for (const [name, prop] of Object.entries<any>(obj?.properties || {})) {
    const raw = String(prop?.id || "");
    if (!raw || seen.has(raw)) continue;
    seen.add(raw);
    rows.push([
      name,
      raw,
      decodeId(raw),
      prop?.type || ""
    ]);
  }

  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  try { sheet.autoResizeColumns(1, headers.length); } catch {}

  if (typeof formatSheet === "function") {
    try { formatSheet(sheet, headers); } catch {}
  }

  return rows.length;
}

/**
 * Convenience: write property IDs from a specific Notion page.
 *
 * @param pageIdOrUrl Page ID or URL.
 * @param sheetName   Destination sheet name.
 */
function writePagePropIdsToSheet(
    pageIdOrUrl: string,
    sheetName = "Notion Page Property IDs"
): number {
  const page = notionGetPage(pageIdOrUrl);
  const n = writePropIdsToSheet(page, sheetName);
  Logger.log(`Wrote ${n} unique properties to "${sheetName}".`);
  return n;
}

/**
 * Convenience: write property IDs from a Notion data source.
 *
 * @param dsIdOrUrl   Data source ID or URL.
 * @param sheetName   Destination sheet name.
 */
function writeDataSourcePropIdsToSheet(
    dsIdOrUrl: string,
    sheetName = "Notion Data Source Property IDs"
): number {
  const ds = notionGetDataSource(dsIdOrUrl);
  const n = writePropIdsToSheet(ds, sheetName);
  Logger.log(`Wrote ${n} unique properties to "${sheetName}".`);
  return n;
}

/**
 * Flatten a Notion property object into a plain string for Sheets.
 *
 * @param prop A Notion property object.
 * @returns    Human-readable string value.
 */
function extractCellValue(prop: any): string {
  if (!prop) return "";
  switch (prop.type) {
    case "title":
      return (prop.title || []).map((t: any) => t?.plain_text || "").join("");
    case "rich_text":
      return (prop.rich_text || []).map((t: any) => t?.plain_text || "").join("");
    case "email":        return prop.email || "";
    case "phone_number": return prop.phone_number || "";
    case "url":          return prop.url || "";
    case "date":         return prop.date?.start || "";
    case "status":       return prop.status?.name || "";
    case "select":       return prop.select?.name || "";
    case "multi_select": return (prop.multi_select || []).map((o: any) => o?.name || "").join(", ");
    case "people":       return (prop.people || []).map((p: any) => p?.name || p?.person?.email || "").join(", ");
    case "number":       return (prop.number ?? "").toString();
    case "checkbox":     return prop.checkbox ? "TRUE" : "FALSE";
    default:             return "";
  }
}
