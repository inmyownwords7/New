/*************************************************
 * sheets/03_writes.ts
 * ------------------------------------------------
 * Handles all row-level write operations to Google Sheets:
 * - Writing or appending rows
 * - Upserting data by key
 * - Writing Notion property IDs
 * - Flattening Notion properties into plain text
 *************************************************/

/**
 * Append rows to a sheet in batches to avoid quota limits.
 * @param sheet  Target sheet.
 * @param rows   2D array of rows.
 * @param batchSize  Maximum rows per batch (default 500).
 */
function appendRowsBatched(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  rows: any[][],
  batchSize: number = 500
): void {
  if (!rows.length) return;

  for (let i = 0; i < rows.length; i += batchSize) {
    const chunk = rows.slice(i, i + batchSize);
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, chunk.length, chunk[0].length).setValues(chunk);
  }
}

/**
 * Upsert rows by matching a key column.
 * - If key exists → update that row
 * - If missing → append new row
 */
function upsertRowsByKey(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  keyHeader: string,
  headers: string[],
  rows: any[][]
): { inserted: number; updated: number } {
  let inserted = 0;
  let updated = 0;

  if (!rows.length) return { inserted, updated };

  const keyIndex = headers.indexOf(keyHeader);
  if (keyIndex === -1) throw new Error(`Key header "${keyHeader}" not found.`);

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const headerRow = values.shift() || [];
  const keyCol = headerRow.indexOf(keyHeader);
  if (keyCol === -1) throw new Error(`Key header "${keyHeader}" not found in sheet.`);

  // Build a map of key → rowIndex
  const map = new Map<string, number>();
  for (let i = 0; i < values.length; i++) {
    const key = String(values[i][keyCol] || "").trim();
    if (key) map.set(key, i + 2); // +2 because header + 1-based index
  }

  // Perform upsert
  for (const row of rows) {
    const key = String(row[keyIndex] || "").trim();
    if (!key) continue;
    const existingRow = map.get(key);
    if (existingRow) {
      sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
      updated++;
    } else {
      sheet.appendRow(row);
      inserted++;
    }
  }

  return { inserted, updated };
}

/**
 * Write contiguous values starting from a given row and column.
 */
function writeContiguousRow(
  sheetName: string,
  row: number,
  startCol: number,
  values: string[]
) {
  const ss = SpreadsheetApp.openById(resolveSpreadsheetId());
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  if (!values.length) return;
  sheet.getRange(row, startCol, 1, values.length).setValues([values]);
}

/**
 * Write a Notion object’s property IDs to a “Property IDs” sheet.
 */
function writePropIdsToSheet(obj: { properties?: any }, sheetName = "Notion Property IDs") {
  const headers = ["Property Name", "Property ID (raw)", "Property ID (pretty)", "Type"];
  const sheet = getOrCreateSheetByName(sheetName, headers);

  // Clear old rows
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
 * Convenience: write page property IDs to a dedicated sheet.
 */
function writePagePropIdsToSheet(pageIdOrUrl: string, sheetName = "Notion Page Property IDs") {
  const page = notionGetPage(pageIdOrUrl);
  const n = writePropIdsToSheet(page, sheetName);
  Logger.log(`Wrote ${n} unique properties to "${sheetName}".`);
  return n;
}

/**
 * Convenience: write data source property IDs to a dedicated sheet.
 */
function writeDataSourcePropIdsToSheet(dsIdOrUrl: string, sheetName = "Notion Data Source Property IDs") {
  const ds = notionGetDataSource(dsIdOrUrl);
  const n = writePropIdsToSheet(ds, sheetName);
  Logger.log(`Wrote ${n} unique properties to "${sheetName}".`);
  return n;
}

/**
 * Flatten a Notion property object into a plain string for Sheets.
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