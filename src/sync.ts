/** Build {label, propId} specs by joining your alias map with Notion's live schema. */

function buildSpecsFromDataSourceWithAliases(
  dsIdOrUrl: string,
  aliases: Record<string, string>
): Array<{ label: string; propId: string; name: string }> {
  const obj = notionGetDataSource(dsIdOrUrl); // or notionGetPage(...)
  const props = obj?.properties || {};
  const specs: Array<{ label: string; propId: string; name: string }> = [];

  for (const [name, prop] of Object.entries<any>(props)) {
    const propIdRaw = String(prop?.id || "");
    const alias = aliases[name] ?? name;        // alias if present, else original
    specs.push({ label: alias, propId: propIdRaw, name });
  }
  return specs;
}
function buildSpecsFromDataSourceWithAliasesOnly(
  dsIdOrUrl: string,
  aliases: Record<string, string>
): Array<{ label: string; propId: string; name: string }> {
  const obj = notionGetDataSource(dsIdOrUrl);
  const props = obj?.properties || {};
  const specs: Array<{ label: string; propId: string; name: string }> = [];

  for (const aliasName of Object.keys(aliases)) {
    const prop = (props as any)[aliasName];
    if (!prop || !prop.id) continue;                 // skip if Notion doesn’t have it
    specs.push({ label: aliases[aliasName] || aliasName, propId: String(prop.id), name: aliasName });
  }
  return specs;
}

function ensureAliasHeadersExact(
  dsIdOrUrl: string,
  sheetName: string,
  aliases: Record<string, string>,
  startCol = 2 // B
) {
  const specs = buildSpecsFromDataSourceWithAliasesOnly(dsIdOrUrl, aliases);
  return ensureHeadersExactByPropId(sheetName, specs, startCol); // your exact writer
}

function ensureHeadersExactByPropId(
  sheetName: string,
  specs: Array<{ label: string; propId: string; name: string }>,
  startCol: number = 2
): { count: number; startCol: number } {
  const ss = SpreadsheetApp.openById(resolveSpreadsheetId());
  const sheet: Sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  // 1) Clear the existing header band we control (row 1 from startCol to end)
  const lastCol = sheet.getMaxColumns();
  if (startCol <= lastCol) {
    sheet.getRange(1, startCol, 1, lastCol - startCol + 1).clearContent().clearNote();
  }

  // 2) Write exactly the headers from specs
  for (let i = 0; i < specs.length; i++) {
    const { label, propId } = specs[i];
    const col = startCol + i;
    setHeaderCellWithId(sheet, 1, col, label, propId); // label in cell, pretty propId in note; no cell-range metadata
  }

  // 3) Rebuild the sheet-level JSON map: { [colNumber]: prettyPropId }
  const colMap: Record<string, string> = {};
  for (let i = 0; i < specs.length; i++) {
    const col = startCol + i;
    colMap[String(col)] = decodeId(specs[i].propId);
  }
  setSheetMeta(sheet, META_KEY_COLMAP, JSON.stringify(colMap));

  // 4) Niceties: freeze header row and resize written columns
  try {
    sheet.setFrozenRows(Math.max(sheet.getFrozenRows(), 1));
    if (specs.length > 0) sheet.autoResizeColumns(startCol, specs.length);
  } catch { }

  Logger.log(`✅ ensureHeadersExactByPropId: wrote ${specs.length} headers at row 1 from col ${startCol}`);
  return { count: specs.length, startCol };
}

/** Write headers at B1 using alias labels; store prop IDs in header NOTES for stable mapping. */
function ensureAliasHeadersFromDataSourceWithMap(
  dsIdOrUrl: string,
  sheetName: string,
  aliases: Record<string, string>
) {
  const specs = buildSpecsFromDataSourceWithAliases(dsIdOrUrl, aliases);
  return ensureHeadersByPropId(sheetName, specs, /*startCol=*/2); // uses notes for IDs
}

/** Warn if your alias map references names Notion doesn't have. */
function verifyAliasCoverage(dsIdOrUrl: string, aliases: Record<string, string>) {
  const obj = notionGetDataSource(dsIdOrUrl);
  const names = new Set(Object.keys(obj?.properties || {}));
  for (const k of Object.keys(aliases)) {
    if (!names.has(k)) Logger.log(`⚠️ alias key not found in Notion schema: "${k}"`);
  }
}

/** Example: write a row by prop IDs without scanning all props each time */
function writePageRowFast(pageIdOrUrl: string, dsIdOrUrl: string, sheetName = "People Sync") {
  // 1) Load the cached ID→name map (assumes you ran refreshIdNameMapFromDataSource)
  const idName = loadIdNameMap("PEOPLE_ID2NAME");

  // 2) Ensure headers exist at B1 with IDs in notes (from previous step)
  //    (call your ensureAliasHeadersFromDataSourceWithMap(...) earlier in your sync)

  // 3) Fetch the page
  const page = notionGetPage(pageIdOrUrl);

  // 4) Pick which IDs/columns you want to write (order arbitrary)
  const ids = [
    "Wp%3DC", // Grey-Box id
    "HA%40l", // Email (Org)
    "QyDj",   // Mandate (Status)
    "title"   // Name (title)
  ];

  // 5) Resolve each value quickly via ID→name map, no O(n) scan
  const values = ids.map(id => extractCellValue(getPropById(page, id, idName)));

  // 6) Drop them into the proper columns by matching header notes (ID)
  const ss = SpreadsheetApp.openById(resolveSpreadsheetId());
  const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  const startCol = 2; // B
  for (let i = 0; i < ids.length; i++) {
    const col = findColumnByPropId(sheet, ids[i], startCol); // uses header notes
    if (col) sheet.getRange(2, col).setValue(values[i]);     // example writes to row 2
  }
}

/** One-shot: fetch schema once (data source), build + save the map */
function refreshIdNameMapFromDataSource(dsIdOrUrl: string, storeKey = "PEOPLE_ID2NAME") {
  const obj = notionGetDataSource(dsIdOrUrl);
  const map = buildIdToNameMap(obj);
  saveIdNameMap(storeKey, map);
  Logger.log(`Saved ID→name map with ${map.size} entries`);
  return map.size;
}
