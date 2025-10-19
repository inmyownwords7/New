/// <reference types="google-apps-script" />

/**
 * sheets/state.ts — Shared state & caches (GAS)
 * ------------------------------------------------
 * Responsibilities:
 *  • Sheet-level JSON map for Notion propId → column (via DeveloperMetadata)
 *  • ID↔name map cache stored in ScriptProperties
 *  • Helpers for reading/writing DeveloperMetadata and cached maps
 *
 * Globals assumed:
 *  - Sheet, DevMeta, decodeId
 */

/* -------------------------------------------------------------------------- */
/*                         DeveloperMetadata accessors                         */
/* -------------------------------------------------------------------------- */

/**
 * Read a developer-metadata value by key from a sheet.
 *
 * @param sheet Target Google Sheet.
 * @param key   Metadata key name.
 * @returns     The stored value or null if not found.
 */
function getSheetMeta(sheet: Sheet, key: string): string | null {
  const metas = sheet.getDeveloperMetadata() as DevMeta[];
  const hit = metas && metas.find(m => m.getKey() === key);
  return hit ? hit.getValue() : null;
}

/**
 * Upsert a developer-metadata key/value pair on a sheet.
 * Removes any existing entries with the same key before adding the new one.
 *
 * @param sheet Target Google Sheet.
 * @param key   Metadata key name.
 * @param value Value to store (as string).
 */
function setSheetMeta(sheet: Sheet, key: string, value: string): void {
  const metas = sheet.getDeveloperMetadata() as DevMeta[];
  metas.forEach(m => { if (m.getKey() === key) m.remove(); });
  sheet.addDeveloperMetadata(key, value);
}

/**
 * Remove a developer-metadata key from a sheet (no-op if missing).
 *
 * @param sheet Target Google Sheet.
 * @param key   Metadata key to remove.
 */
function deleteSheetMeta(sheet: Sheet, key: string): void {
  const metas = sheet.getDeveloperMetadata() as DevMeta[];
  metas.forEach(m => { if (m.getKey() === key) m.remove(); });
}

/* -------------------------------------------------------------------------- */
/*                         Notion ID↔name map helpers                         */
/* -------------------------------------------------------------------------- */

/**
 * Build an ID→name map from a Notion page or database object's properties.
 * Saves both raw and decoded IDs for more robust lookups.
 *
 * @param obj Notion page or database object with `.properties`.
 * @returns   Map of (id raw/decoded) → property name.
 */
function buildIdToNameMap(obj: { properties?: any }): Map<string, string> {
  const m = new Map<string, string>();
  const props = obj?.properties || {};
  for (const [name, prop] of Object.entries<any>(props)) {
    const raw = String(prop?.id || "");
    if (!raw) continue;
    m.set(raw, name);
    m.set(decodeId(raw), name); // store decoded version too
  }
  return m;
}

/**
 * Save an ID→name map to Script Properties (durable cache).
 *
 * @param key  Unique cache key (e.g., "PeopleDirectory_IdNameMap").
 * @param map  Map to save.
 */
function saveIdNameMap(key: string, map: Map<string, string>): void {
  const arr = Array.from(map.entries());
  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(arr));
}

/**
 * Load a previously saved ID→name map from Script Properties.
 *
 * @param key  Unique cache key used when saving.
 * @returns    Restored map (empty if not found or invalid JSON).
 */
function loadIdNameMap(key: string): Map<string, string> {
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) return new Map();
  try {
    return new Map(JSON.parse(raw) as [string, string][]);
  } catch {
    return new Map();
  }
}

/* -------------------------------------------------------------------------- */
/*                        Column→propId map (DeveloperMeta)                   */
/* -------------------------------------------------------------------------- */

/**
 * Load the column → decodedPropId mapping from DeveloperMetadata.
 * Stored as { [colNumber: string]: decodedPropId }.
 *
 * @param sheet Target Google Sheet.
 * @returns     Parsed object mapping columns to propIds.
 */
function loadColumnPropIdMap(sheet: Sheet): Record<string, string> {
  const raw = getSheetMeta(sheet, META_KEY_COLMAP);
  if (!raw) return {};
  try {
    return JSON.parse(raw) as Record<string, string>;
  } catch {
    return {};
  }
}

/**
 * Save the column → decodedPropId mapping to DeveloperMetadata.
 *
 * @param sheet  Target Google Sheet.
 * @param colMap Object like { "2": "Wp@C", "3": "fk^Y", ... }.
 */
function saveColumnPropIdMap(sheet: Sheet, colMap: Record<string, string>): void {
  setSheetMeta(sheet, META_KEY_COLMAP, JSON.stringify(colMap));
}

/**
 * Set or update a single column’s decoded propId in the stored map.
 *
 * @param sheet           Target Google Sheet.
 * @param col1Based       1-based column index.
 * @param decodedPropId   Decoded property id (from decodeId()).
 */
function setColumnPropId(sheet: Sheet, col1Based: number, decodedPropId: string): void {
  const map = loadColumnPropIdMap(sheet);
  map[String(col1Based)] = decodedPropId || "";
  saveColumnPropIdMap(sheet, map);
}

/**
 * Retrieve a single column’s decoded propId from the stored map.
 *
 * @param sheet     Target Google Sheet.
 * @param col1Based 1-based column index.
 * @returns         Decoded property id or "" if not found.
 */
function getColumnPropId(sheet: Sheet, col1Based: number): string {
  const map = loadColumnPropIdMap(sheet);
  return map[String(col1Based)] || "";
}
