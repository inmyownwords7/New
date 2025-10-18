/**
 * state.ts — shared state & caches (GAS)
 * Owns:
 *  - Sheet-level JSON map for Notion propId → column (via DeveloperMetadata)
 *  - ID↔name map cache in ScriptProperties
 *  - Helper to build ID→name from a Notion schema object
 *
 * Globals assumed: Sheet, DevMeta, decodeId
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

/** Build ID→name map from a page/database object */
function buildIdToNameMap(obj: { properties?: any }): Map<string, string> {
  const m = new Map<string, string>();
  const props = obj?.properties || {};
  for (const [name, prop] of Object.entries<any>(props)) {
    const raw = String(prop?.id || "");
    if (!raw) continue;
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