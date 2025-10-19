/// <reference types="google-apps-script" />

/**
 * 07_props.ts — Notion property utilities (GAS)
 * ------------------------------------------------
 * - Title extraction, property ID/name mapping, diagnostics/logging
 * - ScriptProperties-based caching for IDs and names
 *
 * Requires:
 *  - utils/core.ts: decodeId
 *  - notion/resources.ts: notionGetPage, notionGetDataSource
 *
 * Exposes (global):
 *  - titleOf
 *  - notionPropIdsToRows
 *  - getPropByIdFromMap
 *  - logPropertyIds, logPropertyIdsFromPage, logPropertyIdsFromDataSource
 *  - getPropertyNames, getPropertyNamesFromPage, getPropertyNamesFromDataSource
 *  - getPropertyNameIdPairs, getPropertyNameIdPairsFromPage, getPropertyNameIdPairsFromDataSource
 *  - getPropertyIds, getPropertyIdsFromPage, getPropertyIdsFromDataSource
 *  - saveIdsToProps, loadIdsFromProps, savePropertyNames, loadPropertyNames
 *  - cachePropertyNamesFromPage, cachePropertyNamesFromDataSource
 *  - printEmailOrgId, printIdsForPage
 *  - getCellValueByName, notionExtractCellValue
 */

/* -------------------------------------------------------------------------- */
/*                               Basic extractors                              */
/* -------------------------------------------------------------------------- */

/** Return the Notion page title (first 'title' property plain text). */
function titleOf(page: { properties?: any }): string {
  const props = page?.properties || {};
  for (const p of Object.values(props) as any[]) {
    if (p?.type === "title") return (p.title || []).map((t: any) => t.plain_text || "").join("");
  }
  return "";
}

/**
 * Build table rows of [name, rawId, prettyId, type] from a page/database/data_source object.
 */
function notionPropIdsToRows(obj: { properties?: any }): any[] {
  const props = obj?.properties || {};
  const rows: any[] = [];
  for (const [name, prop] of Object.entries(props) as [string, any][]) {
    const idRaw = String(prop?.id ?? "");
    let idPretty = idRaw;
    try { idPretty = decodeURIComponent(idRaw); } catch {}
    rows.push([name, idRaw, idPretty, prop?.type || ""]);
  }
  return rows;
}

/**
 * Get a property object from a page using its ID via a precomputed id→name map.
 * (Renamed to avoid clashing with other modules' `getPropById`.)
 */
function getPropByIdFromMap(page: any, propId: string, idNameMap: Map<string, string>) {
  const name = idNameMap.get(propId) || idNameMap.get(decodeId(propId));
  return name ? page?.properties?.[name] : null;
}

/* -------------------------------------------------------------------------- */
/*                                    Logs                                     */
/* -------------------------------------------------------------------------- */

/** Log all property names → ids (works for Page or Database/Data Source JSON). */
function logPropertyIds(obj: { properties?: any }): void {
  const props = obj?.properties || {};
  for (const [name, prop] of Object.entries(props) as [string, any][]) {
    const raw = String(prop?.id || "");
    const pretty = raw.includes("%") ? decodeId(raw) : raw;
    Logger.log(`${name} → id=${pretty} (raw=${raw}) type=${prop?.type || "?"}`);
  }
}

/** Log property IDs from a PAGE (by id/url). */
function logPropertyIdsFromPage(pageIdOrUrl: string): void {
  const page = notionGetPage(pageIdOrUrl);
  logPropertyIds(page);
}

/** Log property IDs from a DATA SOURCE / DATABASE (by id/url). */
function logPropertyIdsFromDataSource(dsIdOrUrl: string): void {
  const ds = notionGetDataSource(dsIdOrUrl);
  logPropertyIds(ds);
}

/* -------------------------------------------------------------------------- */
/*                        Names / IDs (derived & cached)                       */
/* -------------------------------------------------------------------------- */

/**
 * Return property names array (optionally sorted; title-first).
 * @param obj   Notion page/database-like object with properties
 * @param opts  sort: locale-aware sort; titleFirst: move title property to front
 */
function getPropertyNames(
    obj: { properties?: any },
    opts: { sort?: boolean; titleFirst?: boolean } = {}
): string[] {
  const { sort = false, titleFirst = true } = opts;
  const props = obj?.properties || {};
  const names = Object.keys(props);

  if (titleFirst) {
    const t = names.find(n => props[n]?.type === "title");
    if (t) {
      const i = names.indexOf(t);
      if (i > 0) { names.splice(i, 1); names.unshift(t); }
    }
  }

  if (sort) {
    try {
      const collator = new Intl.Collator("en", { sensitivity: "base" });
      names.sort(collator.compare);
    } catch { names.sort(); }
  }
  return names;
}

/** Convenience: property names from PAGE. */
function getPropertyNamesFromPage(pageIdOrUrl: string, opts: { sort?: boolean; titleFirst?: boolean } = {}): string[] {
  const page = notionGetPage(pageIdOrUrl);
  return getPropertyNames(page, opts);
}

/** Convenience: property names from DATA SOURCE / DB. */
function getPropertyNamesFromDataSource(dsIdOrUrl: string, opts: { sort?: boolean; titleFirst?: boolean } = {}): string[] {
  const ds = notionGetDataSource(dsIdOrUrl);
  return getPropertyNames(ds, opts);
}

/** Return array of { name, idRaw, idPretty, type }. */
function getPropertyNameIdPairs(
    obj: { properties?: any },
    opts: { titleFirst?: boolean; sort?: boolean } = {}
): Array<{ name: string; idRaw: string; idPretty: string; type: string }> {
  const names = getPropertyNames(obj, opts);
  const props = obj?.properties || {};
  return names.map(name => {
    const idRaw = String(props[name]?.id || "");
    const idPretty = decodeId(idRaw);
    return { name, idRaw, idPretty, type: props[name]?.type || "" };
  });
}

/** Convenience: name/id pairs from PAGE. */
function getPropertyNameIdPairsFromPage(pageIdOrUrl: string, opts: { titleFirst?: boolean; sort?: boolean } = {}) {
  const page = notionGetPage(pageIdOrUrl);
  return getPropertyNameIdPairs(page, opts);
}

/** Convenience: name/id pairs from DATA SOURCE / DB. */
function getPropertyNameIdPairsFromDataSource(dsIdOrUrl: string, opts: { titleFirst?: boolean; sort?: boolean } = {}) {
  const ds = notionGetDataSource(dsIdOrUrl);
  return getPropertyNameIdPairs(ds, opts);
}

/** Get an array of property IDs (raw or pretty). */
function getPropertyIds(obj: { properties?: any }, form: "raw" | "pretty" = "raw"): string[] {
  const props = obj?.properties || {};
  const ids: string[] = [];
  for (const p of Object.values(props) as any[]) {
    const raw = String(p?.id || "");
    ids.push(form === "pretty" ? decodeId(raw) : raw);
  }
  return ids;
}

/** Convenience: IDs from PAGE. */
function getPropertyIdsFromPage(pageIdOrUrl: string, form: "raw" | "pretty" = "raw"): string[] {
  const page = notionGetPage(pageIdOrUrl);
  return getPropertyIds(page, form);
}

/** Convenience: IDs from DATA SOURCE / DB. */
function getPropertyIdsFromDataSource(dsIdOrUrl: string, form: "raw" | "pretty" = "raw"): string[] {
  const ds = notionGetDataSource(dsIdOrUrl);
  return getPropertyIds(ds, form);
}

/* -------------------------------------------------------------------------- */
/*                           Properties storage cache                          */
/* -------------------------------------------------------------------------- */

/** Save/load arrays into Script Properties (IDs or names). */
function saveIdsToProps(key: string, ids: string[]): void {
  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(ids));
}
function loadIdsFromProps(key: string): string[] {
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) return [];
  try { return JSON.parse(raw); } catch { return []; }
}

function savePropertyNames(key: string, names: string[]): void {
  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(names));
}
function loadPropertyNames(key: string): string[] {
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) return [];
  try { return JSON.parse(raw); } catch { return []; }
}

/** One-shot cache helpers (persist property names by source/page). */
function cachePropertyNamesFromPage(pageIdOrUrl: string, storeKey: string): string[] {
  const arr = getPropertyNamesFromPage(pageIdOrUrl, { titleFirst: true });
  savePropertyNames(storeKey, arr);
  return arr;
}
function cachePropertyNamesFromDataSource(dsIdOrUrl: string, storeKey: string): string[] {
  const arr = getPropertyNamesFromDataSource(dsIdOrUrl, { titleFirst: true });
  savePropertyNames(storeKey, arr);
  return arr;
}

/* -------------------------------------------------------------------------- */
/*                                Diagnostics                                  */
/* -------------------------------------------------------------------------- */

/** Example: print a specific property ID from a PAGE. */
function printEmailOrgId(pageIdOrUrl: string): void {
  const page = notionGetPage(pageIdOrUrl);
  const raw = String(page.properties?.["Email (Org)"]?.id || "");
  const pretty = decodeId(raw);
  Logger.log(`Email (Org) id = ${pretty} (raw=${raw})`);
}

/** Convenience: print all property IDs from a PAGE. */
function printIdsForPage(pageIdOrUrl: string): void {
  const page = notionGetPage(pageIdOrUrl);
  logPropertyIds(page);
}

/* -------------------------------------------------------------------------- */
/*                           Cell value convenience                            */
/* -------------------------------------------------------------------------- */

/** Return a plain string for a given property NAME on a page. */
function getCellValueByName(page: any, propName: string): string {
  const prop = page?.properties?.[propName];
  return notionExtractCellValue(prop);
}

/**
 * Convert a Notion property object into a plain string for Sheets/logging.
 * (Renamed to avoid colliding with sheets/03_writes.ts's extractCellValue.)
 * Handles common types; unknown types return "".
 */
function notionExtractCellValue(prop: any): string {
  if (!prop) return "";

  switch (prop.type) {
    case "title":
      return (prop.title || []).map((t: any) => t?.plain_text || "").join("");

    case "rich_text":
      return (prop.rich_text || []).map((t: any) => t?.plain_text || "").join("");

    case "email":        return prop.email || "";
    case "phone_number": return prop.phone_number || "";
    case "url":          return prop.url || "";

    case "date":
      // prefer start; if you need ranges, include end as well.
      return prop.date?.start || "";

    case "status":       return prop.status?.name || "";
    case "select":       return prop.select?.name || "";

    case "multi_select":
      return (prop.multi_select || []).map((o: any) => o?.name || "").join(", ");

    case "people":
      // Try display name; fall back to email if present
      return (prop.people || [])
          .map((p: any) => p?.name || p?.person?.email || "")
          .filter(Boolean)
          .join(", ");

    case "number":
      return (prop.number ?? "").toString();

    case "checkbox":
      return prop.checkbox ? "TRUE" : "FALSE";

      // ---- Optional extras you may want ----
    case "relation":
      // Safe default: count of related records.
      return String((prop.relation || []).length || 0);

    case "files":
      // Join file names (or URLs if you prefer)
      return (prop.files || [])
          .map((f: any) => f?.name || f?.file?.url || f?.external?.url || "")
          .filter(Boolean)
          .join(", ");

    case "formula": {
      const f = prop.formula || {};
      switch (f.type) {
        case "string":  return f.string || "";
        case "number":  return (f.number ?? "").toString();
        case "boolean": return f.boolean ? "TRUE" : "FALSE";
        case "date":    return f.date?.start || "";
        default:        return "";
      }
    }

    case "rollup": {
      const r = prop.rollup || {};
      if (r.type === "number") return (r.number ?? "").toString();
      if (r.type === "date")   return r.date?.start || "";
      if (r.type === "array") {
        const arr = r.array || [];
        const values = arr.map((x: any) => {
          if (x?.type) return notionExtractCellValue(x);
          return typeof x === "string" ? x : "";
        }).filter(Boolean);
        return values.join(", ");
      }
      return "";
    }

    case "created_by":
    case "last_edited_by":
      return prop[prop.type]?.name || "";

    case "created_time":
    case "last_edited_time":
      return prop[prop.type] || "";

    default:
      return "";
  }
}