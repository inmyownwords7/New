/**
 * notion/orchestrator.ts
 * High-level flows for pulling from Notion and shaping rows for Sheets.
 *
 * Requires (already loaded as globals):
 * - notion/schema.ts: buildSpecsFromDataSourceWithAliasesOnlyCI, buildSpecsFromDataSourceWithAliases
 * - notion/query.ts:  queryDataSourceAll
 * - notion/resources.ts: notionGetPage, notionGetDataSource
 * - sheets/writes.ts: extractCellValue
 * - sheets/headers.ts: ensureHeaders, ensureHeadersByPropId, ensureHeadersExactByPropId, setHeaderCellWithId, findColumnByPropId
 * - sheets/access.ts: getSheetByNameOrCreate, withSheetLock
 * - sheets/state.ts:  buildIdToNameMap, saveIdNameMap, loadIdNameMap, setSheetMeta, getSheetMeta, META_KEY_COLMAP
 * - utils/core.ts:    decodeId
 */

// ---------- Public Orchestrator API ----------

/** Build specs from aliases, using your existing CI match helper */
function buildSpecsFromAliases(dsIdOrUrl: string, aliases: Record<string, string>): NotionSpec[] {
  return buildSpecsFromDataSourceWithAliasesOnlyCI(dsIdOrUrl, aliases);
}

/** Fetch all pages from a data source */
function fetchAllPages(dsIdOrUrl: string, queryBody: Record<string, unknown> = {}): any[] {
  return queryDataSourceAll(dsIdOrUrl, queryBody);
}

/** Turn specs into column headers (labels) */
function makeHeadersFromSpecs(specs: NotionSpec[]): string[] {
  return specs.map(s => s.label || s.name);
}

/** Convert Notion pages → 2D rows aligned to given specs (one row per page) */
function pagesToRows(pages: any[], specs: NotionSpec[]): any[][] {
  const idNameMap = buildIdNameMapFromPageOrDb(pages?.[0]); // best-effort for getPropById helper if you use it

  return pages.map(page => {
    return specs.map(spec => {
      // Prefer fast lookup via propId→name map when possible
      const propObj = idNameMap
        ? getPropById(page, spec.propId, idNameMap)
        : (page?.properties?.[spec.name] ?? null);
      return stringifyNotionProp(propObj);
    });
  });
}

/**
 * End-to-end: read Notion → ensure headers → append/upsert rows into a sheet.
 * opts.mode = "append" (default) or "upsert"
 * opts.keyLabel = header label used as unique key for upsert (required when mode="upsert")
 * opts.batchSize = write chunk size (default 500)
 */
function syncDataSourceToSheet(
  dsIdOrUrl: string,
  aliases: Record<string, string>,
  sheetName: string,
  opts: { mode?: "append" | "upsert"; keyLabel?: string; batchSize?: number } = {}
) {
  const { mode = "append", keyLabel, batchSize = 500 } = opts;

  // 1) Specs & headers
  const specs = buildSpecsFromAliases(dsIdOrUrl, aliases);
  const headers = makeHeadersFromSpecs(specs);

  // 2) Fetch all pages
  const pages = fetchAllPages(dsIdOrUrl);

  // 3) Shape into rows
  const rows = pagesToRows(pages, specs);

  // 4) Write to sheet
  const ss = SpreadsheetApp.getActive();
  const sheet = getSheetByNameOrCreate(ss, sheetName);

  withSheetLock(() => {
    const { changed } = ensureHeaders(sheet, headers);

    if (mode === "upsert") {
      if (!keyLabel) throw new Error('syncDataSourceToSheet: keyLabel is required when mode="upsert"');
      const res = upsertRowsByKey(sheet, keyLabel, headers, rows);
      Logger.log(`Upsert complete → inserted=${res.inserted}, updated=${res.updated}`);
      return;
    }

    // default: append
    appendRowsBatched(sheet, rows, batchSize);
    Logger.log(`Append complete → rows=${rows.length}`);
  });
}

// ---------- Internals / helpers ----------

/** Build a map of propId(raw/decoded) → name from a representative page or db */
function buildIdNameMapFromPageOrDb(obj: any): Map<string, string> | null {
  const props = obj?.properties;
  if (!props || typeof props !== "object") return null;

  const m = new Map<string, string>();
  for (const [name, p] of Object.entries<any>(props)) {
    const raw = String(p?.id ?? "");
    if (!raw) continue;
    m.set(raw, name);
    try { m.set(decodeURIComponent(raw), name); } catch {}
  }
  return m;
}

/** Convert a single Notion property object to a displayable/flat value */
function stringifyNotionProp(prop: any): any {
  if (!prop || typeof prop !== "object") return "";

  switch (prop.type) {
    case "title":
      return (prop.title || []).map((t: any) => t.plain_text ?? "").join("");
    case "rich_text":
      return (prop.rich_text || []).map((t: any) => t.plain_text ?? "").join("");
    case "number":
      return prop.number ?? "";
    case "checkbox":
      return !!prop.checkbox;
    case "status":
      return prop.status?.name ?? "";
    case "select":
      return prop.select?.name ?? "";
    case "multi_select":
      return (prop.multi_select || []).map((o: any) => o.name).join(", ");
    case "people":
      return (prop.people || []).map((p: any) => p?.name || p?.person?.email || p?.id).join(", ");
    case "email":
      return prop.email ?? "";
    case "phone_number":
      return prop.phone_number ?? "";
    case "url":
      return prop.url ?? "";
    case "date":
      if (!prop.date) return "";
      if (prop.date.start && prop.date.end) return `${prop.date.start} → ${prop.date.end}`;
      return prop.date.start ?? "";
    case "files":
      return (prop.files || [])
        .map((f: any) => f?.name || f?.file?.url || f?.external?.url)
        .filter(Boolean)
        .join(", ");
    case "relation":
      return (prop.relation || []).map((r: any) => r?.id).join(", ");
    case "rollup":
      return stringifyRollup(prop);
    case "formula":
      return stringifyFormula(prop.formula);
    default:
      try { return JSON.stringify(prop); } catch { return String(prop); }
  }
}

function stringifyRollup(prop: any): any {
  if (!prop || prop.type !== "rollup") return "";
  const r = prop.rollup;
  if (!r) return "";

  switch (r.type) {
    case "number":
      return r.number ?? "";
    case "date":
      if (!r.date) return "";
      if (r.date.start && r.date.end) return `${r.date.start} → ${r.date.end}`;
      return r.date.start ?? "";
    case "array":
      return (r.array || [])
        .map((x: any) => {
          if (!x) return "";
          if (x.type === "title") return (x.title || []).map((t: any) => t.plain_text ?? "").join("");
          if (x.type === "rich_text") return (x.rich_text || []).map((t: any) => t.plain_text ?? "").join("");
          if (x.type === "people") return (x.people || []).map((p: any) => p?.name || p?.person?.email || p?.id).join(", ");
          if (x.type === "select") return x.select?.name ?? "";
          if (x.type === "multi_select") return (x.multi_select || []).map((o: any) => o.name).join(", ");
          if (x.type === "status") return x.status?.name ?? "";
          if (x.type === "number") return x.number ?? "";
          try { return JSON.stringify(x); } catch { return String(x); }
        })
        .join("; ");
    default:
      try { return JSON.stringify(r); } catch { return String(r); }
  }
}

function stringifyFormula(f: any): any {
  if (!f || typeof f !== "object") return "";
  switch (f.type) {
    case "string": return f.string ?? "";
    case "number": return f.number ?? "";
    case "boolean": return !!f.boolean;
    case "date":
      if (!f.date) return "";
      if (f.date.start && f.date.end) return `${f.date.start} → ${f.date.end}`;
      return f.date.start ?? "";
    default:
      try { return JSON.stringify(f); } catch { return String(f); }
  }
}