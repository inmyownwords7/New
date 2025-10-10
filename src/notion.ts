/** Notion API (runtime) */
function getHeaderCI(headers: Record<string, string>, name: string): string | undefined {
  const n = name.toLowerCase();
  for (const k in headers) if (k && k.toLowerCase() === n) return headers[k];
  return undefined;
}

function normKey(s: string): string {
  return s
    ? s.normalize("NFKC").trim().replace(/\s+/g, " ").toLowerCase()
    : "";
}

function buildNameIndexCI(props: Record<string, any>): Map<string, { name: string; id: string }> {
  const idx = new Map<string, { name: string; id: string }>();
  for (const [name, prop] of Object.entries(props)) {
    const id = String((prop as any)?.id || "");
    if (!id) continue;
    idx.set(normKey(name), { name, id });
  }
  return idx;
}

/** Aliases-only, case-insensitive match against Notion schema */
function buildSpecsFromDataSourceWithAliasesOnlyCI(
  dsIdOrUrl: string,
  aliases: Record<string, string>
): NotionSpec[] {
  const obj = notionGetDataSource(dsIdOrUrl);
  const props = (obj as any)?.properties || {};
  const index = buildNameIndexCI(props);

  const specs: NotionSpec[] = [];
  for (const aliasKey of Object.keys(aliases)) {
    // 1) If alias key looks like an ID, match by ID
    if (looksLikeId(aliasKey)) {
      const pretty = decodeId(aliasKey);
      const match = Object.entries<any>(props).find(([_, p]) => {
        const id = String(p?.id || "");
        return decodeId(id) === pretty || id === aliasKey;
      });
      if (match) {
        const [realName, p] = match;
        specs.push({ label: aliases[aliasKey], propId: String(p.id), name: realName });
      } else {
        Logger.log(`⚠️ No property with id "${aliasKey}"`);
      }
      continue;
    }

    // 2) Otherwise match by NAME (case/spacing-insensitive)
    const hit = index.get(normKey(aliasKey));
    if (!hit) {
      Logger.log(`⚠️ Alias not found (CI): "${aliasKey}"`);
      continue;
    }
    specs.push({ label: aliases[aliasKey] || hit.name, propId: hit.id, name: hit.name });
  }
  return specs;
}

function notionApi<T = unknown>(params: NotionApiParams): NotionApiResult<T> {
  const NOTION_BASE = "https://api.notion.com";

  const {
    method = "GET",
    path,
    query,
    body,
    token = PropertiesService.getScriptProperties().getProperty("NOTION_TOKEN") || "",
    version = PropertiesService.getScriptProperties().getProperty("NOTION_VERSION") || "2025-09-03",
    throwOnHttpError = true,
    debug = false,
  } = params;

  if (!path || !path.startsWith("/")) throw new Error('notionApi: "path" must start with "/"');
  if (!token) throw new Error("notionApi: missing NOTION_TOKEN");

  // Build URL
  let url = NOTION_BASE + path;
  if (query) {
    const qs: string[] = [];
    for (const [k, v] of Object.entries(query)) {
      if (v == null) continue;
      if (Array.isArray(v)) for (const it of v) qs.push(`${encodeURIComponent(k)}=${encodeURIComponent(String(it))}`);
      else qs.push(`${encodeURIComponent(k)}=${encodeURIComponent(String(v))}`);
    }
    if (qs.length) url += "?" + qs.join("&");
  }

  // Headers
  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    "Notion-Version": version,
  };

  // Options
  const gasMethod = toGasMethod(method);
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: gasMethod,
    headers,
    muteHttpExceptions: true,
  };

  // Body/payload
  const needsBody = gasMethod === "post" || gasMethod === "patch";
  if (needsBody && body !== undefined) {
    if (typeof body === "string") {
      options.payload = body as FetchPayload;
      options.contentType = headers["Content-Type"] || "application/json";
    } else if (typeof body === "object" && "getBytes" in (body as object)) {
      const blob = body as GoogleAppsScript.Base.Blob;
      options.payload = blob as FetchPayload;
      const blobType = blob.getContentType();
      if (blobType && !headers["Content-Type"]) options.contentType = blobType;
    } else {
      options.payload = JSON.stringify(body) as FetchPayload;
      options.contentType = "application/json";
    }
  }

  if (debug) {
    const safe = { ...headers, Authorization: "Bearer ***redacted***" };
    Logger.log(
      `notionApi → ${gasMethod.toUpperCase()} ${url}\nheaders=${JSON.stringify(safe)}\n` +
      `contentType=${options.contentType || "(none)"}  hasPayload=${options.payload != null}`
    );
  }

  const resp = UrlFetchApp.fetch(url, options);
  const status = resp.getResponseCode();
  const respHeaders = resp.getHeaders() as unknown as Record<string, string>;
  const ctype = respHeaders["Content-Type"] || "";
  const text = resp.getContentText();

  let data: unknown = text;
  if (ctype.includes("application/json")) {
    try { data = JSON.parse(text); } catch {}
  }

  if (throwOnHttpError && (status < 200 || status >= 300)) {
    throw new Error(`notionApi: HTTP ${status} → ${text}`);
  }

  return { ok: status >= 200 && status < 300, status, data: data as T, headers: respHeaders, url, method: gasMethod };
}
/** Normalize to GAS's lowercase HttpMethod */
function toGasMethod(m: AnyCaseHttpMethod = "GET"): GoogleAppsScript.URL_Fetch.HttpMethod {
  return (String(m).toLowerCase() as GoogleAppsScript.URL_Fetch.HttpMethod);
}



/** Type guards (runtime) */
function isDataSource(x: unknown): x is NotionDataSource {
  return !!x && typeof x === "object" && (x as any).object === "data_source";
}
function isDatabase(x: unknown): x is NotionDatabase {
  return !!x && typeof x === "object" && (x as any).object === "database";
}

/** Fetchers */
function notionGetDataSource(idOrUrl: string): NotionDataSource | NotionDatabase {
  const id = normalizeUuid(extractId32(idOrUrl));
  if (!id) throw new Error("notionGetDataSource: missing ID/URL");

  const ds = notionApi<NotionDataSource | { object?: "error"; code?: string; message?: string }>({
    method: "GET",
    path: `/v1/data_sources/${id}`,
    throwOnHttpError: false,
  });
  if (ds.ok && isDataSource(ds.data)) return ds.data;

  const db = notionApi<NotionDatabase | { object?: "error"; code?: string; message?: string }>({
    method: "GET",
    path: `/v1/databases/${id}`,
    throwOnHttpError: false,
  });
  if (db.ok && isDatabase(db.data)) return db.data;

  throw new Error(
    `/v1/data_sources/${id} → ${ds.status} ${safeJson(ds.data)}\n` +
    `/v1/databases/${id} → ${db.status} ${safeJson(db.data)}`
  );
}

function notionGetPage(idOrUrl: string): NotionPage {
  const id = normalizeUuid(extractId32(idOrUrl));
  if (!id) throw new Error("notionGetPage: missing ID/URL");
  const r = notionApi<NotionPage | { object?: "error" }>({ method: "GET", path: `/v1/pages/${id}`, throwOnHttpError: false });
  if (r.ok && (r.data as NotionPage).object === "page") return r.data as NotionPage;
  throw new Error(`GET /v1/pages/${id} → ${r.status} ${safeJson(r.data)}`);
}

/** Query all pages from a data source */
function queryDataSourceAll(
  dsIdOrUrl: string,
  queryBody: Record<string, unknown> = {},
  opts: { pageSize?: number; debug?: boolean } = {}
) {
  const { pageSize = 100, debug = false } = opts;
  const id = normalizeUuid(extractId32(dsIdOrUrl));
  const base = { page_size: pageSize, ...queryBody };

  let r = notionApi<{ results?: any[]; has_more?: boolean; next_cursor?: string }>({
    method: "POST",
    path: `/v1/data_sources/${id}/query`,
    body: base,
    throwOnHttpError: false,
    debug,
  });
  if (!r.ok) throw new Error(`POST /v1/data_sources/${id}/query → ${r.status} ${safeJson(r.data)}`);

  const all = [...(r.data.results || [])];
  let cursor = r.data.has_more && r.data.next_cursor ? r.data.next_cursor : null;

  while (cursor) {
    r = notionApi<{ results?: any[]; has_more?: boolean; next_cursor?: string }>({
      method: "POST",
      path: `/v1/data_sources/${id}/query`,
      body: { ...base, start_cursor: cursor },
      throwOnHttpError: false,
      debug,
    });
    if (!r.ok) throw new Error(`pagination → ${r.status} ${safeJson(r.data)}`);
    all.push(...(r.data.results || []));
    cursor = r.data.has_more && r.data.next_cursor ? r.data.next_cursor : null;
  }
  return all;
}

/** Utils */
function titleOf(page: { properties?: any }): string {
  const props = page?.properties || {};
  for (const p of Object.values(props) as any[]) {
    if (p?.type === "title") return (p.title || []).map((t: any) => t.plain_text).join("");
  }
  return "";
}

function safeJson(x: unknown): string {
  try { return JSON.stringify(x); } catch { return String(x); }
}

function notionPropIdsToRows(obj: { properties?: any }) {
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


function sleepMs(ms: number) { Utilities.sleep(ms); }

function notionFetchWithRetry(
  url: string,
  options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions
) {
  const maxAttempts = 5;
  let delayMs = 250;
  let lastErr: any = null;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const resp = UrlFetchApp.fetch(url, options);
      const status = resp.getResponseCode();

      // TS-safe: cast headers to an indexable dictionary
      const headers = resp.getHeaders() as unknown as Record<string, string>;

      if (status === 429 || status >= 500) {
        const ra = Number(getHeaderCI(headers, "Retry-After") || 0);
        const wait = ra > 0 ? ra * 1000 : delayMs;
        Utilities.sleep(wait);
        delayMs *= 2;
        continue;
      }
      return resp;
    } catch (e) {
      lastErr = e;
      Utilities.sleep(delayMs);
      delayMs *= 2;
    }
  }
  throw lastErr || new Error("notionFetchWithRetry: failed after retries");
}

/** Get a property object from a page using its ID (fast via cached map). */
function getPropById(page: any, propId: string, idNameMap: Map<string, string>) {
  const name = idNameMap.get(propId) || idNameMap.get(decodeId(propId));
  return name ? page?.properties?.[name] : null;
}


// Extract a 32-hex id from a string or return the input as string
function extractId32(input: unknown): string {
  if (!input) return "";
  const m = String(input).match(/[0-9a-f]{32}/i);
  return m ? m[0] : String(input);
}

// Normalize 32-hex → UUID with dashes; otherwise return as-is
function normalizeUuid(id: string): string {
  const m = String(id).match(/^[0-9a-fA-F]{32}$/);
  if (!m) return id; // already hyphenated or not 32-hex
  const r = m[0].toLowerCase();
  return `${r.slice(0,8)}-${r.slice(8,12)}-${r.slice(12,16)}-${r.slice(16,20)}-${r.slice(20)}`;
}
/*** ─────────────────────────────────────────────
 * Extras you had before (utilities & diagnostics)
 * Paste these BELOW your current code
 * ──────────────────────────────────────────── ***/

/** decode a Notion property id like 'HA%40l' → 'HA@l' */
function decodeId(id: unknown): string {
  if (!id) return "";
  try { return decodeURIComponent(String(id)); } catch { return String(id); }
}

/** Log all property names → ids (works for Page or Database/Data Source JSON) */
function logPropertyIds(obj: { properties?: any }): void {
  const props = obj?.properties || {};
  for (const [name, prop] of Object.entries(props) as [string, any][]) {
    const raw = String(prop?.id || "");
    const pretty = raw.includes("%") ? decodeId(raw) : raw;
    Logger.log(`${name}  →  id=${pretty}  (raw=${raw})  type=${prop?.type || "?"}`);
  }
}

/** From a PAGE id/url */
function logPropertyIdsFromPage(pageIdOrUrl: string): void {
  const page = notionGetPage(pageIdOrUrl);
  logPropertyIds(page);
}

/** From a DATA SOURCE / DATABASE id/url */
function logPropertyIdsFromDataSource(dsIdOrUrl: string): void {
  const ds = notionGetDataSource(dsIdOrUrl);
  logPropertyIds(ds);
}

/** Return property names array (optionally sorted; title-first) */
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

/** Get property names from PAGE */
function getPropertyNamesFromPage(pageIdOrUrl: string, opts: { sort?: boolean; titleFirst?: boolean } = {}): string[] {
  const page = notionGetPage(pageIdOrUrl);
  return getPropertyNames(page, opts);
}

/** Get property names from DATA SOURCE / DB */
function getPropertyNamesFromDataSource(dsIdOrUrl: string, opts: { sort?: boolean; titleFirst?: boolean } = {}): string[] {
  const ds = notionGetDataSource(dsIdOrUrl);
  return getPropertyNames(ds, opts);
}

/** Return array of { name, idRaw, idPretty, type } */
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

/** Convenience wrappers */
function getPropertyNameIdPairsFromPage(pageIdOrUrl: string, opts: { titleFirst?: boolean; sort?: boolean } = {}) {
  const page = notionGetPage(pageIdOrUrl);
  return getPropertyNameIdPairs(page, opts);
}
function getPropertyNameIdPairsFromDataSource(dsIdOrUrl: string, opts: { titleFirst?: boolean; sort?: boolean } = {}) {
  const ds = notionGetDataSource(dsIdOrUrl);
  return getPropertyNameIdPairs(ds, opts);
}

/** Get an array of property IDs (raw or pretty) */
function getPropertyIds(obj: { properties?: any }, form: "raw" | "pretty" = "raw"): string[] {
  const props = obj?.properties || {};
  const ids: string[] = [];
  for (const p of Object.values(props) as any[]) {
    const raw = String(p?.id || "");
    ids.push(form === "pretty" ? decodeId(raw) : raw);
  }
  return ids;
}
function getPropertyIdsFromPage(pageIdOrUrl: string, form: "raw" | "pretty" = "raw"): string[] {
  const page = notionGetPage(pageIdOrUrl);
  return getPropertyIds(page, form);
}
function getPropertyIdsFromDataSource(dsIdOrUrl: string, form: "raw" | "pretty" = "raw"): string[] {
  const ds = notionGetDataSource(dsIdOrUrl);
  return getPropertyIds(ds, form);
}

/** Save/load arrays into Script Properties (IDs or names) */
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

/** One-shot cache helpers */
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

/** Print specific property id examples */
function printEmailOrgId(pageIdOrUrl: string): void {
  const page = notionGetPage(pageIdOrUrl);
  const raw = String(page.properties?.["Email (Org)"]?.id || "");
  const pretty = decodeId(raw);
  Logger.log(`Email (Org) id = ${pretty} (raw=${raw})`);
}

/** Simple page property id logger */
function printIdsForPage(pageIdOrUrl: string): void {
  const page = notionGetPage(pageIdOrUrl);
  logPropertyIds(page);
}

function looksLikeId(s: string) { 
  return /[%a-z0-9]{2,}/i.test(s); 
}

