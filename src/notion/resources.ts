/**
 * 04_resources.ts — Notion resource fetchers (GAS)
 * ------------------------------------------------
 * - Get Data Source / Database by ID (with database fallback)
 * - Get Page by ID
 * - Runtime type guards for response objects
 *
 * Requires:
 *   - utils/core.ts: extractId32, normalizeUuid, safeJson
 *   - notion/http.ts: notionApi
 *
 * Exposes (global):
 *   - isDataSource
 *   - isDatabase
 *   - notionGetDataSource
 *   - notionGetPage
 */

/** Type guard: Notion Data Source */
function isDataSource(x: unknown): x is NotionDataSource {
  return !!x && typeof x === "object" && (x as any).object === "data_source";
}

/** Type guard: Notion Database */
function isDatabase(x: unknown): x is NotionDatabase {
  return !!x && typeof x === "object" && (x as any).object === "database";
}

/**
 * Fetch a Notion Data Source by ID; if not found, fall back to Database by the same ID.
 * Accepts raw IDs or URLs; normalizes to dashed UUID form.
 */
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

/**
 * Fetch a Notion Page by ID or URL.
 * Normalizes input to dashed UUID before calling the API.
 */
function notionGetPage(idOrUrl: string): NotionPage {
  const id = normalizeUuid(extractId32(idOrUrl));
  if (!id) throw new Error("notionGetPage: missing ID/URL");

  const r = notionApi<NotionPage | { object?: "error" }>({
    method: "GET",
    path: `/v1/pages/${id}`,
    throwOnHttpError: false
  });

  if (r.ok && (r.data as NotionPage).object === "page") return r.data as NotionPage;
  throw new Error(`GET /v1/pages/${id} → ${r.status} ${safeJson(r.data)}`);
}