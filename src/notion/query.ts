/// <reference types="google-apps-script" />

/**
 * 06_query.ts — Notion query pagination (GAS)
 * ------------------------------------------------
 * - queryDataSourceAll: fetches all results with cursor pagination
 *
 * Requires (loaded earlier as globals):
 *  - utils/core.ts: extractId32, normalizeUuid, safeJson
 *  - notion/http.ts: notionApi
 *
 * Exposes (global):
 *  - queryDataSourceAll
 */

/**
 * Query all pages from a Notion data source, following cursors until exhaustion.
 *
 * Accepts a raw 32-hex ID, dashed UUID, or a Notion URL containing the ID.
 * Uses POST /v1/data_sources/{id}/query with optional filters/sorts/pagination body.
 *
 * @param dsIdOrUrl Raw ID (32-hex), dashed UUID, or Notion URL containing the ID.
 * @param queryBody The POST body for /query (filters, sorts, sorts, etc.).
 * @param opts      Optional pagination/debug options: { pageSize?: number; debug?: boolean }.
 * @returns         All result objects concatenated into a single array.
 * @throws          Error on HTTP failure or missing id.
 */
function queryDataSourceAll(
    dsIdOrUrl: string,
    queryBody: Record<string, unknown> = {},
    opts: { pageSize?: number; debug?: boolean } = {}
): any[] {
  const { pageSize = 100, debug = false } = opts;

  const id = normalizeUuid(extractId32(dsIdOrUrl));
  if (!id) throw new Error("queryDataSourceAll: missing data source ID/URL");

  const base = { page_size: pageSize, ...queryBody };

  // First page
  let r = notionApi<{ results?: any[]; has_more?: boolean; next_cursor?: string }>({
    method: "POST",
    path: `/v1/data_sources/${id}/query`,
    body: base,
    throwOnHttpError: false, // we'll throw a cleaner message below
    debug,
  });
  if (!r.ok) {
    throw new Error(`POST /v1/data_sources/${id}/query → ${r.status} ${safeJson(r.data)}`);
  }

  const all = [...(r.data.results || [])];
  let cursor = r.data.has_more && r.data.next_cursor ? r.data.next_cursor : null;

  // Paginate
  while (cursor) {
    r = notionApi<{ results?: any[]; has_more?: boolean; next_cursor?: string }>({
      method: "POST",
      path: `/v1/data_sources/${id}/query`,
      body: { ...base, start_cursor: cursor },
      throwOnHttpError: false,
      debug,
    });

    if (!r.ok) {
      throw new Error(`pagination POST /v1/data_sources/${id}/query → ${r.status} ${safeJson(r.data)}`);
    }

    all.push(...(r.data.results || []));
    cursor = r.data.has_more && r.data.next_cursor ? r.data.next_cursor : null;
  }

  return all;
}