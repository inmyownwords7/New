/**
 * 06_query.ts — Notion query pagination (GAS)
 * ------------------------------------------------
 * - queryDataSourceAll: fetches all results with cursor pagination
 *
 * Requires:
 *   - utils/core.ts: extractId32, normalizeUuid, safeJson
 *   - notion/http.ts: notionApi
 *
 * Exposes:
 *   - queryDataSourceAll
 */

/**
 * Query all pages from a Notion data source, following cursors until exhaustion.
 * @param dsIdOrUrl Raw ID (32-hex), dashed UUID, or a Notion URL containing the ID.
 * @param queryBody The POST body for /query (filters, sorts, etc.)
 * @param opts Optional pagination/debug options.
 * @returns All result objects concatenated into a single array.
 */
function queryDataSourceAll(
  dsIdOrUrl: string,
  queryBody: Record<string, unknown> = {},
  opts: { pageSize?: number; debug?: boolean } = {}
) {
  const { pageSize = 100, debug = false } = opts;

  const id = normalizeUuid(extractId32(dsIdOrUrl));
  if (!id) throw new Error("queryDataSourceAll: missing data source ID/URL");

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