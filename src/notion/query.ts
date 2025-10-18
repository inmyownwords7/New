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