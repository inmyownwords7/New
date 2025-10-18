/// <reference types="google-apps-script" />

let aliasHeaders: Record<string, string> = {
  "Email (Org)": "Email",
  "Name (Org)": "Name",
  "Last edited time": "Edited",
  "Grey-Box id": "id",
  "Mandate (Status)": "Mandate",
  "Position (Current)": "Position",
  "Team (Current)": "Team",
  "Mandate (Date)": "MandateDate",
  "Hours (Initial)": "Hours",
  "Hours (Current)": "HoursCurrent",
  "Created Profile (Date)": "CreatedProfile",
  "Error Detection": "Error",
  "Notion Page URL": "NotionURL",
};

/** Notion API smoke test */
function notionApi_smoke(): NotionApiResult<any> {
  const r = notionApi({ method: "GET", path: "/v1/users/me", debug: true });
  Logger.log(JSON.stringify({ ok: r.ok, status: r.status, object: (r.data as any)?.object, name: (r.data as any)?.name }, null, 2));
  return r;
}

/** Sample query printer (first 10 rows) */
function test_query_all_print(): void {
  const dsId = "a92f493a-6843-4b0d-9812-117d699055db"; // <- replace with yours
  const rows = queryDataSourceAll(
    dsId,
    { sorts: [{ timestamp: "last_edited_time", direction: "ascending" }] },
    { pageSize: 50, debug: true }
  );
  Logger.log(`total=${rows.length}`);
  rows.slice(0, 10).forEach((pg: any, i: number) => Logger.log(`${i + 1}. ${pg.id} — ${titleOf(pg)}`));
}

// Your headers still start at B1:
function demo_headers3() {
  const DS_ID = "a92f493a-6843-4b0d-9812-117d699055db";
  ensureAliasHeadersExact(DS_ID, "People Sync", aliasHeaders, /*startCol=*/1);
}

function demo_headers() {
  const DS_ID = "a92f493a-6843-4b0d-9812-117d699055db";
  ensureAliasHeadersFromDataSourceWithMap(DS_ID, "People Sync", aliasHeaders);
}