# 🧭 Notion → Google Sheets Sync (Apps Script + TypeScript)

This project syncs **Notion** data into **Google Sheets** using Google Apps Script compiled from **TypeScript**.  
It’s structured into modular files that load in a specific order — utilities first, then Notion logic, then Sheets helpers.

---

## 📁 Folder Structure

```text
src/
├─ utils/
│  ├─ 01_core.ts           # decodeId, normKey, safeJson, getHeaderCI, META_KEY_COLMAP, etc.
│  └─ 02_ids.ts            # extractId32, normalizeUuid
│
├─ notion/
│  ├─ 03_http.ts           # toGasMethod, notionApi, notionFetchWithRetry
│  ├─ 04_resources.ts      # isDataSource, isDatabase, notionGetDataSource, notionGetPage
│  ├─ 05_schema.ts         # buildNameIndexCI, buildSpecsFromDataSourceWithAliases*, etc.
│  ├─ 06_query.ts          # queryDataSourceAll
│  ├─ 07_props.ts          # titleOf, getPropById, property logging/caching
│  └─ 08_orchestrator.ts   # buildSpecsFromAliases, fetchAllPages, pagesToRows, stringify*, syncDataSourceToSheet
│
├─ sheets/
│  ├─ 01_access.ts         # getSheetId, resolveSpreadsheetId, getSheetByNameOrCreate, withSheetLock
│  ├─ 02_headers.ts        # ensureHeaders, formatSheet, setHeaderCellWithId, findColumnByPropId
│  ├─ 03_writes.ts         # appendRowsBatched, upsertRowsByKey, writePropIdsToSheet, extractCellValue
│  └─ 04_state.ts          # getSheetMeta, setSheetMeta, buildIdToNameMap, saveIdNameMap, loadIdNameMap
│
├─ app.ts                  # Demo entrypoints & test functions
└─ types/
   └─ global.d.ts          # All global type declarations
⚙️ Setup
Add Script Properties
In your Apps Script project → Project Settings → Script Properties:

Key Example Value
NOTION_TOKEN secret_xxx
NOTION_VERSION 2025-09-03
SPREADSHEET_ID Your Google Sheet ID

Authorize Scopes
First run will request access to:

spreadsheets

script.external_request

script.scriptapp

userinfo.email

Compile or Upload
Use npx tsc to transpile to /build or upload .ts files directly to Apps Script (TS supported natively).

🧩 File Roles
File Purpose
utils/ Base helpers & constants (decodeId, safeJson, etc.)
notion/ API, schema, property, and pagination logic
sheets/ Spreadsheet read/write, header formatting, and metadata state
app.ts Manual test runners for smoke testing and demos
types/global.d.ts Shared ambient type declarations for GAS runtime

🚀 First Run
1️⃣ Smoke Test
ts
Copy code
function notionApi_smoke() {
  const r = notionApi({ method: "GET", path: "/v1/users/me", debug: true });
  Logger.log(JSON.stringify({
    ok: r.ok,
    status: r.status,
    object: (r.data as any)?.object,
    name: (r.data as any)?.name
  }, null, 2));
}
Run this once — if it logs your Notion user, your token works!

2️⃣ Create Headers
ts
Copy code
function demo_headers() {
  const DS_ID = "YOUR_NOTION_DATA_SOURCE_ID";
  ensureAliasHeadersFromDataSourceWithMap(DS_ID, "People Sync", aliasHeaders);
}
or exact placement from col A:

ts
Copy code
function demo_headers3() {
  const DS_ID = "YOUR_NOTION_DATA_SOURCE_ID";
  ensureAliasHeadersExact(DS_ID, "People Sync", aliasHeaders, 1);
}
3️⃣ Sync Data
Append Mode:

ts
Copy code
function run_sync_append() {
  const DS_ID = "YOUR_NOTION_DATA_SOURCE_ID";
  syncDataSourceToSheet(DS_ID, aliasHeaders, "People Sync", { mode: "append" });
}
Upsert Mode:

ts
Copy code
function run_sync_upsert() {
  const DS_ID = "YOUR_NOTION_DATA_SOURCE_ID";
  syncDataSourceToSheet(DS_ID, aliasHeaders, "People Sync", {
    mode: "upsert",
    keyLabel: "Email",
    batchSize: 300
  });
}
🧠 How the Mapping Works
You define an alias map (e.g., "Email (Org)" → "Email").

The script loads the Notion schema, building specs:

ts
Copy code
{ label, propId, name }
Header cells get the label as visible text and store propId in the cell note.

The sheet also saves a JSON { column: propId } map in developer metadata (META_KEY_COLMAP).

During sync:

Each Notion page is flattened via stringifyNotionProp()

Writes are idempotent (ensureHeaders, upsertRowsByKey, etc.)

withSheetLock() ensures safe concurrency

🧰 Utilities
Rebuild Lost Metadata
ts
Copy code
rebuildHeaderMetadataFromNotes("People Sync", 2);
Cache an ID → Name Map
ts
Copy code
refreshIdNameMapFromDataSource("YOUR_DS_ID", "PEOPLE_ID2NAME");
🧩 Troubleshooting
Issue Fix
ReferenceError: META_KEY_COLMAP is not defined Define const META_KEY_COLMAP = "NOTION_SYNC_COLMAP"; in utils/01_core.ts.
Key header "X" not found Ensure keyLabel matches an existing header label.
No data or 401 errors Verify NOTION_TOKEN and integration access.
Header columns shifted Re-run ensureAliasHeadersFromDataSourceWithMap() and formatSheet().

🧱 Conventions
No imports/exports — Apps Script loads globals in filename order.

Idempotent — all sheet operations can re-run safely.

Locked writes — withSheetLock() prevents trigger collisions.

Readable logs — safeJson() truncates large objects cleanly.

🧾 License
MIT © 2025 — Use freely, but handle API tokens responsibly.

yaml
Copy code

---

Would you like me to generate a **Table of Contents (TOC)** block at the top (GitHub-style `[links](#sections
