/// <reference types="google-apps-script" />

/**
 * app.ts — Entrypoint / UI layer for Notion → Sheets sync
 * Requires globals from:
 * - notion/orchestrator.ts
 * - sheets/access.ts: resolveSpreadsheetId, getSheetByNameOrCreate, withSheetLock
 * - sheets/headers.ts: ensureHeaders
 * - sheets/writes.ts: clearDataBelowHeader, appendRowsBatched, upsertRowsByKey
 */

// ---- config you customize ----
const DS_PEOPLE = "a92f493a-6843-4b0d-9812-117d699055db"; // Notion Data Source ID (replace)
const TAB_PEOPLE = "People Sync";                           // Target sheet tab name

// Aliases map: Notion property name → Sheet header label
const aliasHeaders: Record<string, string> = {
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

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu("Notion Sync")
      .addItem("Preview first 10", "app_previewFirst10")
      .addItem("Fix headers (keep data)", "app_fixHeaders")
      .addItem("Wipe & rebuild", "app_wipeAndRebuild")
      .addItem("Append all rows", "app_syncAppendAll")
      .addItem("Upsert by NotionURL", "app_syncUpsertByNotionURL")
      .addToUi();
}

/** Logs first 10 rows (no sheet writes) */
function app_previewFirst10() {
  const specs = buildSpecsFromAliases(DS_PEOPLE, aliasHeaders);
  const headers = makeHeadersFromSpecs(specs);
  const pages = fetchAllPages(DS_PEOPLE, { page_size: 50 });
  const rows = pagesToRows(pages, specs).slice(0, 10);

  Logger.log(`Headers: ${headers.join(" | ")}`);
  rows.forEach((r, i) => Logger.log(`${i + 1}. ${r.join(" | ")}`));
}

/** Ensure header labels exist/are correct (keeps data) */
function app_fixHeaders() {
  const specs = buildSpecsFromAliases(DS_PEOPLE, aliasHeaders);
  const headers = makeHeadersFromSpecs(specs);

  const ss = SpreadsheetApp.getActive();
  const sheet = getSheetByNameOrCreate(ss, TAB_PEOPLE);

  withSheetLock(() => {
    ensureHeaders(sheet, headers);
  });

  Logger.log(`Headers ensured on "${TAB_PEOPLE}" (${headers.length} cols).`);
}

/** Upsert by a unique key (must be a label in headers, e.g. NotionURL) */
function app_syncUpsertByNotionURL() {
  const KEY = "NotionURL";

  const specs = buildSpecsFromAliases(DS_PEOPLE, aliasHeaders);
  const headers = makeHeadersFromSpecs(specs);
  const pages = fetchAllPages(DS_PEOPLE);
  const rows = pagesToRows(pages, specs);

  const ss = SpreadsheetApp.getActive();
  const sheet = getSheetByNameOrCreate(ss, TAB_PEOPLE);

  withSheetLock(() => {
    ensureHeaders(sheet, headers);
    const res = upsertRowsByKey(sheet, KEY, headers, rows);
    Logger.log(`Upsert complete → inserted=${res.inserted}, updated=${res.updated}`);
  });
}

/** Full wipe + rebuild headers + append all */
function app_wipeAndRebuild() {
  const specs = buildSpecsFromAliases(DS_PEOPLE, aliasHeaders);
  const headers = makeHeadersFromSpecs(specs);
  const pages = fetchAllPages(DS_PEOPLE);
  const rows = pagesToRows(pages, specs);

  const ss = SpreadsheetApp.getActive();
  const sheet = getSheetByNameOrCreate(ss, TAB_PEOPLE);

  withSheetLock(() => {
    // hard reset data below header then write
    ensureHeaders(sheet, headers);
    clearDataBelowHeader(sheet);
    appendRowsBatched(sheet, rows, 500);
  });

  Logger.log(`Rebuilt + appended ${rows.length} rows on "${TAB_PEOPLE}".`);
}
