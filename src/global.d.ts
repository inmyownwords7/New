// src/types/global.d.ts
/// <reference types="google-apps-script" />

declare function getSheetByNameOrCreate(
  ss: GoogleAppsScript.Spreadsheet.Spreadsheet,
  sheetName: string
): GoogleAppsScript.Spreadsheet.Sheet;

declare function withSheetLock<T>(fn: () => T): T;

declare function ensureHeaders(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  headers: string[]
): { changed: boolean };

declare function appendRowsBatched(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  rows: any[][],
  batchSize?: number
): void;

declare function upsertRowsByKey(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  keyHeader: string,
  headers: string[],
  rows: any[][]
): { inserted: number; updated: number };


/// <reference types="google-apps-script" />
type DevMeta = GoogleAppsScript.Spreadsheet.DeveloperMetadata;
/** Narrow methods you accept at your API surface (uppercase for callers) */
type NotionHttpMethodUpper = "GET" | "POST" | "PATCH" | "DELETE";
// If you didn’t already:
type Sheet = GoogleAppsScript.Spreadsheet.Sheet;
/** GAS fetch payload type (avoids DOM Blob conflicts) */
type FetchPayload = GoogleAppsScript.URL_Fetch.URLFetchRequestOptions["payload"];

/// <reference types="google-apps-script" />

type NotionSpec = {
  /** What you want to show as the column header (alias or original) */
  label: string;
  /** The Notion property ID (raw, may include % encodes) */
  propId: string;
  /** The actual Notion property name from the schema (optional but handy) */
  name?: string;
};

/** Query values that can be stringified */
type QueryValue =
  | string | number | boolean | null | undefined
  | (string | number | boolean)[];

  // Accept either uppercase or GAS’s lowercase HttpMethod
type AnyCaseHttpMethod =
  | NotionHttpMethodUpper
  | GoogleAppsScript.URL_Fetch.HttpMethod; // "get" | "post" | "patch" | "put" | "delete" | "head" | "options"
/** Wrapper params */
interface NotionApiParams {
  method?: NotionHttpMethodUpper | GoogleAppsScript.URL_Fetch.HttpMethod;
  path: string;                               // must start with "/"
  query?: Record<string, QueryValue>;
  body?: string | Record<string, unknown> | GoogleAppsScript.Base.Blob;
  token?: string;
  version?: string;
  throwOnHttpError?: boolean;
  debug?: boolean;
}

/** Wrapper result */
interface NotionApiResult<T = unknown> {
  ok: boolean;
  status: number;
  data: T;
  headers: Record<string, string>;
  url: string;
  method: GoogleAppsScript.URL_Fetch.HttpMethod;
}

/** Minimal Notion shapes you care about */
interface NotionWithObject { object: string }

interface NotionDataSource extends NotionWithObject {
  object: "data_source";
  id?: string;
  properties?: Record<string, unknown>;
}

interface NotionDatabase extends NotionWithObject {
  object: "database";
  id?: string;
  properties?: Record<string, unknown>;
}

interface NotionPage extends NotionWithObject {
  object: "page";
  id: string;
  properties?: Record<string, any>;
}

/** Type guards (type-only declarations if you prefer to implement elsewhere) */
// If you *only* want the types here and implement in .ts, use `declare function`:
declare function isDataSource(x: unknown): x is NotionDataSource;
declare function isDatabase(x: unknown): x is NotionDatabase;

/** If you reference helpers defined elsewhere, declare them here so TS knows them. */
declare function extractId32(input: unknown): string;
declare function normalizeUuid(id: string): string;
/// <reference types="google-apps-script" />

/* ========= Sheets: access, headers, formatting, metadata ========= */

declare const META_KEY_COLMAP: string;

declare function getSheetId(): string | null;

declare function formatSheet(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  headers: string[]
): void;

declare function resolveSpreadsheetId(): string;

declare function getOrCreateSheetByName(
  sheetName: string,
  headers: string[]
): GoogleAppsScript.Spreadsheet.Sheet;

declare function setHeaderCellWithId(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  row: number,
  col: number,
  label: string,
  propIdRaw: string
): void;

declare function ensureHeadersByPropId(
  sheetName: string,
  specs: Array<{ label: string; propId: string }>,
  startCol?: number
): { count: number; startCol: number };

declare function ensureHeadersExactByPropId(
  sheetName: string,
  specs: Array<{ label: string; propId: string; name: string }>,
  startCol?: number
): { count: number; startCol: number };

declare function findColumnByPropId(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  propIdRaw: string,
  startCol?: number,
  width?: number
): number | null;

/* optional fast-path if you added it */
declare function findColumnByPropIdFromMeta(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  propIdRaw: string
): number | null;

declare function getSheetMeta(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  key: string
): string | null;

declare function setSheetMeta(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  key: string,
  value: string
): void;

declare function rebuildHeaderMetadataFromNotes(
  sheetName: string,
  startCol?: number
): number;

declare function writeContiguousRow(
  sheetName: string,
  row: number,
  startCol: number,
  values: string[]
): void;

declare function writePropIdsToSheet(
  obj: { properties?: any },
  sheetName?: string
): number;

declare function writePagePropIdsToSheet(
  pageIdOrUrl: string,
  sheetName?: string
): number;

declare function writeDataSourcePropIdsToSheet(
  dsIdOrUrl: string,
  sheetName?: string
): number;

/* ========= Caching / ID↔name maps ========= */

declare function buildIdToNameMap(
  obj: { properties?: any }
): Map<string, string>;

declare function saveIdNameMap(
  key: string,
  map: Map<string, string>
): void;

declare function loadIdNameMap(key: string): Map<string, string>;

/* ========= Notion schema & orchestrators ========= */

declare function decodeId(id: unknown): string;

declare function stringifyNotionProp(prop: any): any;

declare function getPropById(
  page: any,
  propId: string,
  idNameMap: Map<string, string>
): any;

declare function buildSpecsFromDataSourceWithAliases(
  dsIdOrUrl: string,
  aliases: Record<string, string>
): Array<{ label: string; propId: string; name: string }>;

declare function buildSpecsFromDataSourceWithAliasesOnly(
  dsIdOrUrl: string,
  aliases: Record<string, string>
): Array<{ label: string; propId: string; name: string }>;

declare function buildNameIndexCI(
  props: Record<string, any>
): Map<string, { name: string; id: string }>;

declare function buildSpecsFromDataSourceWithAliasesOnlyCI(
  dsIdOrUrl: string,
  aliases: Record<string, string>
): Array<{ label: string; propId: string; name: string }>;

/* High-level helpers you pasted in sync.ts / orchestrator */

declare function ensureAliasHeadersExact(
  dsIdOrUrl: string,
  sheetName: string,
  aliases: Record<string, string>,
  startCol?: number
): { count: number; startCol: number };

declare function ensureAliasHeadersFromDataSourceWithMap(
  dsIdOrUrl: string,
  sheetName: string,
  aliases: Record<string, string>
): { count: number; startCol: number };

declare function verifyAliasCoverage(
  dsIdOrUrl: string,
  aliases: Record<string, string>
): void;

declare function extractCellValue(prop: any): string;

declare function writePageRowFast(
  pageIdOrUrl: string,
  dsIdOrUrl: string,
  sheetName?: string
): void;

declare function refreshIdNameMapFromDataSource(
  dsIdOrUrl: string,
  storeKey?: string
): number;
// ===== Notion HTTP/runtime you call from app.ts =====
declare function toGasMethod(m?: AnyCaseHttpMethod): GoogleAppsScript.URL_Fetch.HttpMethod;
declare function notionApi<T = unknown>(params: NotionApiParams): NotionApiResult<T>;
declare function notionFetchWithRetry(
  url: string,
  options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions
): GoogleAppsScript.HTTPResponse;

// ===== Notion resources & queries =====
declare function notionGetDataSource(idOrUrl: string): NotionDataSource | NotionDatabase;
declare function notionGetPage(idOrUrl: string): NotionPage;
declare function queryDataSourceAll(
  dsIdOrUrl: string,
  queryBody?: Record<string, unknown>,
  opts?: { pageSize?: number; debug?: boolean }
): any[];

// ===== Notion property helpers used in app.ts/tests =====
declare function titleOf(page: { properties?: any }): string;

// (optional diagnostics, if you call them anywhere)
declare function logPropertyIds(obj: { properties?: any }): void;
declare function logPropertyIdsFromPage(pageIdOrUrl: string): void;
declare function logPropertyIdsFromDataSource(dsIdOrUrl: string): void;

// ===== Orchestrators (both notions/orchestrator.ts and sheets/orchestrator.ts) =====
declare function buildSpecsFromAliases(dsIdOrUrl: string, aliases: Record<string, string>): NotionSpec[];
declare function fetchAllPages(dsIdOrUrl: string, queryBody?: Record<string, unknown>): any[];
declare function makeHeadersFromSpecs(specs: NotionSpec[]): string[];
declare function pagesToRows(pages: any[], specs: NotionSpec[]): any[][];
declare function syncDataSourceToSheet(
  dsIdOrUrl: string,
  aliases: Record<string, string>,
  sheetName: string,
  opts?: { mode?: "append" | "upsert"; keyLabel?: string; batchSize?: number }
): void;