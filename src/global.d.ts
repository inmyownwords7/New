// src/types/global.d.ts

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
