/// <reference types="google-apps-script" />

/**
 * 03_http.ts — Notion HTTP runtime (GAS)
 * ------------------------------------------------
 * Centralized Notion API calls using UrlFetchApp.
 * - Handles authentication, Notion-Version, and query building.
 * - Parses JSON automatically and throws on HTTP errors.
 * - Retries 429/5xx with exponential backoff.
 *
 * Requires (loaded earlier):
 * - utils/core.ts: getHeaderCI, safeJson
 *
 * Exposes (global):
 * - toGasMethod
 * - notionApi
 * - notionFetchWithRetry
 */

/* -------------------------------------------------------------------------- */
/*                               Method helpers                                */
/* -------------------------------------------------------------------------- */

/**
 * Normalize to GAS's lowercase HttpMethod.
 * Converts "GET" → "get", etc.
 */
function toGasMethod(
    m: AnyCaseHttpMethod = "GET"
): GoogleAppsScript.URL_Fetch.HttpMethod {
  return String(m).toLowerCase() as GoogleAppsScript.URL_Fetch.HttpMethod;
}

/* -------------------------------------------------------------------------- */
/*                              Core API wrapper                               */
/* -------------------------------------------------------------------------- */

/**
 * Core Notion API call wrapper for Google Apps Script.
 * Handles headers, authentication, JSON parsing, and error reporting.
 * Also retries on 429/5xx via notionFetchWithRetry.
 *
 * @template T Parsed data type expected from the response.
 * @param params Parameters for the API call.
 * @returns Wrapped response with status, headers, parsed data, and request info.
 * @throws Error if missing token, malformed path, or non-2xx when throwOnHttpError is true.
 */
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
    headers: extraHeaders = {}, // optional user-supplied headers (declared in your d.ts)
  } = params as NotionApiParams & { headers?: Record<string, string> };

  if (!path || !path.startsWith("/")) {
    throw new Error('notionApi: "path" must start with "/"');
  }
  if (!token) {
    throw new Error("notionApi: missing NOTION_TOKEN");
  }

  // Build URL + query string
  let url = NOTION_BASE + path;
  if (query) {
    const qs: string[] = [];
    for (const [k, v] of Object.entries(query)) {
      if (v == null) continue;
      if (Array.isArray(v)) {
        for (const it of v) qs.push(`${encodeURIComponent(k)}=${encodeURIComponent(String(it))}`);
      } else {
        qs.push(`${encodeURIComponent(k)}=${encodeURIComponent(String(v))}`);
      }
    }
    if (qs.length) url += "?" + qs.join("&");
  }

  // Headers (auth + version + any extras)
  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    "Notion-Version": version,
    ...extraHeaders,
  };

  // Options
  const gasMethod = toGasMethod(method);
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: gasMethod,
    headers,
    muteHttpExceptions: true,
  };

  // Body/payload: only for POST/PATCH by default
  const needsBody = gasMethod === "post" || gasMethod === "patch";
  if (needsBody && body !== undefined) {
    if (typeof body === "string") {
      options.payload = body as FetchPayload;
      options.contentType = headers["Content-Type"] || "application/json";
    } else if (typeof body === "object" && body && "getBytes" in (body as object)) {
      // Blob payload
      const blob = body as GoogleAppsScript.Base.Blob;
      options.payload = blob as FetchPayload;
      const blobType = blob.getContentType();
      if (blobType && !headers["Content-Type"]) options.contentType = blobType;
    } else {
      // JSON payload
      options.payload = JSON.stringify(body) as FetchPayload;
      options.contentType = "application/json";
    }
  }

  if (debug) {
    const safeHeaders = { ...headers, Authorization: "Bearer ***redacted***" };
    Logger.log(
        `[notionApi] → ${gasMethod.toUpperCase()} ${url}\n` +
        `headers=${JSON.stringify(safeHeaders)}\n` +
        `contentType=${options.contentType || "(none)"} hasPayload=${options.payload != null}`
    );
    if (needsBody && typeof options.payload === "string") {
      try { Logger.log(`[notionApi] payload=${options.payload.slice(0, 1000)}`); } catch {}
    }
  }

  // Fetch with retry
  const resp = notionFetchWithRetry(url, options);
  const status = resp.getResponseCode();
  const respHeaders = resp.getHeaders() as unknown as Record<string, string>;
  const ctype = respHeaders["Content-Type"] || respHeaders["content-type"] || "";
  const text = resp.getContentText();

  let data: unknown = text;
  if (ctype.includes("application/json")) {
    try {
      data = JSON.parse(text);
    } catch {
      // leave as raw text if JSON parse fails
    }
  }

  if (throwOnHttpError && (status < 200 || status >= 300)) {
    // Attempt to show a concise error body
    let snippet = text;
    try { snippet = typeof data === "object" ? safeJson(data) : text; } catch {}
    throw new Error(`notionApi: HTTP ${status} → ${snippet}`);
  }

  return {
    ok: status >= 200 && status < 300,
    status,
    data: data as T,
    headers: respHeaders,
    url,
    method: gasMethod,
  };
}

/* -------------------------------------------------------------------------- */
/*                               Retry wrapper                                 */
/* -------------------------------------------------------------------------- */

/**
 * Retry wrapper for UrlFetchApp.fetch() with exponential backoff.
 * Retries 429 (rate-limited) and 5xx server errors up to 5 times.
 *
 * - Obeys `Retry-After` header when present (seconds).
 * - Doubles delay between attempts when header is absent.
 *
 * @param url     The target URL.
 * @param options Fetch options.
 * @returns       The successful response.
 * @throws        Error after exhausting all retries.
 */
function notionFetchWithRetry(
    url: string,
    options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions
): GoogleAppsScript.URL_Fetch.HTTPResponse {
  const maxAttempts = 5;
  let delayMs = 250;
  let lastErr: any = null;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const resp = UrlFetchApp.fetch(url, options);
      const status = resp.getResponseCode();

      if (status === 429 || status >= 500) {
        const headers = resp.getHeaders() as unknown as Record<string, string>;
        const ra = Number(getHeaderCI(headers, "Retry-After") || 0);
        const wait = ra > 0 ? ra * 1000 : delayMs;
        Utilities.sleep(wait);
        delayMs *= 2;
        continue;
      }

      return resp; // success
    } catch (e) {
      lastErr = e;
      Utilities.sleep(delayMs);
      delayMs *= 2;
    }
  }

  throw lastErr || new Error("notionFetchWithRetry: failed after retries");
}