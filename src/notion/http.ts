/**
 * 03_http.ts — Notion HTTP runtime (GAS)
 * ------------------------------------------------
 * Centralized Notion API calls using UrlFetchApp.
 * - Handles authentication, Notion-Version, and query building.
 * - Parses JSON automatically and throws on HTTP errors.
 * - Includes exponential backoff for 429 (rate limit) and 5xx retries.
 *
 * Requires (loaded earlier):
 *   - utils/core.ts: getHeaderCI, safeJson
 *
 * Exposes (global):
 *   - toGasMethod
 *   - notionApi
 *   - notionFetchWithRetry
 */

/**
 * Normalize to GAS's lowercase HttpMethod.
 * Converts "GET" → "get", etc.
 */
function toGasMethod(
  m: AnyCaseHttpMethod = "GET"
): GoogleAppsScript.URL_Fetch.HttpMethod {
  return String(m).toLowerCase() as GoogleAppsScript.URL_Fetch.HttpMethod;
}

/**
 * Core Notion API call wrapper for Google Apps Script.
 * Handles headers, authentication, JSON parsing, and error reporting.
 *
 * @param {NotionApiParams} params - Parameters for the API call.
 * @returns {NotionApiResult<T>} - Wrapped response with status, headers, and parsed data.
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
  } = params;

  if (!path || !path.startsWith("/")) throw new Error('notionApi: "path" must start with "/"');
  if (!token) throw new Error("notionApi: missing NOTION_TOKEN");

  // Build URL with query params
  let url = NOTION_BASE + path;
  if (query) {
    const qs: string[] = [];
    for (const [k, v] of Object.entries(query)) {
      if (v == null) continue;
      if (Array.isArray(v)) {
        for (const it of v)
          qs.push(`${encodeURIComponent(k)}=${encodeURIComponent(String(it))}`);
      } else {
        qs.push(`${encodeURIComponent(k)}=${encodeURIComponent(String(v))}`);
      }
    }
    if (qs.length) url += "?" + qs.join("&");
  }

  // Headers
  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    "Notion-Version": version,
  };

  // Options
  const gasMethod = toGasMethod(method);
  const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
    method: gasMethod,
    headers,
    muteHttpExceptions: true,
  };

  // Body/payload
  const needsBody = gasMethod === "post" || gasMethod === "patch";
  if (needsBody && body !== undefined) {
    if (typeof body === "string") {
      options.payload = body as FetchPayload;
      options.contentType = headers["Content-Type"] || "application/json";
    } else if (typeof body === "object" && "getBytes" in (body as object)) {
      const blob = body as GoogleAppsScript.Base.Blob;
      options.payload = blob as FetchPayload;
      const blobType = blob.getContentType();
      if (blobType && !headers["Content-Type"]) options.contentType = blobType;
    } else {
      options.payload = JSON.stringify(body) as FetchPayload;
      options.contentType = "application/json";
    }
  }

  if (debug) {
    const safe = { ...headers, Authorization: "Bearer ***redacted***" };
    Logger.log(
      `[notionApi] → ${gasMethod.toUpperCase()} ${url}\nheaders=${JSON.stringify(safe)}\n` +
        `contentType=${options.contentType || "(none)"}  hasPayload=${options.payload != null}`
    );
  }

  const resp = UrlFetchApp.fetch(url, options);
  const status = resp.getResponseCode();
  const respHeaders = resp.getHeaders() as unknown as Record<string, string>;
  const ctype = respHeaders["Content-Type"] || "";
  const text = resp.getContentText();

  let data: unknown = text;
  if (ctype.includes("application/json")) {
    try {
      data = JSON.parse(text);
    } catch {}
  }

  if (throwOnHttpError && (status < 200 || status >= 300)) {
    throw new Error(`notionApi: HTTP ${status} → ${text}`);
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

/**
 * Retry wrapper for UrlFetchApp.fetch() with exponential backoff.
 * Retries 429 (rate-limited) and 5xx server errors up to 5 times.
 *
 * @param {string} url - The target URL.
 * @param {GoogleAppsScript.URL_Fetch.URLFetchRequestOptions} options - Fetch options.
 * @returns {GoogleAppsScript.URL_Fetch.HTTPResponse} - The successful response.
 * @throws {Error} After exhausting all retries.
 */
function notionFetchWithRetry(
  url: string,
  options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions
) {
  const maxAttempts = 5;
  let delayMs = 250;
  let lastErr: any = null;

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const resp = UrlFetchApp.fetch(url, options);
      const status = resp.getResponseCode();
      const headers = resp.getHeaders() as unknown as Record<string, string>;

      if (status === 429 || status >= 500) {
        const ra = Number(getHeaderCI(headers, "Retry-After") || 0);
        const wait = ra > 0 ? ra * 1000 : delayMs;
        Utilities.sleep(wait);
        delayMs *= 2;
        continue;
      }

      return resp;
    } catch (e) {
      lastErr = e;
      Utilities.sleep(delayMs);
      delayMs *= 2;
    }
  }

  throw lastErr || new Error("notionFetchWithRetry: failed after retries");
}