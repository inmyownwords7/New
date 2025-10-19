/// <reference types="google-apps-script" />

/**
 * slackApi.ts — Core Slack Web API wrapper for GAS
 * ------------------------------------------------
 * - slackFetchWithRetry: low-level fetch with 429/5xx backoff
 * - slackApi: auth, headers, payload shaping, error normalization
 * - slackPostMessage: chat.postMessage convenience
 * - slackUploadFile: files.upload for text or blob
 * - postSyncSummaryToSlack: tiny notifier for sync jobs
 *
 * Script Properties expected (from slack_utils.ts):
 *  - SLACK_BOT_TOKEN  → via getSlackToken()
 *  - SLACK_DEFAULT_CHANNEL → via getSlackDefaultChannel()
 */

/** Case-insensitive header getter (Slack-local version). */
function slackGetHeaderCI(headers: Record<string, string>, name: string): string | undefined {
    const n = name.toLowerCase();
    for (const k in headers) if (k && k.toLowerCase() === n) return headers[k];
    return undefined;
}

/** Safe JSON stringify for Slack logs. */
function slackSafeJson(x: unknown): string {
    try { return JSON.stringify(x); } catch { return String(x); }
}

/** Low-level fetch with backoff (429 respects Retry-After). */
function slackFetchWithRetry(
    url: string,
    options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions
): GoogleAppsScript.URL_Fetch.HTTPResponse {
    const maxAttempts = 5;
    let delayMs = 300;
    let lastErr: any = null;

    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
        try {
            const resp = UrlFetchApp.fetch(url, options);
            const status = resp.getResponseCode();
            if (status === 429 || status >= 500) {
                const headers = resp.getHeaders() as unknown as Record<string, string>;
                const ra = Number(slackGetHeaderCI(headers, "Retry-After") || 0);
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
    throw lastErr || new Error("slackFetchWithRetry: failed after retries");
}

/** Core API wrapper (auth, payload shaping, error normalization). */
type SlackMethod = "GET" | "POST";
type SlackResult<T = unknown> = { ok: boolean; status: number; data: T; url: string; method: SlackMethod };

function slackApi<T = any>(path: string, opts: {
    method?: SlackMethod;
    token?: string;
    query?: Record<string, string | number | boolean | undefined>;
    body?: Record<string, any> | string | GoogleAppsScript.Base.Blob;
    asForm?: boolean;
    asMultipart?: boolean;
    debug?: boolean;
} = {}): SlackResult<T> {
    const base = "https://slack.com/api";
    const {
        method = "POST",
        token = getSlackToken(), // from slack_utils.ts
        query,
        body,
        asForm = false,
        asMultipart = false,
        debug = false,
    } = opts;

    let url = `${base}/${path.replace(/^\//, "")}`;
    if (query) {
        const qs = Object.entries(query)
            .filter(([, v]) => v !== undefined && v !== null && v !== "")
            .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(String(v))}`)
            .join("&");
        if (qs) url += (url.includes("?") ? "&" : "?") + qs;
    }

    const headers: Record<string, string> = { Authorization: `Bearer ${token}` };
    const options: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {
        method: method.toLowerCase() as GoogleAppsScript.URL_Fetch.HttpMethod,
        headers,
        muteHttpExceptions: true,
    };

    // Handle payloads
    if (method === "POST" && body !== undefined) {
        if (typeof body === "string") {
            options.payload = body;
            options.contentType = "application/json";
        } else if (asForm) {
            const form = Object.entries(body)
                .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(String(v))}`)
                .join("&");
            options.payload = form;
            options.contentType = "application/x-www-form-urlencoded";
        } else {
            options.payload = JSON.stringify(body);
            options.contentType = "application/json";
        }
    }

    if (debug) {
        Logger.log(`[slackApi] → ${method} ${url}`);
        if (options.payload) Logger.log(`Payload: ${typeof options.payload === "string" ? options.payload.slice(0, 400) : "(binary)"}`);
    }

    const resp = slackFetchWithRetry(url, options);
    const status = resp.getResponseCode();
    const text = resp.getContentText();
    let data: any = {};
    try { data = JSON.parse(text); } catch { data = { ok: status >= 200 && status < 300, text }; }

    const ok = data?.ok === true && status >= 200 && status < 300;
    if (!ok && debug) Logger.log(`[slackApi] !ok status=${status} body=${text}`);
    return { ok, status, data: data as T, url, method };
}

/** Post a message easily (no hard dependency on slackChannels.ts). */
function slackPostMessage(args: {
    channel?: string;
    text: string;
    blocks?: any[];
    thread_ts?: string;
    unfurl_links?: boolean;
    unfurl_media?: boolean;
    debug?: boolean;
}): { ok: boolean; ts?: string; channel?: string } {
    const defChan = getSlackDefaultChannel() || ""; // from slack_utils.ts
    const chanIn = args.channel || defChan;
    if (!chanIn) throw new Error("slackPostMessage: no channel provided and SLACK_DEFAULT_CHANNEL missing");

    // Resolve channel without requiring slackEnsureChannelId
    const normalized = normalizeChannelRef(chanIn); // from slack_utils.ts
    const channel = /^[CG][A-Z0-9]+$/i.test(normalized)
        ? normalized
        : (typeof slackFindChannelId === "function" ? (slackFindChannelId(normalized) || chanIn) : chanIn);

    const payload = {
        channel,
        text: args.text || "",
        blocks: args.blocks,
        thread_ts: args.thread_ts,
        unfurl_links: args.unfurl_links,
        unfurl_media: args.unfurl_media,
    };

    const r = slackApi<{ ok: boolean; ts?: string; channel?: string }>("chat.postMessage", {
        method: "POST",
        body: payload,
        debug: !!args.debug,
    });

    if (!r.ok) throw new Error(`chat.postMessage failed → ${r.status} ${slackSafeJson(r.data)}`);
    return r.data;
}

/** Upload files (text or Blob) (no hard dependency on slackChannels.ts). */
function slackUploadFile(args: {
    channels?: string;
    channel?: string;
    filename: string;
    content?: string;
    blob?: GoogleAppsScript.Base.Blob;
    filetype?: string;
    initial_comment?: string;
    thread_ts?: string;
    debug?: boolean;
}): any {
    const defChan = getSlackDefaultChannel() || "";
    let channels = args.channels || args.channel || defChan;
    if (!channels) throw new Error("slackUploadFile: no channels provided");

    // If single channel and it's not an ID, try to resolve name → ID (if finder is present)
    if (!channels.includes(",") && !/^[CG][A-Z0-9]+$/i.test(channels)) {
        const normalized = normalizeChannelRef(channels);
        channels = (typeof slackFindChannelId === "function" ? (slackFindChannelId(normalized) || channels) : channels);
    }

    const body: Record<string, any> = {
        channels,
        filename: args.filename,
        filetype: args.filetype,
        initial_comment: args.initial_comment,
        thread_ts: args.thread_ts,
    };

    if (args.blob) {
        body["file"] = args.blob.setName(args.filename);
        return slackApi("files.upload", { method: "POST", body, asMultipart: true, debug: !!args.debug }).data;
    } else if (args.content) {
        body["content"] = args.content;
        return slackApi("files.upload", { method: "POST", body, asForm: true, debug: !!args.debug }).data;
    } else {
        throw new Error("slackUploadFile: provide either content or blob");
    }
}

/** Post a summary message to Slack after a Notion→Sheets sync. */
function postSyncSummaryToSlack(run: {
    mode: "append" | "upsert";
    sheet: string;
    rows?: number;
    inserted?: number;
    updated?: number;
    channel?: string;
}): void {
    const text =
        run.mode === "upsert"
            ? `✅ Notion → Sheets upsert complete on *${run.sheet}*: ${run.inserted || 0} inserted, ${run.updated || 0} updated.`
            : `✅ Notion → Sheets append complete on *${run.sheet}*: ${run.rows || 0} rows appended.`;

    slackPostMessage({ channel: run.channel, text });
}