/// <reference types="google-apps-script" />

/**
 * slackOrchestrator.ts — High-level flows for Slack (GAS)
 * -------------------------------------------------------
 * Requires (previously loaded as globals):
 *  - slack_utils.ts: getSlackDefaultChannel, normalizeChannelRef, safeJsonSlack
 *  - slackApi.ts: slackPostMessage, slackUploadFile
 *  - slackChannels.ts: slackEnsureChannelId
 *  - slackUsergroups.ts: slackListUsergroups, slackGetUsergroupByHandle, slackUpdateUsergroupMembers
 *  - slackProfiles.ts: slackFindUserByEmail, slackSetProfileFields, slackBatchUpdateProfiles
 *  - slackPermissions.ts (optional): slackEnsureMessagingScopes, slackEnsureDirectoryScopes, slackEnsureUsergroupScopes, slackEnsureProfileWriteScopes
 *
 * What you get:
 *  - postSlackSyncSummary: one-liner to post append/upsert results
 *  - slackNotifyError: send a readable error to Slack
 *  - ensureChannelAndPost / ensureChannelAndUpload: convenience wrappers
 *  - syncUsergroupMembershipFromEmails: resolve emails to user IDs and replace membership
 *  - setProfileFieldsByEmail: set custom profile fields for one user by email
 *  - batchSetProfileFieldsByEmail: batch profile updates from { email, fields }[]
 *  - exportSheetAsCsvToSlack: turn a tab into CSV and upload it
 */

/** One-liner: post Notion→Sheets sync results to Slack. */
function postSlackSyncSummary(run: {
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

/**
 * Post a human-friendly error report to Slack (blocks + fallback text).
 * Good for try/catch wrappers in scheduled jobs.
 */
function slackNotifyError(args: {
    title?: string;
    error: unknown;
    context?: Record<string, unknown>;
    channel?: string;
    thread_ts?: string;
}) {
    try {
        // slackEnsureMessagingScopes?.(); // if you wired permissions helpers
    } catch {}

    const title = args.title || "❌ Task failed";
    const errText =
        typeof args.error === "string"
            ? args.error
            : args.error instanceof Error
                ? `${args.error.message}\n${(args.error.stack || "").split("\n").slice(0, 6).join("\n")}`
                : safeJsonSlack(args.error);

    const ctx = args.context ? "```" + safeJsonSlack(args.context) + "```" : "";

    const blocks = [
        { type: "header", text: { type: "plain_text", text: title } },
        {
            type: "section",
            text: {
                type: "mrkdwn",
                text: `*Error*\n${"```"}${errText}${"```"}${ctx ? `\n*Context*\n${ctx}` : ""}`,
            },
        },
    ];

    slackPostMessage({
        channel: args.channel,
        text: `${title}: ${errText}`,
        blocks,
        thread_ts: args.thread_ts,
        unfurl_links: false,
        unfurl_media: false,
    });
}

/** Ensure a channel (name or ID) is resolvable and post a message. */
function ensureChannelAndPost(args: {
    channel: string; // "#alerts" | "alerts" | "C…"
    text: string;
    blocks?: any[];
    thread_ts?: string;
    unfurl_links?: boolean;
    unfurl_media?: boolean;
}): { ok: boolean; ts?: string; channel?: string } {
    try {
        // slackEnsureMessagingScopes?.();
    } catch {}
    const id = slackEnsureChannelId(args.channel);
    return slackPostMessage({ ...args, channel: id });
}

/** Ensure channel and upload content/blob as a file. */
function ensureChannelAndUpload(args: {
    channel: string; // "#ops" | "C…"
    filename: string;
    content?: string;
    blob?: GoogleAppsScript.Base.Blob;
    filetype?: string;
    initial_comment?: string;
    thread_ts?: string;
}): any {
    try {
        // slackEnsureMessagingScopes?.();
    } catch {}
    const id = slackEnsureChannelId(args.channel);
    return slackUploadFile({ ...args, channel: id });
}

/**
 * Resolve "@handle" or "S…"-id, map emails → user IDs, then replace group membership.
 * Returns a small summary for logs/Slack posts.
 */
function syncUsergroupMembershipFromEmails(args: {
    usergroup: string; // "@handle" (e.g., "oncall") or "S…"
    emails: string[];
}): { usergroupId: string; resolved: number; missing: string[] } {
    try {
        // slackEnsureDirectoryScopes?.();
        // slackEnsureUsergroupScopes?.();
    } catch {}

    const ugInput = String(args.usergroup || "").replace(/^@/, "");
    let usergroupId = ugInput;
    if (!/^S[0-9A-Z]+$/i.test(ugInput)) {
        const ug = slackGetUsergroupByHandle(ugInput);
        if (!ug) throw new Error(`Usergroup handle "${ugInput}" not found`);
        usergroupId = ug.id;
    }

    const missing: string[] = [];
    const userIds: string[] = [];
    for (const email of args.emails || []) {
        if (!email) continue;
        const hit = slackFindUserByEmail(email);
        if (hit?.id) userIds.push(hit.id);
        else missing.push(email);
        Utilities.sleep(120); // gentle rate-limit spacing
    }

    slackUpdateUsergroupMembers(usergroupId, userIds);
    return { usergroupId, resolved: userIds.length, missing };
}

/**
 * Set custom profile fields for a single user, addressed by email.
 * `fields` must use internal field keys (Xf…).
 */
function setProfileFieldsByEmail(email: string, fields: Record<string, string>): { updated: boolean; userId?: string } {
    try {
        // slackEnsureDirectoryScopes?.();
        // slackEnsureProfileWriteScopes?.();
    } catch {}

    const u = slackFindUserByEmail(email);
    if (!u?.id) return { updated: false };
    slackSetProfileFields(u.id, fields);
    return { updated: true, userId: u.id };
}

/** Batch update custom profile fields by email. Returns summary counts. */
function batchSetProfileFieldsByEmail(updates: Array<{ email: string; fields: Record<string, string> }>): {
    updated: number;
    skipped: number;
    errors: number;
} {
    try {
        // slackEnsureDirectoryScopes?.();
        // slackEnsureProfileWriteScopes?.();
    } catch {}
    return slackBatchUpdateProfiles(updates);
}

/**
 * Export a sheet as CSV (header + rows) and upload to Slack.
 * Handy for lightweight reporting without Drive share links.
 */
function exportSheetAsCsvToSlack(args: {
    sheetName: string;
    channel?: string;
    filename?: string; // defaults to "<sheetName>.csv"
    initial_comment?: string;
    thread_ts?: string;
}): any {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(args.sheetName);
    if (!sheet) throw new Error(`exportSheetAsCsvToSlack: sheet "${args.sheetName}" not found`);

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 1 || lastCol < 1) throw new Error(`exportSheetAsCsvToSlack: sheet "${args.sheetName}" is empty`);

    const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const csv = values
        .map(row =>
            row
                .map(cell => {
                    const s = String(cell ?? "");
                    // CSV-escape: wrap if contains comma/quote/newline; escape quotes
                    if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
                    return s;
                })
                .join(",")
        )
        .join("\n");

    const filename = args.filename || `${args.sheetName}.csv`;
    const blob = Utilities.newBlob(csv, "text/csv", filename);

    return slackUploadFile({
        channel: args.channel || getSlackDefaultChannel() || "",
        filename,
        blob,
        filetype: "csv",
        initial_comment: args.initial_comment,
        thread_ts: args.thread_ts,
    });
}

/* ---------------------------
 * Demo helpers (optional)
 * ---------------------------
 * You can bind these to a custom menu or run manually to test your Slack wiring.
 */

function demo_slack_post_hello() {
    ensureChannelAndPost({ channel: getSlackDefaultChannel() || "#general", text: "👋 Hello from Apps Script!" });
}

function demo_slack_export_people_tab() {
    exportSheetAsCsvToSlack({ sheetName: "People Sync", initial_comment: "📤 Latest People export" });
}

function demo_slack_sync_oncall_group() {
    const emails = ["alice@example.com", "bob@example.com"];
    const res = syncUsergroupMembershipFromEmails({ usergroup: "oncall", emails });
    slackPostMessage({
        text: `@oncall updated. Members set: ${res.resolved}. Missing: ${res.missing.length ? res.missing.join(", ") : "none"}.`,
    });
}