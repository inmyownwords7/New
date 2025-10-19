/// <reference types="google-apps-script" />

/**
 * slack_utils.ts — shared helpers for Slack modules (GAS)
 * ------------------------------------------------------
 * - Token & default channel resolution from Script Properties
 * - Channel/name normalization
 * - Safe JSON + log truncation
 * - Sleep alias for consistency
 *
 * Script Properties expected:
 *  - SLACK_BOT_TOKEN
 *  - (optional) SLACK_DEFAULT_CHANNEL
 */

/** Get Slack bot token from Script Properties, or throw if missing. */
function getSlackToken(): string {
    const token = PropertiesService.getScriptProperties().getProperty("SLACK_BOT_TOKEN");
    if (!token) {
        throw new Error(
            "Missing SLACK_BOT_TOKEN in Script Properties. " +
            "Set it in Apps Script → Project Settings → Script Properties."
        );
    }
    return token;
}

/** Get default channel (channel ID like Cxxxx / Gxxxx, or '#name'). Returns null if unset. */
function getSlackDefaultChannel(): string | null {
    const v = PropertiesService.getScriptProperties().getProperty("SLACK_DEFAULT_CHANNEL");
    return v ? String(v) : null;
}

/**
 * Normalize a channel reference.
 * - If it's an ID (C… / G…), return as-is.
 * - If it starts with '#', strip it (Slack API wants names without '#').
 * - Otherwise return the trimmed string.
 * @example normalizeChannelRef("#alerts") → "alerts"
 * @example normalizeChannelRef("C0123ABCD") → "C0123ABCD"
 */
function normalizeChannelRef(ref: string): string {
    const s = String(ref || "").trim();
    if (!s) throw new Error("normalizeChannelRef: empty channel reference");
    if (/^[CG][A-Z0-9]+$/i.test(s)) return s; // already an ID
    return s.replace(/^#/, "");
}

/** Safe JSON stringify for logs/debugging. */
function safeJsonSlack(x: unknown): string {
    try {
        return JSON.stringify(x);
    } catch {
        try { return String(x); } catch { return "[unstringifiable]"; }
    }
}

/**
 * Truncate long strings for logs to avoid quota noise.
 * @param s   Input string
 * @param max Default 1000 chars
 */
function truncateForLog(s: string, max: number = 1000): string {
    const str = String(s ?? "");
    return str.length > max ? str.slice(0, max) + "…" : str;
}

