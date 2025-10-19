/// <reference types="google-apps-script" />

/**
 * slackPermissions.ts — Slack scope & permissions helpers (GAS)
 * -------------------------------------------------------------
 * Requires:
 *  - slackApi.ts: slackApi(...)
 *
 * What it does:
 *  - slackAuthTest(): basic auth info for the bot token
 *  - slackGetGrantedScopes(): read the token's granted scopes (bot & user)
 *  - slackEnsureScopes(): assert scopes are present (throws helpful error)
 *  - Quick ensure helpers for common feature bundles (messaging, directory, etc.)
 *
 * Notes:
 *  Slack has multiple "list scopes" endpoints depending on app type & age.
 *  We try `apps.auth.scopes.list` first, then fall back to `apps.permissions.scopes.list`.
 */

/** Fetch bot auth info (scopes, team, app id) */
function slackAuthTest(): { ok: boolean; url?: string; team?: string; user?: string; bot_id?: string } {
    const r = slackApi<{ ok: boolean; url?: string; team?: string; user?: string; bot_id?: string }>("auth.test", {
        method: "GET",
    });
    // auth.test often returns {ok:false} with HTTP 200 for certain errors; normalize as-is.
    return {
        ok: r.ok && r.data.ok === true,
        url: r.data.url,
        team: r.data.team,
        user: r.data.user,
        bot_id: (r.data as any).bot_id,
    };
}

/**
 * Read the bot token’s granted scopes.
 * Returns { bot: string[], user?: string[] }
 *
 * We attempt in this order:
 *  1) apps.auth.scopes.list   → { scopes: { app_home?: string[], team?: string[], user?: string[] } }
 *  2) apps.permissions.scopes.list → similar structure for older / classic apps
 */
function slackGetGrantedScopes(): { bot: string[]; user?: string[] } {
    // Try modern endpoint
    let r = slackApi<{ ok: boolean; scopes?: Record<string, string[]> }>("apps.auth.scopes.list", {
        method: "GET",
    });

    if (!r.ok || !r.data || !(r.data as any).scopes) {
        // Fallback for older apps
        r = slackApi<{ ok: boolean; scopes?: Record<string, string[]> }>("apps.permissions.scopes.list", {
            method: "GET",
        });
    }

    const scopes = (r.data && (r.data as any).scopes) || {};
    // Slack may organize scopes by buckets (e.g., app_home, team, user). Treat any non-`user` array as bot app scopes.
    const user = Array.isArray(scopes.user) ? scopes.user.slice() : undefined;

    const botSets: string[][] = [];
    for (const [bucket, arr] of Object.entries(scopes)) {
        if (bucket === "user") continue;
        if (Array.isArray(arr)) botSets.push(arr);
    }
    const bot = Array.from(new Set(botSets.flat()));

    return { bot, user };
}

/**
 * Ensure required scopes are present; throw with a helpful message if not.
 * @param requiredBotScopes   e.g., ["chat:write", "files:write"]
 * @param requiredUserScopes  e.g., ["users.profile:write"] (optional)
 */
function slackEnsureScopes(requiredBotScopes: string[], requiredUserScopes?: string[]): void {
    const granted = slackGetGrantedScopes();
    const haveBot = new Set((granted.bot || []).map(String));
    const haveUser = new Set((granted.user || []).map(String));

    const missingBot = (requiredBotScopes || []).filter(s => !haveBot.has(s));
    const missingUser = (requiredUserScopes || []).filter(s => !haveUser.has(s));

    if (missingBot.length || missingUser.length) {
        const blocks: string[] = [];
        if (missingBot.length) blocks.push(`Missing **bot** scopes: ${missingBot.join(", ")}`);
        if (missingUser.length) blocks.push(`Missing **user** scopes: ${missingUser.join(", ")}`);

        const hint =
            "Reinstall your Slack app with the required scopes in its manifest, then re-authorize. " +
            "Docs: https://api.slack.com/authentication/quickstart";

        throw new Error(`Slack scopes check failed.\n${blocks.join("\n")}\n\n${hint}`);
    }
}

/** Quick checks for feature sets */

// e.g., posting messages, uploading files, listing channels
function slackEnsureMessagingScopes(): void {
    const bot = ["chat:write", "files:write", "channels:read", "groups:read"];
    slackEnsureScopes(bot);
}

// e.g., reading user directory (needed for email→userId lookups, profiles reads)
function slackEnsureDirectoryScopes(): void {
    const bot = ["users:read", "users:read.email"];
    slackEnsureScopes(bot);
}

// e.g., writing custom profile fields
function slackEnsureProfileWriteScopes(): void {
    const bot = ["users.profile:write"];
    slackEnsureScopes(bot);
}

// e.g., managing @usergroups (oncall rotations, etc.)
function slackEnsureUsergroupScopes(): void {
    const bot = ["usergroups:read", "usergroups:write"];
    slackEnsureScopes(bot);
}

// e.g., creating/archiving channels, inviting users
function slackEnsureChannelAdminScopes(): void {
    // `conversations.write` is broadly used for membership / admin ops on channels.
    // Some workspaces may still rely on legacy scopes like channels:manage / groups:write.
    const bot = ["conversations:write", "channels:manage", "groups:write"];
    slackEnsureScopes(bot);
}