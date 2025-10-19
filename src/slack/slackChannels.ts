/// <reference types="google-apps-script" />

/**
 * slackChannels.ts — Slack channel helpers (GAS)
 * ------------------------------------------------
 * Requires:
 *  - slackApi.ts: slackApi(...)
 *
 * Scopes typically needed:
 *  - channels:read / groups:read (to list/find)
 *  - channels:manage or conversations:write (to create/archive)
 *  - channels:manage / groups:write (to invite)
 */

/** Normalize "#name" → "name"; IDs passthrough. */
function slackNormalizeChannelRef(ref: string): string {
    if (!ref) throw new Error("slackNormalizeChannelRef: empty channel reference");
    if (/^[CG][A-Z0-9]+$/i.test(ref)) return ref; // already an ID
    return ref.replace(/^#/, "");
}

/** Resolve text like "#alerts" or "alerts" to a channel ID (public channels only). */
function slackEnsureChannelId(channelOrName: string): string {
    const inRef = slackNormalizeChannelRef(channelOrName);
    if (/^[CG][A-Z0-9]+$/i.test(inRef)) return inRef; // already an ID

    const hit = slackFindChannelByName(inRef);
    return hit ? hit.id : channelOrName; // return input if not found (Slack will error on misuse)
}

/**
 * Lookup a public channel by name (exact match).
 * @returns { id, name } or null if not found.
 */
function slackFindChannelByName(name: string): { id: string; name: string } | null {
    const norm = slackNormalizeChannelRef(name);

    let cursor: string | undefined = undefined;
    for (let i = 0; i < 20; i++) {
        const r = slackApi<{
            ok: boolean;
            channels?: Array<{ id: string; name: string }>;
            response_metadata?: { next_cursor?: string };
        }>("conversations.list", {
            method: "GET",
            query: { exclude_archived: true, limit: 1000, cursor },
        });

        if (!r.ok) break;
        const hit = (r.data.channels || []).find(ch => ch.name === norm);
        if (hit) return { id: hit.id, name: hit.name };

        cursor = r.data.response_metadata?.next_cursor || "";
        if (!cursor) break;
    }
    return null;
}

/** Return only the channel ID for a public channel name (or null). */
function slackFindChannelId(name: string): string | null {
    const ch = slackFindChannelByName(name);
    return ch ? ch.id : null;
}

/**
 * List public channels with pagination.
 * @param opts.exclude_archived default true
 * @param opts.limit            page size per API call (Slack permits up to 1000)
 */
function slackListChannels(
    opts: { exclude_archived?: boolean; limit?: number } = {}
): Array<{ id: string; name: string }> {
    const { exclude_archived = true, limit = 1000 } = opts;
    const out: Array<{ id: string; name: string }> = [];

    let cursor: string | undefined = undefined;
    for (let i = 0; i < 50; i++) {
        const r = slackApi<{
            ok: boolean;
            channels?: Array<{ id: string; name: string }>;
            response_metadata?: { next_cursor?: string };
        }>("conversations.list", {
            method: "GET",
            query: { exclude_archived, limit, cursor },
        });

        if (!r.ok) break;
        (r.data.channels || []).forEach(ch => out.push({ id: ch.id, name: ch.name }));

        cursor = r.data.response_metadata?.next_cursor || "";
        if (!cursor) break;
    }
    return out;
}

/**
 * Create a channel (public by default).
 * @param name       channel name (no '#')
 * @param isPrivate  true → create a private channel
 */
function slackCreateChannel(name: string, isPrivate: boolean = false): { id: string; name: string } {
    const norm = slackNormalizeChannelRef(name);
    const r = slackApi<{ ok: boolean; channel?: { id: string; name: string } }>("conversations.create", {
        method: "POST",
        body: { name: norm, is_private: !!isPrivate },
    });
    if (!r.ok || !r.data.channel) {
        throw new Error(`conversations.create failed: ${r.status} ${JSON.stringify(r.data)}`);
    }
    return { id: r.data.channel.id, name: r.data.channel.name };
}

/**
 * Invite a user to a channel.
 * @param channelId  C… or G… ID
 * @param userId     U… ID
 */
function slackInviteUserToChannel(channelId: string, userId: string): void {
    const r = slackApi<{ ok: boolean }>("conversations.invite", {
        method: "POST",
        body: { channel: channelId, users: userId },
    });
    if (!r.ok) throw new Error(`conversations.invite failed: ${r.status} ${JSON.stringify(r.data)}`);
}

/** Archive a channel (requires manage permissions). */
function slackArchiveChannel(channelId: string): void {
    const r = slackApi<{ ok: boolean }>("conversations.archive", {
        method: "POST",
        body: { channel: channelId },
    });
    if (!r.ok) throw new Error(`conversations.archive failed: ${r.status} ${JSON.stringify(r.data)}`);
}

/** Unarchive a channel (requires manage permissions). */
function slackUnarchiveChannel(channelId: string): void {
    const r = slackApi<{ ok: boolean }>("conversations.unarchive", {
        method: "POST",
        body: { channel: channelId },
    });
    if (!r.ok) throw new Error(`conversations.unarchive failed: ${r.status} ${JSON.stringify(r.data)}`);
}