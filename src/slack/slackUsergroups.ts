/// <reference types="google-apps-script" />

/**
 * slackUsergroups.ts — Slack user group helpers (GAS)
 * ---------------------------------------------------
 * Requires:
 *  - slackApi.ts: slackApi(...)
 *
 * Typical scopes:
 *  - usergroups:read
 *  - usergroups:write
 *  - (optional) channels:read — only if you plan to resolve channel *names* to IDs yourself
 */

/** List user groups (optionally include disabled). */
function slackListUsergroups(
    include_disabled?: boolean
): Array<{ id: string; handle: string; name: string; is_usergroup: boolean }> {
    const r = slackApi<{
        ok: boolean;
        usergroups?: Array<{ id: string; handle: string; name: string; is_usergroup: boolean; date_delete?: number }>;
    }>("usergroups.list", {
        method: "GET",
        query: { include_disabled: !!include_disabled },
    });

    if (!r.ok) throw new Error(`usergroups.list failed: ${r.status} ${JSON.stringify(r.data)}`);

    const arr = r.data.usergroups || [];
    return arr.map(ug => ({
        id: ug.id,
        handle: ug.handle,
        name: ug.name,
        is_usergroup: ug.is_usergroup,
    }));
}

/** Get a user group by its @handle (e.g., "oncall"). */
function slackGetUsergroupByHandle(handle: string): { id: string; handle: string; name: string } | null {
    const list = slackListUsergroups(true);
    const hit = list.find(u => u.handle === handle);
    return hit ? { id: hit.id, handle: hit.handle, name: hit.name } : null;
}

/**
 * Update user group metadata.
 * Note: when passing `channels`, provide channel IDs (C…/G…); Slack expects a comma-separated list.
 */
function slackUpdateUsergroup(
    usergroupId: string,
    fields: { name?: string; handle?: string; description?: string; channels?: string[] }
): void {
    const body: Record<string, any> = { usergroup: usergroupId };
    if (fields.name != null) body.name = fields.name;
    if (fields.handle != null) body.handle = fields.handle;
    if (fields.description != null) body.description = fields.description;
    if (fields.channels && fields.channels.length) body.channels = fields.channels.join(",");

    const r = slackApi<{ ok: boolean }>("usergroups.update", {
        method: "POST",
        body,
    });
    if (!r.ok) throw new Error(`usergroups.update failed: ${r.status} ${JSON.stringify(r.data)}`);
}

/**
 * Replace membership for a user group with the given set of user IDs.
 * @param usergroupId S… user group ID
 * @param userIds     Array of U… user IDs
 */
function slackUpdateUsergroupMembers(usergroupId: string, userIds: string[]): void {
    const r = slackApi<{ ok: boolean }>("usergroups.users.update", {
        method: "POST",
        body: { usergroup: usergroupId, users: (userIds || []).join(",") },
    });
    if (!r.ok) throw new Error(`usergroups.users.update failed: ${r.status} ${JSON.stringify(r.data)}`);
}

/** Disable (deactivate) a user group. */
function slackDeactivateUsergroup(usergroupId: string): void {
    const r = slackApi<{ ok: boolean }>("usergroups.disable", {
        method: "POST",
        body: { usergroup: usergroupId },
    });
    if (!r.ok) throw new Error(`usergroups.disable failed: ${r.status} ${JSON.stringify(r.data)}`);
}

/** Enable (reactivate) a user group. */
function slackActivateUsergroup(usergroupId: string): void {
    const r = slackApi<{ ok: boolean }>("usergroups.enable", {
        method: "POST",
        body: { usergroup: usergroupId },
    });
    if (!r.ok) throw new Error(`usergroups.enable failed: ${r.status} ${JSON.stringify(r.data)}`);
}