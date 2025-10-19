/// <reference types="google-apps-script" />

/**
 * slackProfiles.ts — Slack user/profile helpers (GAS)
 * ---------------------------------------------------
 * Requires:
 *  - slackApi.ts: slackApi(...)
 * (optionally) slackPermissions.ts: slackEnsureDirectoryScopes, slackEnsureProfileWriteScopes
 *
 * Typical scopes:
 *  - users:read, users:read.email
 *  - users.profile:read (optional for reading full custom fields)
 *  - users.profile:write (for setting custom fields)
 */

/** Look up a user by email → return minimal identity (id, name…). */
function slackFindUserByEmail(email: string): { id: string; team_id?: string; email?: string; name?: string } | null {
    const e = String(email || "").trim();
    if (!e) throw new Error("slackFindUserByEmail: email is empty");


    const r = slackApi<{
        ok: boolean;
        user?: { id: string; team_id?: string; profile?: { email?: string; real_name?: string; display_name?: string; real_name_normalized?: string; display_name_normalized?: string } };
    }>("users.lookupByEmail", {
        method: "GET",
        query: { email: e },
    });

    if (!r.ok || !r.data.user) return null;

    const u = r.data.user;
    const prof = u.profile || {};
    const name =
        prof.display_name_normalized ||
        prof.display_name ||
        prof.real_name_normalized ||
        prof.real_name ||
        undefined;

    return { id: u.id, team_id: u.team_id, email: prof.email || e, name };
}

/** Get a user's full profile (raw Slack structure). */
function slackGetUserProfile(userId: string): any {
    const uid = String(userId || "").trim();
    if (!uid) throw new Error("slackGetUserProfile: userId is empty");

    const r = slackApi<{ ok: boolean; profile?: any }>("users.profile.get", {
        method: "GET",
        query: { user: uid },
    });
    if (!r.ok) throw new Error(`users.profile.get failed: ${r.status} ${JSON.stringify(r.data)}`);
    return (r.data as any).profile || {};
}

/**
 * Set custom profile fields (internal keys, e.g., "XfABC123").
 * NOTE: Slack expects a JSON string in 'profile' with shape:
 *   { "fields": { "XfABC123": { "value": "foo" }, ... } }
 */
function slackSetProfileFields(userId: string, fields: Record<string, string>): void {
    const uid = String(userId || "").trim();
    if (!uid) throw new Error("slackSetProfileFields: userId is empty");
    if (!fields || Object.keys(fields).length === 0) return;

    // If you wired permissions helpers, uncomment:
    // try { slackEnsureProfileWriteScopes(); } catch {}

    // Build profile.fields map
    const profile: any = { fields: {} as Record<string, { value: string }> };
    for (const [k, v] of Object.entries(fields)) {
        if (!k) continue;
        profile.fields[k] = { value: String(v ?? "") };
    }

    // Slack prefers x-www-form-urlencoded with `profile` JSON.
    const r = slackApi<{ ok: boolean }>("users.profile.set", {
        method: "POST",
        asForm: true,
        body: { user: uid, profile: JSON.stringify(profile) },
    });

    if (!r.ok) throw new Error(`users.profile.set failed: ${r.status} ${JSON.stringify(r.data)}`);
}

/**
 * Batch update by email — convenient for HR/Directory sync.
 * - Looks up each email → userId
 * - Writes provided custom field keys (Xf…)
 * - Returns a small summary object
 */
function slackBatchUpdateProfiles(
    updates: Array<{ email: string; fields: Record<string, string> }>
): { updated: number; skipped: number; errors: number } {
    let updated = 0, skipped = 0, errors = 0;

    // If you wired permissions helpers, uncomment:
    // try { slackEnsureDirectoryScopes(); slackEnsureProfileWriteScopes(); } catch {}

    for (const item of updates || []) {
        const email = String(item?.email || "").trim();
        const fields = item?.fields || {};
        if (!email || Object.keys(fields).length === 0) { skipped++; continue; }

        try {
            const user = slackFindUserByEmail(email);
            if (!user || !user.id) { skipped++; continue; }

            slackSetProfileFields(user.id, fields);
            updated++;

            // Gentle pacing to avoid noisy 429s on large batches
            Utilities.sleep(300);
        } catch (e) {
            errors++;
            // optional: Logger.log(`slackBatchUpdateProfiles error for ${email}: ${e}`);
            Utilities.sleep(300);
        }
    }

    return { updated, skipped, errors };
}