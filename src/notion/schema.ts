/**
 * 05_schema.ts — Notion schema utilities (GAS)
 * ----------------------------------------------------------
 * - Case/spacing-insensitive property-name index
 * - Alias→spec builder (ID-first, then name)
 *
 * Requires:
 *   - utils/core.ts: normKey, decodeId, looksLikeId
 *   - notion/04_resources.ts: notionGetDataSource
 *
 * Exposes (global):
 *   - buildNameIndexCI
 *   - buildSpecsFromDataSourceWithAliases
 *   - buildSpecsFromDataSourceWithAliasesOnly
 *   - buildSpecsFromDataSourceWithAliasesOnlyCI
 */

/**
 * Build {label, propId} specs by joining your alias map with Notion’s live schema.
 * Includes all properties from the data source, even if alias not found.
 */
function buildSpecsFromDataSourceWithAliases(
  dsIdOrUrl: string,
  aliases: Record<string, string>
): Array<{ label: string; propId: string; name: string }> {
  const obj = notionGetDataSource(dsIdOrUrl);
  const props = obj?.properties || {};
  const specs: Array<{ label: string; propId: string; name: string }> = [];

  for (const [name, prop] of Object.entries<any>(props)) {
    const propIdRaw = String(prop?.id || "");
    const alias = aliases[name] ?? name; // alias if present, else original
    specs.push({ label: alias, propId: propIdRaw, name });
  }

  return specs;
}

/**
 * Build specs but only for aliases that exist in the given map.
 * Skips properties not present in Notion’s schema.
 */
function buildSpecsFromDataSourceWithAliasesOnly(
  dsIdOrUrl: string,
  aliases: Record<string, string>
): Array<{ label: string; propId: string; name: string }> {
  const obj = notionGetDataSource(dsIdOrUrl);
  const props = obj?.properties || {};
  const specs: Array<{ label: string; propId: string; name: string }> = [];

  for (const aliasName of Object.keys(aliases)) {
    const prop = (props as any)[aliasName];
    if (!prop || !prop.id) continue; // skip if Notion doesn’t have it
    specs.push({
      label: aliases[aliasName] || aliasName,
      propId: String(prop.id),
      name: aliasName,
    });
  }

  return specs;
}

/**
 * Build a case- and spacing-insensitive name→id index.
 * Converts property names like "Full Name" → "fullname".
 */
function buildNameIndexCI(
  props: Record<string, any>
): Map<string, { name: string; id: string }> {
  const idx = new Map<string, { name: string; id: string }>();

  for (const [name, prop] of Object.entries(props)) {
    const id = String((prop as any)?.id || "");
    if (!id) continue;
    idx.set(normKey(name), { name, id });
  }

  return idx;
}

/**
 * Aliases-only, case-insensitive match against Notion schema.
 * Matches both property IDs and property names (case/spacing-insensitive).
 */
function buildSpecsFromDataSourceWithAliasesOnlyCI(
  dsIdOrUrl: string,
  aliases: Record<string, string>
): NotionSpec[] {
  const obj = notionGetDataSource(dsIdOrUrl);
  const props = (obj as any)?.properties || {};
  const index = buildNameIndexCI(props);
  const specs: NotionSpec[] = [];

  for (const aliasKey of Object.keys(aliases)) {
    // 1️⃣ Match by property ID first (e.g., "HA%40l" or "HA@l")
    if (looksLikeId(aliasKey)) {
      const pretty = decodeId(aliasKey);
      const match = Object.entries<any>(props).find(([_, p]) => {
        const id = String(p?.id || "");
        return decodeId(id) === pretty || id === aliasKey;
      });
      if (match) {
        const [realName, p] = match;
        specs.push({
          label: aliases[aliasKey],
          propId: String(p.id),
          name: realName,
        });
      } else {
        Logger.log(`⚠️ No property with id "${aliasKey}"`);
      }
      continue;
    }

    // 2️⃣ Otherwise, match by NAME (case/spacing-insensitive)
    const hit = index.get(normKey(aliasKey));
    if (!hit) {
      Logger.log(`⚠️ Alias not found (CI): "${aliasKey}"`);
      continue;
    }

    specs.push({
      label: aliases[aliasKey] || hit.name,
      propId: hit.id,
      name: hit.name,
    });
  }

  return specs;
}