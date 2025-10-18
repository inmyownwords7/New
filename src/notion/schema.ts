
function buildNameIndexCI(props: Record<string, any>): Map<string, { name: string; id: string }> {
  const idx = new Map<string, { name: string; id: string }>();
  for (const [name, prop] of Object.entries(props)) {
    const id = String((prop as any)?.id || "");
    if (!id) continue;
    idx.set(normKey(name), { name, id });
  }
  return idx;
}

/** Aliases-only, case-insensitive match against Notion schema */
function buildSpecsFromDataSourceWithAliasesOnlyCI(
  dsIdOrUrl: string,
  aliases: Record<string, string>
): NotionSpec[] {
  const obj = notionGetDataSource(dsIdOrUrl);
  const props = (obj as any)?.properties || {};
  const index = buildNameIndexCI(props);

  const specs: NotionSpec[] = [];
  for (const aliasKey of Object.keys(aliases)) {
    // 1) If alias key looks like an ID, match by ID
    if (looksLikeId(aliasKey)) {
      const pretty = decodeId(aliasKey);
      const match = Object.entries<any>(props).find(([_, p]) => {
        const id = String(p?.id || "");
        return decodeId(id) === pretty || id === aliasKey;
      });
      if (match) {
        const [realName, p] = match;
        specs.push({ label: aliases[aliasKey], propId: String(p.id), name: realName });
      } else {
        Logger.log(`⚠️ No property with id "${aliasKey}"`);
      }
      continue;
    }

    // 2) Otherwise match by NAME (case/spacing-insensitive)
    const hit = index.get(normKey(aliasKey));
    if (!hit) {
      Logger.log(`⚠️ Alias not found (CI): "${aliasKey}"`);
      continue;
    }
    specs.push({ label: aliases[aliasKey] || hit.name, propId: hit.id, name: hit.name });
  }
  return specs;
}
