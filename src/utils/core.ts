/// <reference types="google-apps-script" />

/** =========================
 * utils/core.ts (GAS globals)
 * Small, reusable helpers for Notion + Sheets
 * ========================= */

/** Case-insensitive header lookup from an object of headers. */
function getHeaderCI(headers: Record<string, string>, name: string): string | undefined {
  const n = String(name || "").toLowerCase();
  for (const k in headers) if (k && k.toLowerCase() === n) return headers[k];
  return undefined;
}

/** Normalize a string for case-insensitive comparisons / map keys. */
function normKey(s: string): string {
  return s ? s.normalize("NFKC").trim().replace(/\s+/g, " ").toLowerCase() : "";
}

/** Safe JSON stringify for logs/debug. */
function safeJson(x: unknown): string {
  try { return JSON.stringify(x); } catch { return String(x); }
}

/** Heuristic: looks like some kind of ID token (very loose). */
function looksLikeId(s: string): boolean {
  return /[%a-z0-9]{2,}/i.test(String(s || ""));
}

/** Stricter: looks like a Notion property id (percent-encoded or 32-hex). */
function looksLikePropId(s: string): boolean {
  const str = String(s || "");
  return /%[0-9A-Fa-f]{2}/.test(str) || /^[0-9A-Fa-f]{32}$/.test(str);
}

/** Decode a Notion property id like `HA%40l` → `HA@l` (no-throw). */
function decodeId(id: unknown): string {
  if (!id) return "";
  try { return decodeURIComponent(String(id)); } catch { return String(id); }
}

/** Safe percent-decode wrapper reused across helpers. */
function safeDecode(s: string): string {
  try { return decodeURIComponent(String(s)); } catch { return String(s); }
}

/** True if a value is null/undefined/empty after trim. */
function isBlank(v: unknown): boolean {
  return String(v ?? "").trim() === "";
}

/** Sleep N ms (GAS). */
function sleepMs(ms: number): void { Utilities.sleep(ms); }

/** Extract a 32-hex id from a string or return the input as string. */
function extractId32(input: unknown): string {
  if (!input) return "";
  const m = String(input).match(/[0-9a-f]{32}/i);
  return m ? m[0] : String(input);
}

/** Normalize 32-hex → UUID with dashes; otherwise return as-is. */
function normalizeUuid(id: string): string {
  const m = String(id).match(/^[0-9a-fA-F]{32}$/);
  if (!m) return id; // already hyphenated or not 32-hex
  const r = m[0].toLowerCase();
  return `${r.slice(0,8)}-${r.slice(8,12)}-${r.slice(12,16)}-${r.slice(16,20)}-${r.slice(20)}`;
}

/** Best-effort normalization of a Notion prop id (raw + decoded + UUID). */
function normalizePropId(input: unknown): { raw: string; decoded: string; dashed32: string } {
  const raw = String(input ?? "");
  const decoded = safeDecode(raw);
  const hex32 = extractId32(decoded);
  const dashed32 = /^[0-9a-fA-F]{32}$/.test(hex32) ? normalizeUuid(hex32) : "";
  return { raw, decoded, dashed32 };
}

/** Find a header index case-insensitively (returns −1 if not found). */
function findHeaderIndexCI(headers: string[], label: string): number {
  const needle = String(label ?? "").trim().toLowerCase();
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i] ?? "").trim().toLowerCase() === needle) return i;
  }
  return -1;
}

/** Ensure row has exactly `width` columns (pads/trims with empty strings). */
function padRowToWidth(row: any[], width: number): any[] {
  const out = row.slice(0, width);
  while (out.length < width) out.push("");
  return out;
}

/** Simple array chunker for batching. */
function chunkArray<T>(arr: T[], size: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

/**
 * Build a dual map for a Notion page/database-like object:
 *  - byName: name → property object (enriched with .name)
 *  - idToName: id (raw/decoded) → name
 *  - map: name → object AND id (raw/decoded) → name (dual-purpose)
 *
 * Helpers:
 *  - get(key): if key is a name → object, if id → name
 *  - getPropByName(name): property object
 *  - getNameById(id): property name
 */
function buildPropFlipMap(obj: { properties?: Record<string, any> }) {
  const props = obj?.properties || {};
  const byName = new Map<string, any>();
  const idToName = new Map<string, string>();
  const map = new Map<string, any>(); // name → object and id → name

  for (const [name, prop] of Object.entries(props)) {
    const enriched = { ...prop, name };
    byName.set(name, enriched);

    const raw = String(prop?.id ?? "");
    if (raw) {
      const dec = safeDecode(raw);
      idToName.set(raw, name);
      if (dec !== raw) idToName.set(dec, name);

      map.set(name, enriched);  // NAME → OBJECT
      map.set(raw, name);       // RAW ID → NAME
      if (dec !== raw) map.set(dec, name); // DECODED → NAME
    }
  }

  const get = (key: string) => map.get(key) ?? null;
  const getPropByName = (n: string) => byName.get(n) ?? null;
  const getNameById = (id: string) => idToName.get(id) ?? null;

  return { get, getPropByName, getNameById, byName, idToName, map };
}