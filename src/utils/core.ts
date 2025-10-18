/** Notion API (runtime) */
function getHeaderCI(headers: Record<string, string>, name: string): string | undefined {
  const n = name.toLowerCase();
  for (const k in headers) if (k && k.toLowerCase() === n) return headers[k];
  return undefined;
}


function normKey(s: string): string {
  return s
    ? s.normalize("NFKC").trim().replace(/\s+/g, " ").toLowerCase()
    : "";
}

function safeJson(x: unknown): string {
  try { return JSON.stringify(x); } catch { return String(x); }
}

function looksLikeId(s: string) { 
  return /[%a-z0-9]{2,}/i.test(s); 
}

/*** ─────────────────────────────────────────────
 * Extras you had before (utilities & diagnostics)
 * Paste these BELOW your current code
 * ──────────────────────────────────────────── ***/

/** decode a Notion property id like 'HA%40l' → 'HA@l' */
function decodeId(id: unknown): string {
  if (!id) return "";
  try { return decodeURIComponent(String(id)); } catch { return String(id); }
}

function sleepMs(ms: number) { Utilities.sleep(ms); }