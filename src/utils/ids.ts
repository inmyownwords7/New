// Extract a 32-hex id from a string or return the input as string
function extractId32(input: unknown): string {
  if (!input) return "";
  const m = String(input).match(/[0-9a-f]{32}/i);
  return m ? m[0] : String(input);
}

// Normalize 32-hex → UUID with dashes; otherwise return as-is
function normalizeUuid(id: string): string {
  const m = String(id).match(/^[0-9a-fA-F]{32}$/);
  if (!m) return id; // already hyphenated or not 32-hex
  const r = m[0].toLowerCase();
  return `${r.slice(0,8)}-${r.slice(8,12)}-${r.slice(12,16)}-${r.slice(16,20)}-${r.slice(20)}`;
}