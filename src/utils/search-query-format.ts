/**
 * Intelligent search query formatting for Microsoft Graph KQL (e.g. email).
 * Detects person names and common intents (from/to) to build optimal $search values.
 * Handles composite OR queries so the entire expression is one quoted KQL string (per Microsoft Graph docs).
 */

const PROPERTY_VALUE_PATTERN = /^([a-zA-Z]+):(.+)$/i;
/** Person-like: letters, umlauts, spaces, dots, hyphens; 1–4 words */
const PERSON_NAME_PATTERN = /^[\wäöüßÄÖÜ\s.-]+$/;
/** "an Maria Müller" / "to John" → recipient intent */
const TO_INTENT_PATTERN = /^(an|to)\s+(.+)$/i;
/** Composite OR for mail: "from:X OR from:Y" or "to:A OR from:B" */
const COMPOSITE_OR_PATTERN = /\s+OR\s+/i;
/** Extract email from text (simple RFC-style) */
const EMAIL_IN_QUERY_PATTERN = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/;

/**
 * Returns true if the value looks like a person name (no KQL prefix, 1–4 words, letters/umlauts).
 */
export function looksLikePersonName(value: string): boolean {
  const trimmed = value.trim();
  if (!trimmed) return false;
  const wordCount = trimmed.split(/\s+/).length;
  return wordCount >= 1 && wordCount <= 4 && PERSON_NAME_PATTERN.test(trimmed);
}

/**
 * Returns true if the value looks like a composite mail KQL query (e.g. from:X OR from:Y).
 * Such queries must be normalized to one quoted string and not parsed as a single property:value.
 */
export function isCompositeMailSearchQuery(value: string): boolean {
  const trimmed = value.trim();
  if (!trimmed || !COMPOSITE_OR_PATTERN.test(trimmed)) return false;
  const parts = trimmed.split(COMPOSITE_OR_PATTERN);
  return parts.length >= 2 && parts.every((p) => PROPERTY_VALUE_PATTERN.test(p.trim()));
}

/**
 * Normalizes a composite mail search (e.g. from:"Name" OR from:email) to a single KQL-quoted string.
 * Microsoft Graph expects the entire $search value as one double-quoted string: "from:A OR from:B".
 * @param rawValue - Raw composite query (may contain inner quotes)
 * @returns One quoted KQL string for $search
 */
export function normalizeCompositeEmailSearchQuery(rawValue: string): string {
  const trimmed = rawValue.trim();
  const parts = trimmed.split(COMPOSITE_OR_PATTERN).map((p) => p.trim());
  const clauses: string[] = [];
  for (const part of parts) {
    const m = part.match(PROPERTY_VALUE_PATTERN);
    if (!m) continue;
    const prop = m[1];
    let val = m[2].trim();
    if (val.startsWith('"') && val.endsWith('"')) val = val.slice(1, -1);
    if (val.startsWith('\\"') && val.endsWith('\\"')) val = val.slice(2, -2);
    clauses.push(`${prop}:${val}`);
  }
  if (clauses.length === 0) return '""';
  const inner = clauses.join(' OR ');
  return `"${inner.replace(/"/g, '\\"')}"`;
}

/**
 * Extracts the first email address from a search string (e.g. for fallback from:email-only query).
 */
export function extractEmailFromSearch(search: string): string | null {
  if (!search || typeof search !== 'string') return null;
  const match = search.match(EMAIL_IN_QUERY_PATTERN);
  return match ? match[0] : null;
}

/**
 * Formats a raw search value for /me/messages $search (KQL).
 * - If value already has a property prefix (from:, to:, subject:, etc.), preserves it and quotes.
 * - If value starts with "an " or "to ", uses to: (e.g. "E-Mails an Maria" → from: / to:).
 * - If value looks like a person name, uses from: (e.g. "Maria Müller" → "from:Maria Müller").
 * - Otherwise returns plain quoted value (searches from, subject, body).
 * @param rawValue - Raw $search value (may include surrounding quotes)
 * @returns Quoted KQL string for $search parameter
 */
export function formatEmailSearchQuery(rawValue: string): string {
  if (!rawValue || typeof rawValue !== 'string') return '""';
  const quotePattern = /^["']?(.*?)["']?$/s;
  const trimmed = rawValue.replace(quotePattern, '$1').trim();
  if (!trimmed) return '""';

  // Composite OR (e.g. from:"Name" OR from:email) → one quoted KQL string; do not parse as single property
  if (isCompositeMailSearchQuery(trimmed)) {
    return normalizeCompositeEmailSearchQuery(trimmed);
  }

  // Already has property prefix → ensure proper quoting
  const propertyMatch = trimmed.match(PROPERTY_VALUE_PATTERN);
  if (propertyMatch) {
    const property = propertyMatch[1];
    let value = propertyMatch[2].trim();
    if (value.startsWith('"') && value.endsWith('"')) value = value.slice(1, -1);
    if (value.includes(' ')) {
      const inner = `${property}:"${value}"`;
      return `"${inner.replace(/"/g, '\\"')}"`;
    }
    return `"${property}:${value}"`;
  }

  // "an X" / "to X" → recipient ("E-Mails an Maria Müller" / "emails to John")
  const toMatch = trimmed.match(TO_INTENT_PATTERN);
  if (toMatch) {
    const recipient = toMatch[2].trim();
    return `"to:${recipient}"`;
  }

  // Person-like → sender ("letzte E-Mail von Maria Müller" etc.)
  if (looksLikePersonName(trimmed)) {
    return `"from:${trimmed}"`;
  }

  // Plain text search
  return `"${trimmed.replace(/"/g, '\\"')}"`;
}
