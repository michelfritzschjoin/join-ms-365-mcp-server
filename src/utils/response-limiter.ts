/**
 * Response limiter: optionally truncate tool response text to stay within context limits.
 * Set MS365_MCP_MAX_RESPONSE_CHARS > 0 to enable (default 0 = no limit).
 */

const TRUNCATION_SUFFIX = '\n\n[... Antwort gekürzt; max. Zeichen erreicht.]';

/**
 * If MS365_MCP_MAX_RESPONSE_CHARS is set and text exceeds it, truncate and append a marker.
 * Otherwise return text unchanged.
 */
export function applyResponseLimit(text: string): string {
  if (typeof text !== 'string') return text;
  const maxChars = parseInt(process.env.MS365_MCP_MAX_RESPONSE_CHARS || '0', 10);
  if (maxChars <= 0) return text;
  if (text.length <= maxChars) return text;
  return text.slice(0, maxChars - TRUNCATION_SUFFIX.length) + TRUNCATION_SUFFIX;
}
