/**
 * Performance configuration for Graph requests and limits.
 * Centralizes FAST_MODE and env-based timeouts, retries, and result limits
 * so responses can complete within a few seconds when desired.
 */

/** Default Graph request timeout (ms) when not set via env. */
const DEFAULT_GRAPH_REQUEST_TIMEOUT_MS = 15000;

/** Graph request timeout (ms) in fast mode. */
const FAST_MODE_GRAPH_REQUEST_TIMEOUT_MS = 6000;

/** Default max pagination pages when not set via env. */
const DEFAULT_MAX_PAGES = 20;

/** Max pagination pages in fast mode. */
const FAST_MODE_MAX_PAGES = 5;

/** Default max results per query when not set via env. */
const DEFAULT_MAX_RESULTS = 500;

/** Max results in fast mode. */
const FAST_MODE_MAX_RESULTS = 50;

/** Default max aggregate items when not set via env. */
const DEFAULT_MAX_AGGREGATE_ITEMS = 500;

/** Max aggregate items in fast mode. */
const FAST_MODE_MAX_AGGREGATE_ITEMS = 100;

/** Default page size ($top) for single-page requests (e.g. calendar view). */
const DEFAULT_PAGE_SIZE = 50;

/** Page size in fast mode. */
const FAST_MODE_PAGE_SIZE = 20;

/** Default max retries for Graph requests. */
const DEFAULT_GRAPH_MAX_RETRIES = 3;

/** Max retries in fast mode. */
const FAST_MODE_GRAPH_MAX_RETRIES = 1;

/** Default max backoff delay (ms) for retries. */
const DEFAULT_GRAPH_RETRY_MAX_DELAY_MS = 30000;

/** Max retry delay (ms) in fast mode. */
const FAST_MODE_GRAPH_RETRY_MAX_DELAY_MS = 5000;

function parsePositiveInt(value: string | undefined, fallback: number): number {
  if (value === undefined || value === '') return fallback;
  const parsed = parseInt(value, 10);
  return Number.isNaN(parsed) || parsed < 0 ? fallback : parsed;
}

/**
 * Whether fast mode is enabled (lower limits, shorter timeouts, fewer retries).
 */
export function isFastMode(): boolean {
  return process.env.MS365_MCP_FAST_MODE === 'true';
}

/**
 * Graph request timeout in milliseconds.
 * Used for fetch AbortSignal in performRequest and performBatch.
 */
export function getGraphRequestTimeoutMs(): number {
  if (isFastMode()) return FAST_MODE_GRAPH_REQUEST_TIMEOUT_MS;
  const env = process.env.MS365_MCP_GRAPH_REQUEST_TIMEOUT_MS;
  const parsed = parsePositiveInt(env, DEFAULT_GRAPH_REQUEST_TIMEOUT_MS);
  return parsed > 0 ? parsed : DEFAULT_GRAPH_REQUEST_TIMEOUT_MS;
}

/**
 * Max pagination pages (e.g. graph-tools, compound-tools).
 */
export function getMaxPages(): number {
  if (isFastMode()) return FAST_MODE_MAX_PAGES;
  return parsePositiveInt(process.env.MS365_MCP_MAX_PAGES, DEFAULT_MAX_PAGES);
}

/**
 * Max results per query (discovery, intelligent search, etc.).
 */
export function getMaxResults(): number {
  if (isFastMode()) return FAST_MODE_MAX_RESULTS;
  return parsePositiveInt(process.env.MS365_MCP_MAX_RESULTS, DEFAULT_MAX_RESULTS);
}

/**
 * Max aggregate items (data-aggregator, discovery).
 */
export function getMaxAggregateItems(): number {
  if (isFastMode()) return FAST_MODE_MAX_AGGREGATE_ITEMS;
  return parsePositiveInt(process.env.MS365_MCP_MAX_AGGREGATE_ITEMS, DEFAULT_MAX_AGGREGATE_ITEMS);
}

/**
 * Default page size ($top) for endpoints that benefit from a smaller first page (e.g. calendar).
 */
export function getDefaultPageSize(): number {
  if (isFastMode()) return FAST_MODE_PAGE_SIZE;
  return DEFAULT_PAGE_SIZE;
}

/**
 * Max retries for Graph API calls (makeRequest, callGraphWithRetry).
 */
export function getGraphMaxRetries(): number {
  if (isFastMode()) return FAST_MODE_GRAPH_MAX_RETRIES;
  return parsePositiveInt(process.env.MS365_MCP_GRAPH_MAX_RETRIES, DEFAULT_GRAPH_MAX_RETRIES);
}

/**
 * Max backoff delay in ms for retries (429/5xx).
 */
export function getGraphRetryMaxDelayMs(): number {
  if (isFastMode()) return FAST_MODE_GRAPH_RETRY_MAX_DELAY_MS;
  return parsePositiveInt(
    process.env.MS365_MCP_GRAPH_RETRY_MAX_DELAY_MS,
    DEFAULT_GRAPH_RETRY_MAX_DELAY_MS
  );
}
