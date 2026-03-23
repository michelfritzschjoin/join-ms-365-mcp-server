/**
 * Graph API OData pagination helper.
 * Follows @odata.nextLink and merges value arrays for SharePoint/Files list responses.
 */

import logger from '../logger.js';
import { getMaxPages } from '../perf-config.js';

/** Minimal client interface for pagination (makeRequest). */
export interface GraphClientLike {
  makeRequest(
    endpoint: string,
    options?: { method?: string; queryParams?: Record<string, string> }
  ): Promise<unknown>;
}

/** OData collection response shape. */
export interface ODataCollectionResponse {
  value?: unknown[];
  '@odata.nextLink'?: string;
  '@odata.count'?: number;
  [key: string]: unknown;
}

/**
 * Fetches all pages of an OData collection by following @odata.nextLink.
 * Merges value arrays and returns a single combined response object.
 *
 * @param graphClient - Graph client with makeRequest (path, options)
 * @param firstResponse - First page: JSON string or parsed object
 * @param maxPages - Maximum number of pages to fetch (default from getMaxPages())
 * @param thinking - Optional array to push pagination progress messages
 * @returns Combined response object with merged value and no @odata.nextLink
 */
export async function fetchAllODataPages(
  graphClient: GraphClientLike,
  firstResponse: string | ODataCollectionResponse,
  maxPages?: number,
  thinking?: string[]
): Promise<ODataCollectionResponse> {
  const limit = maxPages ?? getMaxPages();
  let parsed: ODataCollectionResponse;
  if (typeof firstResponse === 'string') {
    try {
      parsed = JSON.parse(firstResponse) as ODataCollectionResponse;
    } catch {
      return { value: [] };
    }
  } else {
    parsed = { ...firstResponse };
  }

  const value = parsed.value;
  if (!Array.isArray(value)) {
    return parsed;
  }

  let allItems = [...value];
  let nextLink = parsed['@odata.nextLink'];
  let pageCount = 1;

  if (nextLink && thinking) {
    thinking.push(`Pagination: ${allItems.length} items from page 1, max ${limit} pages`);
  }

  while (nextLink && pageCount < limit) {
    try {
      const url = new URL(nextLink);
      const nextPath = url.pathname.replace(/^\/v1\.0/, '');
      const nextQueryParams: Record<string, string> = {};
      for (const [key, val] of url.searchParams.entries()) {
        nextQueryParams[key] = val;
      }

      logger.debug(`Fetching OData page ${pageCount + 1} from: ${nextPath}`);
      const nextResponse = (await graphClient.makeRequest(nextPath, {
        method: 'GET',
        queryParams: nextQueryParams,
      })) as ODataCollectionResponse;

      if (nextResponse?.value && Array.isArray(nextResponse.value)) {
        allItems = allItems.concat(nextResponse.value);
      }
      nextLink = nextResponse?.['@odata.nextLink'];
      pageCount++;

      if (thinking) {
        thinking.push(`Loaded page ${pageCount}: ${allItems.length} items total`);
      }
    } catch (err) {
      logger.warn(`OData pagination error on page ${pageCount + 1}: ${(err as Error).message}`);
      if (thinking) {
        thinking.push(`Stopped after page ${pageCount} due to error`);
      }
      break;
    }
  }

  if (pageCount >= limit && nextLink) {
    logger.warn(`Reached maximum page limit (${limit}) for OData pagination`);
    if (thinking) {
      thinking.push(`Further pages available; use skip/top for next page.`);
    }
  }

  const combined: ODataCollectionResponse = {
    ...parsed,
    value: allItems,
  };
  if (combined['@odata.count'] !== undefined) {
    combined['@odata.count'] = allItems.length;
  }
  delete combined['@odata.nextLink'];

  return combined;
}
