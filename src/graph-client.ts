import logger from './logger.js';
import AuthManager from './auth.js';
import { refreshAccessToken } from './lib/microsoft-auth.js';
import { encode as toonEncode } from '@toon-format/toon';
import type { AppSecrets } from './secrets.js';
import { getCloudEndpoints } from './cloud-config.js';
import { getRequestTokens } from './request-context.js';
import {
  GraphApiError,
  RateLimitError,
  ServiceUnavailableError,
  AuthenticationError,
  AuthorizationError,
  NotFoundError,
  isRetryableError,
  getRetryAfter,
} from './errors.js';
import { GraphApiRepairManager } from './graph-api-repair.js';
import {
  getGraphRequestTimeoutMs,
  getGraphMaxRetries,
  getGraphRetryMaxDelayMs,
} from './perf-config.js';

interface GraphRequestOptions {
  headers?: Record<string, string>;
  method?: string;
  body?: string;
  rawResponse?: boolean;
  includeHeaders?: boolean;
  excludeResponse?: boolean;
  accessToken?: string;
  refreshToken?: string;
  queryParams?: Record<string, string>;

  [key: string]: unknown;
}

/** Single request for Microsoft Graph JSON batching (POST /$batch). URL is relative to /v1.0 (e.g. /me/messages). */
export interface GraphBatchRequest {
  id: string;
  method: 'GET' | 'POST' | 'PATCH' | 'DELETE';
  url: string;
  body?: unknown;
  headers?: Record<string, string>;
}

/** Single response from a batch request. */
export interface GraphBatchResponse {
  id: string;
  status: number;
  body?: unknown;
  headers?: Record<string, string>;
}

interface ContentItem {
  type: 'text';
  text: string;

  [key: string]: unknown;
}

interface McpResponse {
  content: ContentItem[];
  _meta?: Record<string, unknown>;
  isError?: boolean;

  [key: string]: unknown;
}

class GraphClient {
  private authManager: AuthManager;
  private secrets: AppSecrets;
  private readonly outputFormat: 'json' | 'toon' = 'json';
  private repairManager: GraphApiRepairManager | null = null;

  constructor(
    authManager: AuthManager,
    secrets: AppSecrets,
    outputFormat: 'json' | 'toon' = 'json'
  ) {
    this.authManager = authManager;
    this.secrets = secrets;
    this.outputFormat = outputFormat;

    // Initialize repair manager if enabled
    try {
      this.repairManager = new GraphApiRepairManager();
      if (this.repairManager.isEnabled()) {
        logger.info('Graph API Self-Repair system enabled');
      }
    } catch (error) {
      logger.warn(`Failed to initialize repair manager: ${error}`);
    }
  }

  /**
   * Get the current output format
   */
  getOutputFormat(): 'json' | 'toon' {
    return this.outputFormat;
  }

  async makeRequest(
    endpoint: string,
    options: GraphRequestOptions = {},
    maxRetries = getGraphMaxRetries()
  ): Promise<unknown> {
    const contextTokens = getRequestTokens();

    // Try to get token from various sources
    let accessToken: string | null = null;
    try {
      accessToken =
        options.accessToken ?? contextTokens?.accessToken ?? (await this.authManager.getToken());
    } catch (error) {
      // Token acquisition failed - this is OK, we'll throw a helpful error below
      logger.debug('Token acquisition failed, will return authentication error', {
        error: (error as Error).message,
      });
    }

    const refreshToken = options.refreshToken ?? contextTokens?.refreshToken;

    if (!accessToken) {
      throw new AuthenticationError(
        'AUTHENTICATION REQUIRED: You are not logged in to Microsoft 365. ' +
          'If you are using an MCP client with OAuth (for example Open WebUI), complete the OAuth authorization flow ' +
          'and retry the request. If you are running this server in CLI/stdio mode, call the "login" tool to ' +
          'authenticate via device code flow.'
      );
    }

    let lastError: unknown;
    let attempt = 0;

    while (attempt <= maxRetries) {
      try {
        let response = await this.performRequest(endpoint, accessToken, options);

        // Handle 401 - Token expired
        if (response.status === 401 && refreshToken) {
          // Token expired, try to refresh
          const newTokens = await this.refreshAccessToken(refreshToken);
          accessToken = newTokens.accessToken;

          // Retry the request with new token
          response = await this.performRequest(endpoint, accessToken, options);
        }

        // Handle non-OK responses - try self-repair before throwing
        if (!response.ok) {
          const errorText = await response.text();
          let error: GraphApiError;

          // Create appropriate error type with actionable next steps for LLM
          if (response.status === 401) {
            throw new AuthenticationError(
              'Re-authenticate: use login or refresh token. If using CLI/stdio, call the "login" tool; if using OAuth, complete the authorization flow and retry. ' +
                (errorText ? errorText.slice(0, 200) : ''),
              response
            );
          }
          if (response.status === 403) {
            const base =
              errorText.includes('scope') || errorText.includes('permission')
                ? `Microsoft Graph API scope error - ${errorText.slice(0, 300)}. This tool may require organization mode (--org-mode).`
                : `Microsoft Graph API error: ${response.status} - ${errorText.slice(0, 300)}`;
            error = new AuthorizationError(
              base + ' Next step: Check Azure AD app permissions for this operation.',
              response
            );
          } else if (response.status === 404) {
            error = new NotFoundError(
              'Recipient or resource not found. Next step: Use list-users or search to resolve user/email before sending; or verify the resource ID. ' +
                (errorText ? errorText.slice(0, 200) : ''),
              response
            );
          } else if (response.status === 429) {
            const retryAfterHeader = response.headers.get('Retry-After');
            const retryAfter = retryAfterHeader ? parseInt(retryAfterHeader, 10) : undefined;
            error = new RateLimitError(retryAfter, response);
          } else if (response.status === 503) {
            const retryAfterHeader = response.headers.get('Retry-After');
            const retryAfter = retryAfterHeader ? parseInt(retryAfterHeader, 10) : undefined;
            error = new ServiceUnavailableError(retryAfter, response);
          } else {
            error = new GraphApiError(
              `Microsoft Graph API error: ${response.status} ${response.statusText} - ${errorText}`,
              response.status,
              false,
              undefined,
              response
            );
          }

          // Try self-repair if enabled and not authentication error
          if (
            this.repairManager?.isEnabled() &&
            !(error instanceof AuthenticationError) &&
            attempt < maxRetries
          ) {
            try {
              const repairRequest = this.repairManager.createRepairRequest(
                endpoint,
                options,
                error,
                errorText
              );

              const repairResult = await this.repairManager.attemptRepair(repairRequest);

              if (repairResult?.success && repairResult.repairedRequest) {
                const repaired = repairResult.repairedRequest;
                logger.info(`Self-repair successful, retrying with repaired request`, {
                  strategy: repairResult.strategy,
                  originalEndpoint: endpoint,
                  repairedEndpoint: repaired.endpoint,
                });

                // Update endpoint and options from repair
                endpoint = repaired.endpoint;
                Object.assign(options, repaired.options);

                // Retry with repaired request
                attempt++;
                continue;
              }
            } catch (repairError) {
              logger.warn(`Self-repair attempt failed: ${repairError}`);
              // Continue to throw original error
            }
          }

          // No repair or repair failed, throw error
          throw error;
        }

        const text = await response.text();
        let result: unknown;

        if (text === '') {
          result = { message: 'OK!' };
        } else {
          try {
            result = JSON.parse(text);
          } catch {
            result = { message: 'OK!', rawResponse: text };
          }
        }

        // If includeHeaders is requested, add response headers to the result
        if (options.includeHeaders) {
          const etag = response.headers.get('ETag') || response.headers.get('etag');

          // Simple approach: just add ETag to the result if it's an object
          if (result && typeof result === 'object' && !Array.isArray(result)) {
            return {
              ...result,
              _etag: etag || 'no-etag-found',
            };
          }
        }

        return result;
      } catch (error) {
        lastError = error;

        // Check if error is retryable
        if (isRetryableError(error) && attempt < maxRetries) {
          const retryAfter = getRetryAfter(error);
          const backoffDelay = this.calculateBackoff(attempt, retryAfter);

          logger.warn(
            `Retryable error (attempt ${attempt + 1}/${maxRetries}): ${(error as Error).message}. Retrying in ${backoffDelay}ms`
          );

          await this.sleep(backoffDelay);
          attempt++;
          continue;
        }

        // Not retryable or max retries reached
        if (error instanceof GraphApiError) {
          throw error;
        }

        logger.error('Microsoft Graph API request failed:', error);
        throw new GraphApiError(
          (error as Error).message || 'Unknown error',
          500,
          false,
          undefined,
          error
        );
      }
    }

    // Should not reach here, but handle it anyway
    throw lastError;
  }

  /**
   * Fetch binary content from a Graph endpoint (e.g. /me/drive/items/{id}/content).
   * Use for Office file downloads (Word, PowerPoint) for content extraction.
   */
  async getBinaryContent(endpoint: string): Promise<ArrayBuffer> {
    const contextTokens = getRequestTokens();
    let accessToken: string | null = null;
    try {
      accessToken = contextTokens?.accessToken ?? (await this.authManager.getToken());
    } catch {
      logger.debug('Token acquisition failed for binary content');
    }
    const refreshToken = contextTokens?.refreshToken;
    if (!accessToken) {
      throw new AuthenticationError(
        'AUTHENTICATION REQUIRED: You are not logged in to Microsoft 365.'
      );
    }
    let response = await this.performRequest(endpoint, accessToken, { method: 'GET' });
    if (response.status === 401 && refreshToken) {
      const newTokens = await this.refreshAccessToken(refreshToken);
      accessToken = newTokens.accessToken;
      response = await this.performRequest(endpoint, accessToken, { method: 'GET' });
    }
    if (!response.ok) {
      const errorText = await response.text();
      throw new GraphApiError(
        `Graph API error ${response.status}: ${errorText.slice(0, 200)}`,
        response.status,
        false,
        undefined,
        response
      );
    }
    return response.arrayBuffer();
  }

  /**
   * Calculate exponential backoff delay (capped by perf-config max delay).
   */
  private calculateBackoff(attempt: number, retryAfter?: number): number {
    // If Retry-After header is provided, use it (capped by max delay)
    const maxDelay = getGraphRetryMaxDelayMs();
    if (retryAfter) {
      return Math.min(retryAfter * 1000, maxDelay);
    }

    // Exponential backoff: 1s, 2s, 4s, 8s, etc., capped by config
    const baseDelay = 1000; // 1 second
    const delay = Math.min(baseDelay * Math.pow(2, attempt), maxDelay);

    // Add jitter to prevent thundering herd
    const jitter = Math.random() * 1000; // 0-1 second jitter
    return delay + jitter;
  }

  /**
   * Sleep for specified milliseconds
   */
  private sleep(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  private async refreshAccessToken(
    refreshToken: string
  ): Promise<{ accessToken: string; refreshToken?: string }> {
    const tenantId = this.secrets.tenantId || 'common';
    const clientId = this.secrets.clientId;
    const clientSecret = this.secrets.clientSecret;

    // Log whether using public or confidential client
    if (clientSecret) {
      logger.info('GraphClient: Refreshing token with confidential client');
    } else {
      logger.info('GraphClient: Refreshing token with public client');
    }

    const response = await refreshAccessToken(
      refreshToken,
      clientId,
      clientSecret,
      tenantId,
      this.secrets.cloudType
    );

    return {
      accessToken: response.access_token,
      refreshToken: response.refresh_token,
    };
  }

  private async performRequest(
    endpoint: string,
    accessToken: string,
    options: GraphRequestOptions
  ): Promise<Response> {
    const cloudEndpoints = getCloudEndpoints(this.secrets.cloudType);
    let url = `${cloudEndpoints.graphApi}/v1.0${endpoint}`;

    // Append queryParams to URL if provided
    if (options.queryParams && Object.keys(options.queryParams).length > 0) {
      const queryString = Object.entries(options.queryParams)
        .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
        .join('&');
      url = `${url}${url.includes('?') ? '&' : '?'}${queryString}`;
    }

    const headers: Record<string, string> = {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      ...options.headers,
    };

    const timeoutMs = getGraphRequestTimeoutMs();
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), timeoutMs);

    try {
      const response = await fetch(url, {
        method: options.method || 'GET',
        headers,
        body: options.body,
        signal: controller.signal,
      });
      clearTimeout(timeoutId);
      return response;
    } catch (error) {
      clearTimeout(timeoutId);
      if ((error as Error).name === 'AbortError') {
        throw new GraphApiError(
          `Microsoft Graph API request timed out after ${timeoutMs}ms. Consider MS365_MCP_GRAPH_REQUEST_TIMEOUT_MS or MS365_MCP_FAST_MODE for faster failure.`,
          408,
          true
        );
      }
      throw error;
    }
  }

  private serializeData(data: unknown, outputFormat: 'json' | 'toon', pretty = false): string {
    if (outputFormat === 'toon') {
      try {
        return toonEncode(data);
      } catch (error) {
        logger.warn(`Failed to encode as TOON, falling back to JSON: ${error}`);
        return JSON.stringify(data, null, pretty ? 2 : undefined);
      }
    }
    return JSON.stringify(data, null, pretty ? 2 : undefined);
  }

  /**
   * Execute multiple Graph API requests in a single HTTP call using JSON batching.
   * Uses POST /v1.0/$batch (max 20 requests per batch). Each request url must be
   * relative to /v1.0 (e.g. /me/calendar/calendarView?startDateTime=...).
   *
   * @param requests - Array of batch requests (id, method, url, optional body/headers)
   * @returns Array of responses in the same order as requests (id, status, body)
   */
  async performBatch(requests: GraphBatchRequest[]): Promise<GraphBatchResponse[]> {
    if (requests.length === 0) return [];
    if (requests.length > 20) {
      throw new GraphApiError(
        'Graph batch supports at most 20 requests per call. Split into multiple batches.',
        400,
        false
      );
    }

    const contextTokens = getRequestTokens();
    let accessToken: string | null = null;
    try {
      accessToken = contextTokens?.accessToken ?? (await this.authManager.getToken());
    } catch (error) {
      logger.debug('Token acquisition failed for batch', { error: (error as Error).message });
    }
    const refreshToken = contextTokens?.refreshToken;

    if (!accessToken) {
      throw new AuthenticationError(
        'AUTHENTICATION REQUIRED: You are not logged in to Microsoft 365. ' +
          'Call the "login" tool to authenticate or complete OAuth flow.'
      );
    }

    const batchBody = {
      requests: requests.map((r) => {
        const req: {
          id: string;
          method: string;
          url: string;
          body?: unknown;
          headers?: Record<string, string>;
        } = {
          id: r.id,
          method: r.method,
          url: r.url.startsWith('/') ? r.url : `/${r.url}`,
        };
        if (r.body !== undefined) {
          req.body = r.body;
          req.headers = { 'Content-Type': 'application/json', ...r.headers };
        } else if (r.headers && Object.keys(r.headers).length > 0) {
          req.headers = r.headers;
        }
        return req;
      }),
    };

    const cloudEndpoints = getCloudEndpoints(this.secrets.cloudType);
    const url = `${cloudEndpoints.graphApi}/v1.0/$batch`;
    const timeoutMs = getGraphRequestTimeoutMs();
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), timeoutMs);

    try {
      let response = await fetch(url, {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(batchBody),
        signal: controller.signal,
      });

      if (response.status === 401 && refreshToken) {
        const newTokens = await this.refreshAccessToken(refreshToken);
        response = await fetch(url, {
          method: 'POST',
          headers: {
            Authorization: `Bearer ${newTokens.accessToken}`,
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(batchBody),
          signal: controller.signal,
        });
      }

      clearTimeout(timeoutId);

      if (!response.ok) {
        const errorText = await response.text();
        if (response.status === 401) {
          throw new AuthenticationError(
            'Re-authenticate: use login or refresh token. ' + errorText,
            undefined
          );
        }
        if (response.status === 429) {
          const retryAfterHeader = response.headers.get('Retry-After');
          const retryAfter = retryAfterHeader ? parseInt(retryAfterHeader, 10) : undefined;
          throw new RateLimitError(retryAfter, response);
        }
        throw new GraphApiError(
          `Graph batch request failed: ${response.status} ${response.statusText} - ${errorText}`,
          response.status,
          false,
          undefined,
          response
        );
      }

      const data = (await response.json()) as {
        responses?: Array<{
          id: string;
          status: number;
          body?: unknown;
          headers?: Record<string, string>;
        }>;
      };
      const responses = data.responses ?? [];
      return responses.map((r) => ({
        id: r.id,
        status: r.status,
        body: r.body,
        headers: r.headers,
      }));
    } catch (error) {
      clearTimeout(timeoutId);
      if ((error as Error).name === 'AbortError') {
        throw new GraphApiError(
          `Microsoft Graph batch request timed out after ${timeoutMs}ms. Consider MS365_MCP_GRAPH_REQUEST_TIMEOUT_MS or MS365_MCP_FAST_MODE.`,
          408,
          true
        );
      }
      throw error;
    }
  }

  async graphRequest(endpoint: string, options: GraphRequestOptions = {}): Promise<McpResponse> {
    try {
      logger.info(`Calling ${endpoint} with options: ${JSON.stringify(options)}`);

      // Use new OAuth-aware request method
      const result = await this.makeRequest(endpoint, options);

      return this.formatJsonResponse(result, options.rawResponse, options.excludeResponse);
    } catch (error) {
      logger.error(`Error in Graph API request: ${error}`);
      return {
        content: [{ type: 'text', text: JSON.stringify({ error: (error as Error).message }) }],
        isError: true,
      };
    }
  }

  formatJsonResponse(data: unknown, rawResponse = false, excludeResponse = false): McpResponse {
    // If excludeResponse is true, only return success indication
    if (excludeResponse) {
      return {
        content: [{ type: 'text', text: this.serializeData({ success: true }, this.outputFormat) }],
      };
    }

    // Handle the case where data includes headers metadata
    if (data && typeof data === 'object' && '_headers' in data) {
      const responseData = data as {
        data: unknown;
        _headers: Record<string, string>;
        _etag?: string;
      };

      const meta: Record<string, unknown> = {};
      if (responseData._etag) {
        meta.etag = responseData._etag;
      }
      if (responseData._headers) {
        meta.headers = responseData._headers;
      }

      if (rawResponse) {
        return {
          content: [
            { type: 'text', text: this.serializeData(responseData.data, this.outputFormat) },
          ],
          _meta: meta,
        };
      }

      if (responseData.data === null || responseData.data === undefined) {
        return {
          content: [
            { type: 'text', text: this.serializeData({ success: true }, this.outputFormat) },
          ],
          _meta: meta,
        };
      }

      // Remove OData properties
      const removeODataProps = (obj: Record<string, unknown>): void => {
        if (typeof obj === 'object' && obj !== null) {
          Object.keys(obj).forEach((key) => {
            if (key.startsWith('@odata.')) {
              delete obj[key];
            } else if (typeof obj[key] === 'object') {
              removeODataProps(obj[key] as Record<string, unknown>);
            }
          });
        }
      };

      removeODataProps(responseData.data as Record<string, unknown>);

      return {
        content: [
          { type: 'text', text: this.serializeData(responseData.data, this.outputFormat, true) },
        ],
        _meta: meta,
      };
    }

    // Original handling for backward compatibility
    if (rawResponse) {
      return {
        content: [{ type: 'text', text: this.serializeData(data, this.outputFormat) }],
      };
    }

    if (data === null || data === undefined) {
      return {
        content: [{ type: 'text', text: this.serializeData({ success: true }, this.outputFormat) }],
      };
    }

    // Remove OData properties
    const removeODataProps = (obj: Record<string, unknown>): void => {
      if (typeof obj === 'object' && obj !== null) {
        Object.keys(obj).forEach((key) => {
          if (key.startsWith('@odata.')) {
            delete obj[key];
          } else if (typeof obj[key] === 'object') {
            removeODataProps(obj[key] as Record<string, unknown>);
          }
        });
      }
    };

    removeODataProps(data as Record<string, unknown>);

    return {
      content: [{ type: 'text', text: this.serializeData(data, this.outputFormat, true) }],
    };
  }
}

export default GraphClient;
