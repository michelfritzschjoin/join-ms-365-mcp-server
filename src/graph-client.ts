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
  isRetryableError,
  getRetryAfter,
} from './errors.js';
import { GraphApiRepairManager } from './graph-api-repair.js';

interface GraphRequestOptions {
  headers?: Record<string, string>;
  method?: string;
  body?: string;
  rawResponse?: boolean;
  includeHeaders?: boolean;
  excludeResponse?: boolean;
  accessToken?: string;
  refreshToken?: string;

  [key: string]: unknown;
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

  async makeRequest(
    endpoint: string,
    options: GraphRequestOptions = {},
    maxRetries = 3
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
          'Please use the "login" tool first to authenticate before using any Microsoft 365 tools. ' +
          'After running the "login" tool, follow the device code instructions to complete authentication.'
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

          // Create appropriate error type
          if (response.status === 403) {
            if (errorText.includes('scope') || errorText.includes('permission')) {
              error = new AuthorizationError(
                `Microsoft Graph API scope error: ${response.status} ${response.statusText} - ${errorText}. This tool requires organization mode. Please restart with --org-mode flag.`,
                response
              );
            } else {
              error = new AuthorizationError(
                `Microsoft Graph API error: ${response.status} ${response.statusText} - ${errorText}`,
                response
              );
            }
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
   * Calculate exponential backoff delay
   */
  private calculateBackoff(attempt: number, retryAfter?: number): number {
    // If Retry-After header is provided, use it
    if (retryAfter) {
      return retryAfter * 1000; // Convert seconds to milliseconds
    }

    // Exponential backoff: 1s, 2s, 4s, 8s, etc.
    const baseDelay = 1000; // 1 second
    const maxDelay = 30000; // 30 seconds max
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
    const url = `${cloudEndpoints.graphApi}/v1.0${endpoint}`;

    const headers: Record<string, string> = {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      ...options.headers,
    };

    return fetch(url, {
      method: options.method || 'GET',
      headers,
      body: options.body,
    });
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
