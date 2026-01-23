/**
 * Repair strategies for Graph API Self-Repair System
 */

import logger from './logger.js';
import type { ParsedGraphApiError } from './errors.js';
import {
  getAlternativeEndpoint,
  normalizeODataParam,
  getParentScopes,
  PARAMETER_RULES,
} from './repair-config.js';

/**
 * Repair request context
 */
export interface RepairRequest {
  endpoint: string;
  method: string;
  options: {
    headers?: Record<string, string>;
    body?: string;
    [key: string]: unknown;
  };
  params?: Record<string, unknown>;
  originalError: ParsedGraphApiError;
}

/**
 * Repair result
 */
export interface RepairResult {
  success: boolean;
  repairedRequest?: RepairRequest;
  strategy: string;
  message: string;
  metadata?: Record<string, unknown>;
}

/**
 * Base interface for repair strategies
 */
export interface RepairStrategy {
  name: string;
  canRepair(error: ParsedGraphApiError): boolean;
  repair(request: RepairRequest): Promise<RepairResult>;
}

/**
 * Endpoint Repair Strategy
 * Tries alternative endpoint paths when 404 occurs
 */
export class EndpointRepairStrategy implements RepairStrategy {
  name = 'endpoint';

  canRepair(error: ParsedGraphApiError): boolean {
    return error.statusCode === 404 || error.resourceNotFound === true;
  }

  async repair(request: RepairRequest): Promise<RepairResult> {
    try {
      const { endpoint, params = {} } = request;
      const alternative = getAlternativeEndpoint(endpoint, params);

      if (alternative && alternative !== endpoint) {
        logger.info(`Endpoint repair: ${endpoint} -> ${alternative}`);
        return {
          success: true,
          repairedRequest: {
            ...request,
            endpoint: alternative,
          },
          strategy: this.name,
          message: `Trying alternative endpoint: ${alternative}`,
          metadata: {
            originalEndpoint: endpoint,
            alternativeEndpoint: alternative,
          },
        };
      }

      // Try /me -> /users/{userId} if we can extract userId
      if (endpoint.startsWith('/me/')) {
        const rest = endpoint.substring(4);
        // Try to get userId from params or context
        const userId = params.userId || params['user-id'] || params.id;
        if (userId) {
          const alternativeEndpoint = `/users/${userId}${rest}`;
          logger.info(`Endpoint repair: ${endpoint} -> ${alternativeEndpoint}`);
          return {
            success: true,
            repairedRequest: {
              ...request,
              endpoint: alternativeEndpoint,
            },
            strategy: this.name,
            message: `Trying alternative endpoint: ${alternativeEndpoint}`,
            metadata: {
              originalEndpoint: endpoint,
              alternativeEndpoint: alternativeEndpoint,
            },
          };
        }
      }

      // Try /v1.0 -> /beta
      if (endpoint.includes('/v1.0/')) {
        const betaEndpoint = endpoint.replace('/v1.0/', '/beta/');
        logger.info(`Endpoint repair: ${endpoint} -> ${betaEndpoint}`);
        return {
          success: true,
          repairedRequest: {
            ...request,
            endpoint: betaEndpoint,
          },
          strategy: this.name,
          message: `Trying beta endpoint: ${betaEndpoint}`,
          metadata: {
            originalEndpoint: endpoint,
            alternativeEndpoint: betaEndpoint,
          },
        };
      }

      return {
        success: false,
        strategy: this.name,
        message: 'No alternative endpoint found',
      };
    } catch (error) {
      logger.error(`Endpoint repair strategy error: ${error}`);
      return {
        success: false,
        strategy: this.name,
        message: `Endpoint repair failed: ${(error as Error).message}`,
      };
    }
  }
}

/**
 * Parameter Repair Strategy
 * Fixes invalid or malformed parameters
 */
export class ParameterRepairStrategy implements RepairStrategy {
  name = 'parameter';

  canRepair(error: ParsedGraphApiError): boolean {
    return (
      error.statusCode === 400 ||
      error.invalidParameters !== undefined ||
      error.message.toLowerCase().includes('invalid') ||
      error.message.toLowerCase().includes('parameter')
    );
  }

  async repair(request: RepairRequest): Promise<RepairResult> {
    try {
      const { options, params = {} } = request;
      const repairedOptions = { ...options };
      const repairedParams = { ...params };
      let repaired = false;
      const fixes: string[] = [];

      // Normalize OData parameters
      if (repairedOptions.headers) {
        const headers = { ...repairedOptions.headers };
        let headerFixed = false;

        // Fix ConsistencyLevel header for search
        if (params.search && !headers['ConsistencyLevel']) {
          headers['ConsistencyLevel'] = 'eventual';
          headerFixed = true;
          fixes.push('Added ConsistencyLevel header for search');
        }

        if (headerFixed) {
          repairedOptions.headers = headers;
          repaired = true;
        }
      }

      // Fix query parameters
      const queryParams: Record<string, string> = {};
      if (options.queryParams) {
        Object.assign(queryParams, options.queryParams);
      }

      // Normalize OData parameter names
      for (const [key, value] of Object.entries(queryParams)) {
        const normalized = normalizeODataParam(key);
        if (normalized !== key) {
          delete queryParams[key];
          queryParams[normalized] = String(value);
          fixes.push(`Normalized parameter: ${key} -> ${normalized}`);
          repaired = true;
        }
      }

      // Apply parameter rules
      for (const rule of PARAMETER_RULES) {
        const paramValue = repairedParams[rule.name] || queryParams[rule.name];
        if (paramValue !== undefined && rule.fixer) {
          const fixed = rule.fixer(paramValue);
          if (fixed !== paramValue) {
            if (repairedParams[rule.name]) {
              repairedParams[rule.name] = fixed;
            } else {
              queryParams[rule.name] = String(fixed);
            }
            fixes.push(`Fixed ${rule.name} parameter format`);
            repaired = true;
          }
        }
      }

      // Fix search parameter format
      if (queryParams.$search || queryParams.search) {
        const searchParam = queryParams.$search || queryParams.search;
        if (typeof searchParam === 'string') {
          // Ensure search is wrapped in quotes if not already
          let fixedSearch = searchParam.trim();
          if (!fixedSearch.startsWith('"') && !fixedSearch.startsWith("'")) {
            fixedSearch = `"${fixedSearch}"`;
            if (queryParams.$search) {
              queryParams.$search = fixedSearch;
            } else {
              delete queryParams.search;
              queryParams.$search = fixedSearch;
            }
            fixes.push('Fixed search parameter format (added quotes)');
            repaired = true;
          }
        }
      }

      if (repaired) {
        repairedOptions.queryParams = queryParams;
        return {
          success: true,
          repairedRequest: {
            ...request,
            options: repairedOptions,
            params: repairedParams,
          },
          strategy: this.name,
          message: `Fixed parameters: ${fixes.join(', ')}`,
          metadata: {
            fixes,
          },
        };
      }

      return {
        success: false,
        strategy: this.name,
        message: 'No parameter fixes applicable',
      };
    } catch (error) {
      logger.error(`Parameter repair strategy error: ${error}`);
      return {
        success: false,
        strategy: this.name,
        message: `Parameter repair failed: ${(error as Error).message}`,
      };
    }
  }
}

/**
 * Scope Repair Strategy
 * Attempts to use parent scopes or alternative endpoints
 */
export class ScopeRepairStrategy implements RepairStrategy {
  name = 'scope';

  canRepair(error: ParsedGraphApiError): boolean {
    return (
      error.statusCode === 403 ||
      error.missingScopes !== undefined ||
      error.message.toLowerCase().includes('permission') ||
      error.message.toLowerCase().includes('scope') ||
      error.message.toLowerCase().includes('authorization')
    );
  }

  async repair(request: RepairRequest): Promise<RepairResult> {
    try {
      const { originalError } = request;

      // If we have missing scopes, try to find parent scopes
      if (originalError.missingScopes && originalError.missingScopes.length > 0) {
        const missingScopes = originalError.missingScopes;
        const parentScopes: string[] = [];

        for (const scope of missingScopes) {
          const parents = getParentScopes(scope);
          parentScopes.push(...parents);
        }

        if (parentScopes.length > 0) {
          logger.info(
            `Scope repair: Missing scopes ${missingScopes.join(', ')}, trying parent scopes: ${parentScopes.join(', ')}`
          );
          return {
            success: true,
            repairedRequest: request,
            strategy: this.name,
            message: `Request requires parent scopes: ${parentScopes.join(', ')}`,
            metadata: {
              missingScopes,
              suggestedScopes: parentScopes,
              note: 'User may need to re-authenticate with additional scopes',
            },
          };
        }
      }

      // Try alternative endpoint patterns that might require different scopes
      // For example, /me/* might work with different scopes than /users/{id}/*
      const { endpoint } = request;
      if (endpoint.startsWith('/users/') && !endpoint.startsWith('/me/')) {
        // Try /me instead
        const meEndpoint = endpoint.replace(/^\/users\/[^/]+/, '/me');
        logger.info(`Scope repair: Trying /me endpoint: ${meEndpoint}`);
        return {
          success: true,
          repairedRequest: {
            ...request,
            endpoint: meEndpoint,
          },
          strategy: this.name,
          message: `Trying /me endpoint which may require different scopes`,
          metadata: {
            originalEndpoint: endpoint,
            alternativeEndpoint: meEndpoint,
          },
        };
      }

      return {
        success: false,
        strategy: this.name,
        message: 'No scope repair available - user may need to re-authenticate',
        metadata: {
          note: 'User may need to re-authenticate with additional scopes',
        },
      };
    } catch (error) {
      logger.error(`Scope repair strategy error: ${error}`);
      return {
        success: false,
        strategy: this.name,
        message: `Scope repair failed: ${(error as Error).message}`,
      };
    }
  }
}

/**
 * Version Repair Strategy
 * Tries beta version when v1.0 fails
 */
export class VersionRepairStrategy implements RepairStrategy {
  name = 'version';

  canRepair(error: ParsedGraphApiError): boolean {
    return (
      error.deprecatedApi === true ||
      error.message.toLowerCase().includes('deprecated') ||
      error.message.toLowerCase().includes('obsolete') ||
      error.message.toLowerCase().includes('version')
    );
  }

  async repair(request: RepairRequest): Promise<RepairResult> {
    try {
      const { endpoint } = request;

      // Try beta version if using v1.0
      if (endpoint.includes('/v1.0/')) {
        const betaEndpoint = endpoint.replace('/v1.0/', '/beta/');
        logger.info(`Version repair: ${endpoint} -> ${betaEndpoint}`);
        return {
          success: true,
          repairedRequest: {
            ...request,
            endpoint: betaEndpoint,
          },
          strategy: this.name,
          message: `Trying beta API version: ${betaEndpoint}`,
          metadata: {
            originalEndpoint: endpoint,
            betaEndpoint,
          },
        };
      }

      // If already on beta, try v1.0
      if (endpoint.includes('/beta/')) {
        const v1Endpoint = endpoint.replace('/beta/', '/v1.0/');
        logger.info(`Version repair: ${endpoint} -> ${v1Endpoint}`);
        return {
          success: true,
          repairedRequest: {
            ...request,
            endpoint: v1Endpoint,
          },
          strategy: this.name,
          message: `Trying v1.0 API version: ${v1Endpoint}`,
          metadata: {
            originalEndpoint: endpoint,
            v1Endpoint,
          },
        };
      }

      return {
        success: false,
        strategy: this.name,
        message: 'No version repair available',
      };
    } catch (error) {
      logger.error(`Version repair strategy error: ${error}`);
      return {
        success: false,
        strategy: this.name,
        message: `Version repair failed: ${(error as Error).message}`,
      };
    }
  }
}

/**
 * Rate Limit Repair Strategy
 * Implements intelligent request queuing and backoff
 */
export class RateLimitRepairStrategy implements RepairStrategy {
  name = 'ratelimit';
  private requestQueue: Array<{
    request: RepairRequest;
    resolve: (result: RepairResult) => void;
    timestamp: number;
  }> = [];
  private processing = false;

  canRepair(error: ParsedGraphApiError): boolean {
    return error.statusCode === 429;
  }

  async repair(request: RepairRequest): Promise<RepairResult> {
    try {
      const { originalError } = request;
      const retryAfter = originalError.rawError
        ? (originalError.rawError as { retryAfter?: number }).retryAfter
        : undefined;

      // Calculate backoff delay
      const delay = retryAfter ? retryAfter * 1000 : this.calculateBackoff();

      logger.info(`Rate limit repair: Waiting ${delay}ms before retry`);

      // Queue the request for retry
      return new Promise((resolve) => {
        this.requestQueue.push({
          request,
          resolve,
          timestamp: Date.now(),
        });

        // Process queue if not already processing
        if (!this.processing) {
          this.processQueue();
        }
      });
    } catch (error) {
      logger.error(`Rate limit repair strategy error: ${error}`);
      return {
        success: false,
        strategy: this.name,
        message: `Rate limit repair failed: ${(error as Error).message}`,
      };
    }
  }

  private calculateBackoff(): number {
    // Exponential backoff: 1s, 2s, 4s, 8s, max 30s
    const baseDelay = 1000;
    const maxDelay = 30000;
    const queueLength = this.requestQueue.length;
    const delay = Math.min(baseDelay * Math.pow(2, queueLength), maxDelay);
    // Add jitter
    return delay + Math.random() * 1000;
  }

  private async processQueue(): Promise<void> {
    if (this.processing || this.requestQueue.length === 0) {
      return;
    }

    this.processing = true;

    while (this.requestQueue.length > 0) {
      const item = this.requestQueue.shift();
      if (!item) break;

      const { request, resolve } = item;
      const delay = this.calculateBackoff();

      // Wait for calculated delay
      await new Promise((r) => setTimeout(r, delay));

      // Return repaired request (same request, just delayed)
      resolve({
        success: true,
        repairedRequest: request,
        strategy: this.name,
        message: `Rate limit backoff applied, retrying after ${delay}ms`,
        metadata: {
          delay,
          queuePosition: this.requestQueue.length,
        },
      });
    }

    this.processing = false;
  }
}

/**
 * Get all available repair strategies
 */
export function getRepairStrategies(): RepairStrategy[] {
  return [
    new EndpointRepairStrategy(),
    new ParameterRepairStrategy(),
    new ScopeRepairStrategy(),
    new VersionRepairStrategy(),
    new RateLimitRepairStrategy(),
  ];
}
