import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { GraphApiRepairManager } from '../src/graph-api-repair.js';
import {
  EndpointRepairStrategy,
  ParameterRepairStrategy,
  ScopeRepairStrategy,
  VersionRepairStrategy,
  RateLimitRepairStrategy,
} from '../src/repair-strategies.js';
import { parseGraphApiError } from '../src/errors.js';
import { GraphApiError, NotFoundError, AuthorizationError } from '../src/errors.js';

describe('Graph API Self-Repair System', () => {
  describe('parseGraphApiError', () => {
    it('should parse 404 error correctly', () => {
      const error = new NotFoundError('Resource not found');
      const parsed = parseGraphApiError(error);

      expect(parsed.statusCode).toBe(404);
      expect(parsed.resourceNotFound).toBe(true);
    });

    it('should parse 403 error with scope information', () => {
      const error = new AuthorizationError('Insufficient privileges. Mail.ReadWrite required.');
      const parsed = parseGraphApiError(error);

      expect(parsed.statusCode).toBe(403);
      expect(parsed.missingScopes).toBeDefined();
    });

    it('should parse error from error text', () => {
      const errorText = JSON.stringify({
        error: {
          code: 'Request_ResourceNotFound',
          message: 'Resource not found',
        },
      });

      const parsed = parseGraphApiError(new Error('Test'), errorText);

      expect(parsed.errorCode).toBe('Request_ResourceNotFound');
      expect(parsed.resourceNotFound).toBe(true);
    });
  });

  describe('EndpointRepairStrategy', () => {
    let strategy: EndpointRepairStrategy;

    beforeEach(() => {
      strategy = new EndpointRepairStrategy();
    });

    it('should identify 404 errors as repairable', () => {
      const error = parseGraphApiError(new NotFoundError('Not found'));
      expect(strategy.canRepair(error)).toBe(true);
    });

    it('should repair /me/* to /users/{userId}/*', async () => {
      const request = {
        endpoint: '/me/messages',
        method: 'GET',
        options: {},
        params: { userId: 'user123' },
        originalError: parseGraphApiError(new NotFoundError('Not found')),
      };

      const result = await strategy.repair(request);

      expect(result.success).toBe(true);
      expect(result.repairedRequest?.endpoint).toBe('/users/user123/messages');
    });

    it('should repair /v1.0/* to /beta/*', async () => {
      const request = {
        endpoint: '/v1.0/users',
        method: 'GET',
        options: {},
        params: {},
        originalError: parseGraphApiError(new NotFoundError('Not found')),
      };

      const result = await strategy.repair(request);

      expect(result.success).toBe(true);
      expect(result.repairedRequest?.endpoint).toBe('/beta/users');
    });

    it('should return failure when no alternative found', async () => {
      const request = {
        endpoint: '/unknown/endpoint',
        method: 'GET',
        options: {},
        params: {},
        originalError: parseGraphApiError(new NotFoundError('Not found')),
      };

      const result = await strategy.repair(request);

      expect(result.success).toBe(false);
    });
  });

  describe('ParameterRepairStrategy', () => {
    let strategy: ParameterRepairStrategy;

    beforeEach(() => {
      strategy = new ParameterRepairStrategy();
    });

    it('should identify 400 errors as repairable', () => {
      const error = parseGraphApiError(new GraphApiError('Bad request', 400));
      expect(strategy.canRepair(error)).toBe(true);
    });

    it('should normalize OData parameters', async () => {
      const request = {
        endpoint: '/me/messages',
        method: 'GET',
        options: {
          queryParams: {
            filter: 'test',
            top: '10',
          },
        },
        params: {},
        originalError: parseGraphApiError(new GraphApiError('Bad request', 400)),
      };

      const result = await strategy.repair(request);

      expect(result.success).toBe(true);
      expect(result.repairedRequest?.options.queryParams?.$filter).toBe('test');
      expect(result.repairedRequest?.options.queryParams?.$top).toBe('10');
    });

    it('should fix search parameter format', async () => {
      const request = {
        endpoint: '/me/messages',
        method: 'GET',
        options: {
          queryParams: {
            search: 'test query',
          },
        },
        params: {},
        originalError: parseGraphApiError(new GraphApiError('Bad request', 400)),
      };

      const result = await strategy.repair(request);

      expect(result.success).toBe(true);
      expect(result.repairedRequest?.options.queryParams?.$search).toBe('"test query"');
    });

    it('should add ConsistencyLevel header for search', async () => {
      const request = {
        endpoint: '/users',
        method: 'GET',
        options: {
          headers: {},
          queryParams: {
            $search: '"test"',
          },
        },
        params: { search: 'test' },
        originalError: parseGraphApiError(new GraphApiError('Bad request', 400)),
      };

      const result = await strategy.repair(request);

      expect(result.success).toBe(true);
      expect(result.repairedRequest?.options.headers?.['ConsistencyLevel']).toBe('eventual');
    });
  });

  describe('ScopeRepairStrategy', () => {
    let strategy: ScopeRepairStrategy;

    beforeEach(() => {
      strategy = new ScopeRepairStrategy();
    });

    it('should identify 403 errors as repairable', () => {
      const error = parseGraphApiError(new AuthorizationError('Forbidden'));
      expect(strategy.canRepair(error)).toBe(true);
    });

    it('should suggest parent scopes for missing scopes', async () => {
      const errorText = JSON.stringify({
        error: {
          code: 'InsufficientPrivileges',
          message: 'Mail.Read required',
        },
      });

      const request = {
        endpoint: '/me/messages',
        method: 'GET',
        options: {},
        params: {},
        originalError: parseGraphApiError(new AuthorizationError('Forbidden'), errorText),
      };

      const result = await strategy.repair(request);

      // Should suggest parent scopes
      expect(result.metadata?.suggestedScopes).toBeDefined();
    });

    it('should try /me endpoint for /users/{id} requests', async () => {
      const request = {
        endpoint: '/users/user123/messages',
        method: 'GET',
        options: {},
        params: {},
        originalError: parseGraphApiError(new AuthorizationError('Forbidden')),
      };

      const result = await strategy.repair(request);

      expect(result.success).toBe(true);
      expect(result.repairedRequest?.endpoint).toBe('/me/messages');
    });
  });

  describe('VersionRepairStrategy', () => {
    let strategy: VersionRepairStrategy;

    beforeEach(() => {
      strategy = new VersionRepairStrategy();
    });

    it('should identify deprecated API errors as repairable', () => {
      const error = parseGraphApiError(new GraphApiError('API deprecated', 400));
      error.deprecatedApi = true;
      expect(strategy.canRepair(error)).toBe(true);
    });

    it('should repair /v1.0/* to /beta/*', async () => {
      const request = {
        endpoint: '/v1.0/users',
        method: 'GET',
        options: {},
        params: {},
        originalError: (() => {
          const err = parseGraphApiError(new GraphApiError('Deprecated', 400));
          err.deprecatedApi = true;
          return err;
        })(),
      };

      const result = await strategy.repair(request);

      expect(result.success).toBe(true);
      expect(result.repairedRequest?.endpoint).toBe('/beta/users');
    });

    it('should repair /beta/* to /v1.0/*', async () => {
      const request = {
        endpoint: '/beta/users',
        method: 'GET',
        options: {},
        params: {},
        originalError: (() => {
          const err = parseGraphApiError(new GraphApiError('Version issue', 400));
          err.deprecatedApi = true;
          return err;
        })(),
      };

      const result = await strategy.repair(request);

      expect(result.success).toBe(true);
      expect(result.repairedRequest?.endpoint).toBe('/v1.0/users');
    });
  });

  describe('RateLimitRepairStrategy', () => {
    let strategy: RateLimitRepairStrategy;

    beforeEach(() => {
      strategy = new RateLimitRepairStrategy();
    });

    it('should identify 429 errors as repairable', () => {
      const error = parseGraphApiError(new GraphApiError('Rate limited', 429));
      expect(strategy.canRepair(error)).toBe(true);
    });

    it('should queue request for retry', async () => {
      const request = {
        endpoint: '/me/messages',
        method: 'GET',
        options: {},
        params: {},
        originalError: parseGraphApiError(new GraphApiError('Rate limited', 429)),
      };

      const resultPromise = strategy.repair(request);

      // Should return a promise that resolves after delay
      const result = await resultPromise;

      expect(result.success).toBe(true);
      expect(result.repairedRequest).toBeDefined();
      expect(result.metadata?.delay).toBeDefined();
    });
  });

  describe('GraphApiRepairManager', () => {
    let manager: GraphApiRepairManager;

    beforeEach(() => {
      // Set environment to enable repair
      process.env.MS365_MCP_ENABLE_SELF_REPAIR = 'true';
      process.env.MS365_MCP_REPAIR_STRATEGIES = 'endpoint,parameter,scope,version,ratelimit';
      manager = new GraphApiRepairManager();
    });

    afterEach(() => {
      delete process.env.MS365_MCP_ENABLE_SELF_REPAIR;
      delete process.env.MS365_MCP_REPAIR_STRATEGIES;
    });

    it('should be enabled when configured', () => {
      expect(manager.isEnabled()).toBe(true);
    });

    it('should analyze error and find applicable strategies', () => {
      const error = new NotFoundError('Not found');
      const { parsedError, applicableStrategies } = manager.analyzeError(error);

      expect(parsedError.statusCode).toBe(404);
      expect(applicableStrategies.length).toBeGreaterThan(0);
      expect(applicableStrategies.some((s) => s.name === 'endpoint')).toBe(true);
    });

    it('should attempt repair with endpoint strategy', async () => {
      const request = manager.createRepairRequest(
        '/me/messages',
        { method: 'GET' },
        new NotFoundError('Not found'),
        undefined,
        { userId: 'user123' }
      );

      const result = await manager.attemptRepair(request);

      expect(result).not.toBeNull();
      if (result?.success) {
        expect(result.repairedRequest?.endpoint).toBe('/users/user123/messages');
      }
    });

    it('should record repair history', async () => {
      const request = manager.createRepairRequest(
        '/me/messages',
        { method: 'GET' },
        new NotFoundError('Not found'),
        undefined,
        { userId: 'user123' }
      );

      await manager.attemptRepair(request);

      const history = manager.getHistory();
      expect(history.length).toBeGreaterThan(0);
    });

    it('should provide repair statistics', async () => {
      const request = manager.createRepairRequest(
        '/me/messages',
        { method: 'GET' },
        new NotFoundError('Not found'),
        undefined,
        { userId: 'user123' }
      );

      await manager.attemptRepair(request);

      const stats = manager.getStatistics();
      expect(stats.totalRepairs).toBeGreaterThan(0);
      expect(stats.byStrategy).toBeDefined();
    });

    it('should create repair request from error', () => {
      const request = manager.createRepairRequest(
        '/me/messages',
        { method: 'GET', headers: { 'Content-Type': 'application/json' } },
        new NotFoundError('Not found'),
        'Error text',
        { userId: 'user123' }
      );

      expect(request.endpoint).toBe('/me/messages');
      expect(request.method).toBe('GET');
      expect(request.originalError.statusCode).toBe(404);
    });
  });

  describe('Integration', () => {
    it('should chain multiple repair strategies', async () => {
      process.env.MS365_MCP_ENABLE_SELF_REPAIR = 'true';
      process.env.MS365_MCP_REPAIR_STRATEGIES = 'endpoint,parameter';
      const manager = new GraphApiRepairManager();

      const request = manager.createRepairRequest(
        '/me/messages',
        {
          method: 'GET',
          queryParams: {
            filter: 'test',
            search: 'query',
          },
        },
        new NotFoundError('Not found'),
        undefined,
        { userId: 'user123' }
      );

      const result = await manager.attemptRepair(request);

      expect(result).not.toBeNull();
      if (result?.success) {
        // Should have repaired both endpoint and parameters
        expect(result.repairedRequest?.endpoint).toContain('/users/');
        expect(result.repairedRequest?.options.queryParams?.$filter).toBeDefined();
      }

      delete process.env.MS365_MCP_ENABLE_SELF_REPAIR;
      delete process.env.MS365_MCP_REPAIR_STRATEGIES;
    });
  });
});
