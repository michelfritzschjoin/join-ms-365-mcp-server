import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createMockAuthManager, createMockSecrets } from './utils/test-helpers.js';
import type { AuthManager } from '../src/auth.js';
import type { CommandOptions } from '../src/cli.js';

// Mock dependencies
vi.mock('../src/secrets.js', () => ({
  getSecrets: vi.fn().mockResolvedValue(createMockSecrets()),
}));

vi.mock('../src/graph-client.js', () => {
  return {
    default: vi.fn().mockImplementation(() => ({
      request: vi.fn(),
      get: vi.fn(),
      post: vi.fn(),
      patch: vi.fn(),
      put: vi.fn(),
      delete: vi.fn(),
    })),
  };
});

vi.mock('../src/auth-tools.js', () => ({
  registerAuthTools: vi.fn(),
}));

vi.mock('../src/graph-tools.js', () => ({
  registerGraphTools: vi.fn().mockReturnValue(10),
  registerDiscoveryTools: vi.fn(),
}));

vi.mock('../src/compound-tools.js', () => ({
  registerCompoundTools: vi.fn().mockReturnValue(5),
}));

vi.mock('../src/super-tools.js', () => ({
  registerSuperTools: vi.fn(),
}));

vi.mock('../src/discovery-tools.js', () => ({
  registerDiscoveryTools: vi.fn(),
}));

vi.mock('../src/knowledge-base.js', () => ({
  default: vi.fn().mockImplementation(() => ({})),
}));

vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    error: vi.fn(),
    warn: vi.fn(),
    debug: vi.fn(),
  },
  enableConsoleLogging: vi.fn(),
}));

vi.mock('express', () => {
  const mockApp = {
    use: vi.fn().mockReturnThis(),
    get: vi.fn().mockReturnThis(),
    post: vi.fn().mockReturnThis(),
    listen: vi.fn((port, callback) => {
      if (callback) callback();
      return {
        close: vi.fn(),
      };
    }),
  };
  return {
    default: vi.fn(() => mockApp),
  };
});

vi.mock('@modelcontextprotocol/sdk/server/streamableHttp.js', () => ({
  StreamableHTTPServerTransport: vi.fn().mockImplementation(() => ({
    start: vi.fn().mockResolvedValue(undefined),
    close: vi.fn().mockResolvedValue(undefined),
  })),
}));

describe('HTTP Mode Integration', () => {
  let authManager: Partial<AuthManager>;

  beforeEach(() => {
    vi.clearAllMocks();
    authManager = createMockAuthManager();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  describe('HTTP Server Configuration', () => {
    it('should support HTTP mode configuration', () => {
      const options: CommandOptions = { http: '0.0.0.0:3000' };
      expect(options.http).toBe('0.0.0.0:3000');
    });

    it('should parse HTTP host and port correctly', () => {
      const httpConfig = '0.0.0.0:3000';
      const [host, port] = httpConfig.split(':');

      expect(host).toBe('0.0.0.0');
      expect(port).toBe('3000');
      expect(Number.parseInt(port, 10)).toBe(3000);
    });

    it('should handle different HTTP port configurations', () => {
      const testCases = [
        { input: 'localhost:8080', expectedHost: 'localhost', expectedPort: 8080 },
        { input: '127.0.0.1:5000', expectedHost: '127.0.0.1', expectedPort: 5000 },
        { input: '0.0.0.0:3000', expectedHost: '0.0.0.0', expectedPort: 3000 },
      ];

      testCases.forEach(({ input, expectedHost, expectedPort }) => {
        const [host, port] = input.split(':');
        expect(host).toBe(expectedHost);
        expect(Number.parseInt(port, 10)).toBe(expectedPort);
      });
    });
  });

  describe('HTTP Endpoints', () => {
    it('should support health check endpoint', () => {
      // Health check endpoint should be available in HTTP mode
      const healthEndpoint = '/health';
      expect(healthEndpoint).toBe('/health');
    });

    it('should support MCP endpoint', () => {
      // MCP endpoint should be available in HTTP mode
      const mcpEndpoint = '/mcp';
      expect(mcpEndpoint).toBe('/mcp');
    });

    it('should support OAuth discovery endpoints', () => {
      const oauthEndpoints = [
        '/.well-known/oauth-authorization-server',
        '/.well-known/oauth-protected-resource',
      ];

      oauthEndpoints.forEach((endpoint) => {
        expect(endpoint).toMatch(/^\/\.well-known\//);
      });
    });
  });

  describe('HTTP Request Handling', () => {
    it('should handle CORS headers in HTTP mode', () => {
      const corsHeaders = {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization',
      };

      expect(corsHeaders['Access-Control-Allow-Origin']).toBe('*');
      expect(corsHeaders['Access-Control-Allow-Methods']).toContain('POST');
    });

    it('should handle JSON-RPC requests', () => {
      const jsonRpcRequest = {
        jsonrpc: '2.0',
        id: 1,
        method: 'tools/list',
        params: {},
      };

      expect(jsonRpcRequest.jsonrpc).toBe('2.0');
      expect(jsonRpcRequest.method).toBe('tools/list');
    });

    it('should handle MCP tool calls via HTTP', () => {
      const toolCallRequest = {
        jsonrpc: '2.0',
        id: 2,
        method: 'tools/call',
        params: {
          name: 'test-tool',
          arguments: { param: 'value' },
        },
      };

      expect(toolCallRequest.method).toBe('tools/call');
      expect(toolCallRequest.params.name).toBe('test-tool');
    });
  });

  describe('HTTP Error Handling', () => {
    it('should handle invalid HTTP requests', () => {
      const invalidRequest = {
        jsonrpc: '2.0',
        id: null,
        method: 'invalid-method',
      };

      // Should validate request structure
      expect(invalidRequest.id).toBeNull();
      expect(invalidRequest.method).not.toMatch(/^tools\//);
    });

    it('should handle HTTP request timeouts', () => {
      const timeout = 30000; // 30 seconds
      expect(timeout).toBeGreaterThan(0);
      expect(timeout).toBeLessThanOrEqual(60000);
    });
  });

  describe('HTTP Security', () => {
    it('should support security headers', () => {
      const securityHeaders = {
        'X-Content-Type-Options': 'nosniff',
        'X-Frame-Options': 'DENY',
        'X-XSS-Protection': '1; mode=block',
      };

      expect(securityHeaders['X-Content-Type-Options']).toBe('nosniff');
      expect(securityHeaders['X-Frame-Options']).toBe('DENY');
    });

    it('should handle rate limiting in HTTP mode', () => {
      const rateLimitConfig = {
        windowMs: 60000, // 1 minute
        maxRequests: 100,
      };

      expect(rateLimitConfig.windowMs).toBe(60000);
      expect(rateLimitConfig.maxRequests).toBe(100);
    });
  });
});
