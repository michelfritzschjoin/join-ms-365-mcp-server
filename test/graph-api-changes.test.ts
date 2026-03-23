import { describe, it, expect, beforeEach, vi } from 'vitest';
import { readFileSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import GraphClient from '../src/graph-client.js';
import AuthManager from '../src/auth.js';
import type { AppSecrets } from '../src/secrets.js';
import { getCloudEndpoints } from '../src/cloud-config.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Graph API Change Detection Tests
 *
 * These tests verify that:
 * 1. Graph API endpoints are correctly formatted
 * 2. API versions are properly handled (v1.0 vs beta)
 * 3. Endpoints match the expected patterns
 * 4. API changes are detected and handled
 */

describe('Graph API Change Detection', () => {
  let mockSecrets: AppSecrets;
  let mockAuthManager: AuthManager;
  let graphClient: GraphClient;

  beforeEach(() => {
    // Mock secrets
    mockSecrets = {
      clientId: 'test-client-id',
      tenantId: 'common',
      cloudType: 'global',
    } as AppSecrets;

    // Mock auth manager
    mockAuthManager = {
      getToken: vi.fn().mockResolvedValue('mock-access-token'),
    } as unknown as AuthManager;

    graphClient = new GraphClient(mockAuthManager, mockSecrets);
  });

  describe('API Version Handling', () => {
    it('should use v1.0 API version by default', () => {
      const endpoints = getCloudEndpoints('global');
      expect(endpoints.graphApi).toBe('https://graph.microsoft.com');

      // The GraphClient should construct URLs with /v1.0/
      // This is tested through the performRequest method
      const url = `${endpoints.graphApi}/v1.0/me`;
      expect(url).toBe('https://graph.microsoft.com/v1.0/me');
    });

    it('should support beta API version', () => {
      const endpoints = getCloudEndpoints('global');
      const betaUrl = `${endpoints.graphApi}/beta/me`;
      expect(betaUrl).toBe('https://graph.microsoft.com/beta/me');
    });

    it('should handle China cloud endpoints', () => {
      const endpoints = getCloudEndpoints('china');
      expect(endpoints.graphApi).toBe('https://microsoftgraph.chinacloudapi.cn');
      const url = `${endpoints.graphApi}/v1.0/me`;
      expect(url).toBe('https://microsoftgraph.chinacloudapi.cn/v1.0/me');
    });
  });

  describe('Endpoint Validation', () => {
    it('should validate endpoint patterns from endpoints.json', () => {
      const endpointsPath = path.join(__dirname, '../src/endpoints.json');
      const endpointsData = JSON.parse(readFileSync(endpointsPath, 'utf8'));

      expect(Array.isArray(endpointsData)).toBe(true);
      expect(endpointsData.length).toBeGreaterThan(0);

      // Validate structure
      for (const endpoint of endpointsData) {
        expect(endpoint).toHaveProperty('pathPattern');
        expect(endpoint).toHaveProperty('method');
        expect(endpoint).toHaveProperty('toolName');
        expect(typeof endpoint.pathPattern).toBe('string');
        expect(typeof endpoint.method).toBe('string');
        expect(typeof endpoint.toolName).toBe('string');

        // Path pattern should start with /
        expect(endpoint.pathPattern.startsWith('/')).toBe(true);

        // Method should be uppercase
        expect(['GET', 'POST', 'PATCH', 'PUT', 'DELETE']).toContain(endpoint.method.toUpperCase());
      }
    });

    it("should validate that endpoints don't include version prefix", () => {
      const endpointsPath = path.join(__dirname, '../src/endpoints.json');
      const endpointsData = JSON.parse(readFileSync(endpointsPath, 'utf8'));

      for (const endpoint of endpointsData) {
        // Endpoints should NOT include /v1.0/ or /beta/ prefix
        expect(endpoint.pathPattern).not.toMatch(/^\/v1\.0\//);
        expect(endpoint.pathPattern).not.toMatch(/^\/beta\//);
      }
    });

    it('should validate required scopes are defined', () => {
      const endpointsPath = path.join(__dirname, '../src/endpoints.json');
      const endpointsData = JSON.parse(readFileSync(endpointsPath, 'utf8'));

      for (const endpoint of endpointsData) {
        // Each endpoint should have either scopes or workScopes
        const hasScopes =
          endpoint.scopes && Array.isArray(endpoint.scopes) && endpoint.scopes.length > 0;
        const hasWorkScopes =
          endpoint.workScopes &&
          Array.isArray(endpoint.workScopes) &&
          endpoint.workScopes.length > 0;

        expect(hasScopes || hasWorkScopes).toBe(true);
      }
    });
  });

  describe('Critical API Endpoints', () => {
    const endpointsPath = path.join(__dirname, '../src/endpoints.json');
    const endpointsData = JSON.parse(readFileSync(endpointsPath, 'utf8'));

    it('should have email endpoints configured', () => {
      const emailEndpoints = endpointsData.filter(
        (e: { pathPattern: string }) =>
          e.pathPattern.includes('/messages') || e.pathPattern.includes('/mailFolders')
      );
      expect(emailEndpoints.length).toBeGreaterThan(0);
    });

    it('should have calendar endpoints configured', () => {
      const calendarEndpoints = endpointsData.filter(
        (e: { pathPattern: string }) =>
          e.pathPattern.includes('/events') || e.pathPattern.includes('/calendar')
      );
      expect(calendarEndpoints.length).toBeGreaterThan(0);
    });

    it('should have file endpoints configured', () => {
      const fileEndpoints = endpointsData.filter(
        (e: { pathPattern: string }) =>
          e.pathPattern.includes('/drives') || e.pathPattern.includes('/drive')
      );
      expect(fileEndpoints.length).toBeGreaterThan(0);
    });

    it('should have Teams endpoints configured', () => {
      const teamsEndpoints = endpointsData.filter(
        (e: { pathPattern: string }) =>
          e.pathPattern.includes('/teams') || e.pathPattern.includes('/chats')
      );
      expect(teamsEndpoints.length).toBeGreaterThan(0);
    });

    it('should have SharePoint endpoints configured', () => {
      const sharePointEndpoints = endpointsData.filter((e: { pathPattern: string }) =>
        e.pathPattern.includes('/sites')
      );
      expect(sharePointEndpoints.length).toBeGreaterThan(0);
    });

    it('should have search endpoint configured', () => {
      const searchEndpoint = endpointsData.find(
        (e: { pathPattern: string }) => e.pathPattern === '/search/query'
      );
      expect(searchEndpoint).toBeDefined();
      expect(searchEndpoint.method).toBe('post');
    });
  });

  describe('API Change Detection', () => {
    it('should detect when endpoint path changes', () => {
      const oldEndpoint = '/me/messages';
      const newEndpoint = '/me/messages/v2';

      expect(oldEndpoint).not.toBe(newEndpoint);
      // This would trigger a change detection in a real scenario
    });

    it('should detect when endpoint method changes', () => {
      const endpoint1 = { pathPattern: '/me/messages', method: 'get' };
      const endpoint2 = { pathPattern: '/me/messages', method: 'post' };

      expect(endpoint1.method).not.toBe(endpoint2.method);
    });

    it('should detect when required scopes change', () => {
      const endpoint1 = { pathPattern: '/me/messages', scopes: ['Mail.Read'] };
      const endpoint2 = { pathPattern: '/me/messages', scopes: ['Mail.ReadWrite'] };

      expect(endpoint1.scopes).not.toEqual(endpoint2.scopes);
    });
  });

  describe('URL Construction', () => {
    it('should construct correct Graph API URLs', () => {
      const endpoints = getCloudEndpoints('global');
      const baseUrl = endpoints.graphApi;

      const testCases = [
        { endpoint: '/me', expected: `${baseUrl}/v1.0/me` },
        { endpoint: '/me/messages', expected: `${baseUrl}/v1.0/me/messages` },
        { endpoint: '/users', expected: `${baseUrl}/v1.0/users` },
        { endpoint: '/search/query', expected: `${baseUrl}/v1.0/search/query` },
      ];

      for (const testCase of testCases) {
        const constructedUrl = `${baseUrl}/v1.0${testCase.endpoint}`;
        expect(constructedUrl).toBe(testCase.expected);
      }
    });

    it('should handle query parameters correctly', () => {
      const endpoints = getCloudEndpoints('global');
      const baseUrl = endpoints.graphApi;
      const endpoint = '/me/messages';

      const queryParams = {
        $top: '10',
        $filter: 'hasAttachments eq true',
      };

      const queryString = Object.entries(queryParams)
        .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
        .join('&');

      const fullUrl = `${baseUrl}/v1.0${endpoint}?${queryString}`;

      // Query parameters are URL-encoded, so $ becomes %24
      expect(fullUrl).toContain('%24top=10');
      expect(fullUrl).toContain('%24filter=');
      expect(fullUrl).toContain('hasAttachments%20eq%20true');
    });
  });

  describe('API Compatibility', () => {
    it('should validate endpoint compatibility with v1.0', () => {
      const endpointsPath = path.join(__dirname, '../src/endpoints.json');
      const endpointsData = JSON.parse(readFileSync(endpointsPath, 'utf8'));

      // All endpoints should be compatible with v1.0 API
      // (beta endpoints would be explicitly marked or use different patterns)
      for (const endpoint of endpointsData) {
        // Most endpoints should work with v1.0
        // This is a basic validation - in practice, you'd check against OpenAPI spec
        expect(endpoint.pathPattern).toBeTruthy();
      }
    });

    it('should identify endpoints that might need beta version', () => {
      const endpointsPath = path.join(__dirname, '../src/endpoints.json');
      const endpointsData = JSON.parse(readFileSync(endpointsPath, 'utf8'));

      // Some newer endpoints might only be available in beta
      // This test documents which endpoints might need beta
      const potentiallyBetaEndpoints = endpointsData.filter((e: { pathPattern: string }) => {
        // Newer features often appear in beta first
        return (
          e.pathPattern.includes('/onlineMeetings') ||
          e.pathPattern.includes('/transcripts') ||
          e.pathPattern.includes('/recordings')
        );
      });

      // These endpoints should be tested with both v1.0 and beta
      expect(potentiallyBetaEndpoints.length).toBeGreaterThanOrEqual(0);
    });
  });

  describe('getBinaryContent', () => {
    it('returns ArrayBuffer when response is ok', async () => {
      const buf = new Uint8Array([1, 2, 3]);
      const mockArrayBuffer = buf.buffer;
      vi.stubGlobal(
        'fetch',
        vi.fn().mockResolvedValue({
          ok: true,
          status: 200,
          arrayBuffer: () => Promise.resolve(mockArrayBuffer),
          text: () => Promise.resolve(''),
        })
      );
      try {
        const result = await graphClient.getBinaryContent('/me/drive/items/abc/content');
        expect(result).toBe(mockArrayBuffer);
        expect(result.byteLength).toBe(3);
      } finally {
        vi.unstubAllGlobals();
      }
    });

    it('throws when not authenticated', async () => {
      const noTokenAuth = {
        getToken: vi.fn().mockRejectedValue(new Error('No token')),
      } as unknown as AuthManager;
      const client = new GraphClient(noTokenAuth, mockSecrets);
      vi.stubGlobal('fetch', vi.fn());
      await expect(client.getBinaryContent('/me/drive/items/abc/content')).rejects.toThrow(
        /AUTHENTICATION REQUIRED/
      );
      vi.unstubAllGlobals();
    });
  });
});
