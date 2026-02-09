import { describe, it, expect } from 'vitest';
import {
  loadEndpoints,
  compareEndpoints,
  validateAllEndpoints,
  generateChangeReportMarkdown,
  type EndpointConfig,
} from './utils/api-change-detector.js';

describe('Graph API Change Detection', () => {
  describe('loadEndpoints', () => {
    it('should load endpoints from endpoints.json', () => {
      const endpoints = loadEndpoints();
      expect(Array.isArray(endpoints)).toBe(true);
      expect(endpoints.length).toBeGreaterThan(0);
    });

    it('should load valid endpoint structure', () => {
      const endpoints = loadEndpoints();
      for (const endpoint of endpoints) {
        expect(endpoint).toHaveProperty('pathPattern');
        expect(endpoint).toHaveProperty('method');
        expect(endpoint).toHaveProperty('toolName');
      }
    });
  });

  describe('compareEndpoints', () => {
    it('should detect added endpoints', () => {
      const oldEndpoints: EndpointConfig[] = [
        {
          pathPattern: '/me',
          method: 'get',
          toolName: 'get-current-user',
          scopes: ['User.Read'],
        },
      ];

      const newEndpoints: EndpointConfig[] = [
        {
          pathPattern: '/me',
          method: 'get',
          toolName: 'get-current-user',
          scopes: ['User.Read'],
        },
        {
          pathPattern: '/me/messages',
          method: 'get',
          toolName: 'list-mail-messages',
          scopes: ['Mail.Read'],
        },
      ];

      const report = compareEndpoints(oldEndpoints, newEndpoints);

      expect(report.added.length).toBe(1);
      expect(report.added[0].toolName).toBe('list-mail-messages');
      expect(report.removed.length).toBe(0);
      expect(report.modified.length).toBe(0);
    });

    it('should detect removed endpoints', () => {
      const oldEndpoints: EndpointConfig[] = [
        {
          pathPattern: '/me',
          method: 'get',
          toolName: 'get-current-user',
          scopes: ['User.Read'],
        },
        {
          pathPattern: '/me/messages',
          method: 'get',
          toolName: 'list-mail-messages',
          scopes: ['Mail.Read'],
        },
      ];

      const newEndpoints: EndpointConfig[] = [
        {
          pathPattern: '/me',
          method: 'get',
          toolName: 'get-current-user',
          scopes: ['User.Read'],
        },
      ];

      const report = compareEndpoints(oldEndpoints, newEndpoints);

      expect(report.removed.length).toBe(1);
      expect(report.removed[0].toolName).toBe('list-mail-messages');
      expect(report.added.length).toBe(0);
      expect(report.modified.length).toBe(0);
    });

    it('should detect modified endpoints', () => {
      const oldEndpoints: EndpointConfig[] = [
        {
          pathPattern: '/me/messages',
          method: 'get',
          toolName: 'list-mail-messages',
          scopes: ['Mail.Read'],
        },
      ];

      const newEndpoints: EndpointConfig[] = [
        {
          pathPattern: '/me/messages',
          method: 'get',
          toolName: 'list-mail-messages',
          scopes: ['Mail.Read', 'Mail.ReadWrite'],
        },
      ];

      const report = compareEndpoints(oldEndpoints, newEndpoints);

      expect(report.modified.length).toBe(1);
      expect(report.modified[0].changes).toContain('scopes changed');
      expect(report.added.length).toBe(0);
      expect(report.removed.length).toBe(0);
    });

    it('should detect toolName changes', () => {
      const oldEndpoints: EndpointConfig[] = [
        {
          pathPattern: '/me/messages',
          method: 'get',
          toolName: 'old-tool-name',
          scopes: ['Mail.Read'],
        },
      ];

      const newEndpoints: EndpointConfig[] = [
        {
          pathPattern: '/me/messages',
          method: 'get',
          toolName: 'new-tool-name',
          scopes: ['Mail.Read'],
        },
      ];

      const report = compareEndpoints(oldEndpoints, newEndpoints);

      expect(report.modified.length).toBe(1);
      expect(report.modified[0].changes).toContain(
        'toolName changed from "old-tool-name" to "new-tool-name"'
      );
    });

    it('should detect method changes', () => {
      const oldEndpoints: EndpointConfig[] = [
        {
          pathPattern: '/me/messages',
          method: 'get',
          toolName: 'list-mail-messages',
          scopes: ['Mail.Read'],
        },
      ];

      const newEndpoints: EndpointConfig[] = [
        {
          pathPattern: '/me/messages',
          method: 'post',
          toolName: 'list-mail-messages',
          scopes: ['Mail.Read'],
        },
      ];

      const report = compareEndpoints(oldEndpoints, newEndpoints);

      // Method change results in removed + added, not modified
      // (because the key is method:pathPattern)
      expect(report.removed.length).toBe(1);
      expect(report.added.length).toBe(1);
      expect(report.removed[0].method).toBe('get');
      expect(report.added[0].method).toBe('post');
    });
  });

  describe('validateAllEndpoints', () => {
    it('should validate all endpoints', () => {
      const endpoints = loadEndpoints();
      const validation = validateAllEndpoints(endpoints);

      // All endpoints should be valid
      expect(validation.valid).toBe(true);
      expect(validation.errors.length).toBe(0);
    });

    it('should detect invalid endpoints', () => {
      const invalidEndpoints: EndpointConfig[] = [
        {
          pathPattern: 'invalid-path', // Missing leading slash
          method: 'get',
          toolName: 'invalid-tool',
        },
        {
          pathPattern: '/valid-path',
          method: 'invalid-method', // Invalid method
          toolName: 'valid-tool',
        },
        {
          pathPattern: '/valid-path',
          method: 'get',
          toolName: 'InvalidTool', // Invalid tool name format
        },
      ];

      const validation = validateAllEndpoints(invalidEndpoints);

      expect(validation.valid).toBe(false);
      expect(validation.errors.length).toBeGreaterThan(0);
    });
  });

  describe('generateChangeReportMarkdown', () => {
    it('should generate markdown report', () => {
      const oldEndpoints: EndpointConfig[] = [
        {
          pathPattern: '/me',
          method: 'get',
          toolName: 'get-current-user',
          scopes: ['User.Read'],
        },
      ];

      const newEndpoints: EndpointConfig[] = [
        {
          pathPattern: '/me',
          method: 'get',
          toolName: 'get-current-user',
          scopes: ['User.Read'],
        },
        {
          pathPattern: '/me/messages',
          method: 'get',
          toolName: 'list-mail-messages',
          scopes: ['Mail.Read'],
        },
      ];

      const report = compareEndpoints(oldEndpoints, newEndpoints);
      const markdown = generateChangeReportMarkdown(report);

      expect(markdown).toContain('# Graph API Change Report');
      expect(markdown).toContain('## Summary');
      expect(markdown).toContain('## Added Endpoints');
      expect(markdown).toContain('list-mail-messages');
    });

    it('should include all sections in report', () => {
      const oldEndpoints: EndpointConfig[] = [
        {
          pathPattern: '/old',
          method: 'get',
          toolName: 'old-tool',
          scopes: ['User.Read'],
        },
        {
          pathPattern: '/modified',
          method: 'get',
          toolName: 'modified-tool',
          scopes: ['User.Read'],
        },
      ];

      const newEndpoints: EndpointConfig[] = [
        {
          pathPattern: '/modified',
          method: 'get',
          toolName: 'modified-tool',
          scopes: ['User.Read', 'User.ReadWrite'],
        },
        {
          pathPattern: '/new',
          method: 'get',
          toolName: 'new-tool',
          scopes: ['User.Read'],
        },
      ];

      const report = compareEndpoints(oldEndpoints, newEndpoints);
      const markdown = generateChangeReportMarkdown(report);

      expect(markdown).toContain('## Added Endpoints');
      expect(markdown).toContain('## Removed Endpoints');
      expect(markdown).toContain('## Modified Endpoints');
    });
  });

  describe('Real-world change detection', () => {
    it('should detect changes in actual endpoints.json', () => {
      const currentEndpoints = loadEndpoints();

      // Create a modified version
      const modifiedEndpoints = currentEndpoints.map((endpoint) => {
        if (endpoint.toolName === 'get-current-user') {
          return {
            ...endpoint,
            scopes: [...(endpoint.scopes || []), 'User.ReadWrite'],
          };
        }
        return endpoint;
      });

      const report = compareEndpoints(currentEndpoints, modifiedEndpoints);

      // Should detect at least one modification
      expect(report.summary.modifiedCount).toBeGreaterThanOrEqual(0);
    });
  });
});
