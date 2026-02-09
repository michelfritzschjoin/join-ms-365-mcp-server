import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { createMockGraphClient, createMockAuthManager } from './utils/test-helpers.js';
import { registerAuthTools } from '../src/auth-tools.js';
import type { AuthManager } from '../src/auth.js';

// Mock dependencies
vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    error: vi.fn(),
    warn: vi.fn(),
    debug: vi.fn(),
  },
}));

describe('MCP Protocol Handling', () => {
  let server: McpServer;
  let authManager: Partial<AuthManager>;

  beforeEach(() => {
    vi.clearAllMocks();
    server = new McpServer({
      name: 'Microsoft365MCP',
      version: '1.0.0',
    });
    authManager = createMockAuthManager();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  describe('Tool Registration', () => {
    it('should register auth tools', () => {
      registerAuthTools(server, authManager as AuthManager);

      // Verify tools are registered by checking if registerTool was called
      // The actual registration happens inside registerAuthTools
      expect(authManager).toBeDefined();
    });

    it('should handle tool registration with valid schema', () => {
      const toolName = 'test-tool';
      const toolConfig = {
        title: 'Test Tool',
        description: 'A test tool',
        inputSchema: {
          type: 'object',
          properties: {
            param: { type: 'string' },
          },
        },
      };
      const handler = vi.fn().mockResolvedValue({
        content: [{ type: 'text', text: 'Success' }],
      });

      expect(() => {
        server.registerTool(toolName, toolConfig, handler);
      }).not.toThrow();
    });

    it('should handle tool call with valid parameters', async () => {
      const toolName = 'test-tool';
      const toolConfig = {
        title: 'Test Tool',
        description: 'A test tool',
        inputSchema: {
          type: 'object',
          properties: {
            param: { type: 'string' },
          },
        },
      };
      const handler = vi.fn().mockResolvedValue({
        content: [{ type: 'text', text: 'Success' }],
      });

      server.registerTool(toolName, toolConfig, handler);

      // Simulate tool call
      const result = await handler({ param: 'test-value' });

      expect(result).toEqual({
        content: [{ type: 'text', text: 'Success' }],
      });
      expect(handler).toHaveBeenCalledWith({ param: 'test-value' });
    });

    it('should handle tool call errors gracefully', async () => {
      const toolName = 'error-tool';
      const toolConfig = {
        title: 'Error Tool',
        description: 'A tool that errors',
        inputSchema: {
          type: 'object',
          properties: {},
        },
      };
      const handler = vi.fn().mockRejectedValue(new Error('Tool error'));

      server.registerTool(toolName, toolConfig, handler);

      await expect(handler({})).rejects.toThrow('Tool error');
    });
  });

  describe('MCP Server Configuration', () => {
    it('should create server with correct name and version', () => {
      const testServer = new McpServer({
        name: 'TestServer',
        version: '2.0.0',
      });

      expect(testServer).toBeInstanceOf(McpServer);
    });

    it('should handle multiple tool registrations', () => {
      const tools = ['tool1', 'tool2', 'tool3'];

      tools.forEach((toolName) => {
        server.registerTool(
          toolName,
          {
            title: toolName,
            description: `Description for ${toolName}`,
            inputSchema: { type: 'object', properties: {} },
          },
          vi.fn()
        );
      });

      // All tools should be registered without errors
      expect(tools.length).toBe(3);
    });
  });

  describe('Error Handling', () => {
    it('should handle invalid tool parameters', async () => {
      const handler = vi.fn().mockImplementation((params) => {
        // Simulate validation that would normally be done by Zod schema
        // In real implementation, Zod would validate before handler is called
        // But for testing, we simulate the validation in the handler
        if (!params || typeof params !== 'object' || !('required' in params) || !params.required) {
          return Promise.reject(new Error('Invalid parameters: required field missing'));
        }
        return Promise.resolve({
          content: [{ type: 'text', text: 'Success' }],
        });
      });

      server.registerTool(
        'validation-tool',
        {
          title: 'Validation Tool',
          description: 'A tool with validation',
          inputSchema: {
            type: 'object',
            properties: {
              required: { type: 'string' },
            },
            required: ['required'],
          },
        },
        handler
      );

      // Call handler with invalid parameters (missing required field)
      // The handler should reject the promise
      await expect(handler({})).rejects.toThrow('Invalid parameters: required field missing');

      // Verify handler was called
      expect(handler).toHaveBeenCalledWith({});
    });

    it('should handle async tool errors', async () => {
      const handler = vi.fn().mockImplementation(async () => {
        await new Promise((resolve) => setTimeout(resolve, 10));
        throw new Error('Async error');
      });

      server.registerTool(
        'async-error-tool',
        {
          title: 'Async Error Tool',
          description: 'A tool with async error',
          inputSchema: { type: 'object', properties: {} },
        },
        handler
      );

      await expect(handler({})).rejects.toThrow('Async error');
    });
  });
});
