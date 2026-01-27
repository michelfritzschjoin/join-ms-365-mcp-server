import { beforeEach, describe, expect, it, vi } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { registerGraphTools } from '../src/graph-tools.js';
import GraphClient from '../src/graph-client.js';

vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    error: vi.fn(),
  },
}));

vi.mock('../src/generated/client.js', () => ({
  api: {
    endpoints: [
      {
        alias: 'list-mail-messages',
        method: 'GET',
        path: '/me/messages',
        description: 'List mail messages',
      },
      { alias: 'send-mail', method: 'POST', path: '/me/sendMail', description: 'Send mail' },
      {
        alias: 'list-calendar-events',
        method: 'GET',
        path: '/me/events',
        description: 'List calendar events',
      },
      {
        alias: 'list-excel-worksheets',
        method: 'GET',
        path: '/workbook/worksheets',
        description: 'List Excel worksheets',
      },
      { alias: 'get-current-user', method: 'GET', path: '/me', description: 'Get current user' },
    ],
  },
}));

describe('Tool Filtering', () => {
  let server: McpServer;
  let graphClient: GraphClient;
  let toolCalls: string[];

  beforeEach(() => {
    server = new McpServer({ name: 'test', version: '1.0.0' });
    graphClient = {} as GraphClient;
    toolCalls = [];

    // Mock both tool() and registerTool() to capture tool names
    vi.spyOn(server, 'tool').mockImplementation((name: string) => {
      toolCalls.push(name);
    });
    vi.spyOn(server, 'registerTool').mockImplementation((name: string) => {
      toolCalls.push(name);
    });
  });

  it('should register all tools when no filter is provided', () => {
    registerGraphTools(server, graphClient, false);

    expect(toolCalls.length).toBe(5);
    expect(toolCalls).toContain('list-mail-messages');
    expect(toolCalls).toContain('send-mail');
    expect(toolCalls).toContain('list-calendar-events');
    expect(toolCalls).toContain('list-excel-worksheets');
    expect(toolCalls).toContain('get-current-user');
  });

  it('should filter tools by regex pattern - mail only', () => {
    registerGraphTools(server, graphClient, false, 'mail');

    expect(toolCalls.length).toBe(2);
    expect(toolCalls).toContain('list-mail-messages');
    expect(toolCalls).toContain('send-mail');
  });

  it('should filter tools by regex pattern - calendar or excel', () => {
    registerGraphTools(server, graphClient, false, 'calendar|excel');

    expect(toolCalls.length).toBe(2);
    expect(toolCalls).toContain('list-calendar-events');
    expect(toolCalls).toContain('list-excel-worksheets');
  });

  it('should handle invalid regex patterns gracefully', () => {
    registerGraphTools(server, graphClient, false, '[invalid regex');

    expect(toolCalls.length).toBe(5);
  });

  it('should combine read-only and filtering correctly', () => {
    registerGraphTools(server, graphClient, true, 'mail');

    expect(toolCalls.length).toBe(1);
    expect(toolCalls).toContain('list-mail-messages');
  });

  it('should register no tools when pattern matches nothing', () => {
    registerGraphTools(server, graphClient, false, 'nonexistent');

    expect(toolCalls.length).toBe(0);
  });
});
