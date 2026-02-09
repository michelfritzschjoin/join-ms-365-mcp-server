/**
 * Test Utilities and Helpers
 *
 * Common mocks, fixtures, and helper functions for tests
 */

import { vi } from 'vitest';
import type { GraphClient } from '../src/graph-client.js';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { AuthManager } from '../src/auth.js';
import type { AppSecrets } from '../src/secrets.js';

/**
 * Create a mock GraphClient
 */
export function createMockGraphClient(): Partial<GraphClient> {
  return {
    request: vi.fn(),
    get: vi.fn(),
    post: vi.fn(),
    patch: vi.fn(),
    put: vi.fn(),
    delete: vi.fn(),
  };
}

/**
 * Create a mock MCP Server
 */
export function createMockMcpServer(): Partial<McpServer> {
  const tools = new Map<string, unknown>();

  return {
    tool: vi.fn((name: string, config: unknown, handler: unknown) => {
      tools.set(name, { config, handler });
    }),
    registerTool: vi.fn((name: string, config: unknown, handler: unknown) => {
      tools.set(name, { config, handler });
    }),
    listTools: vi.fn(() => Array.from(tools.keys())),
  };
}

/**
 * Create a mock AuthManager
 */
export function createMockAuthManager(overrides?: Partial<AuthManager>): Partial<AuthManager> {
  return {
    acquireTokenByDeviceCode: vi.fn(),
    testLogin: vi.fn().mockResolvedValue({
      success: true,
      userData: { displayName: 'Test User', mail: 'test@example.com' },
    }),
    getAccessToken: vi.fn().mockResolvedValue('mock-access-token'),
    loadTokenCache: vi.fn().mockResolvedValue(undefined),
    logout: vi.fn().mockResolvedValue(undefined),
    listAccounts: vi.fn().mockResolvedValue([]),
    selectAccount: vi.fn().mockResolvedValue(true),
    removeAccount: vi.fn().mockResolvedValue(true),
    getSelectedAccountId: vi.fn().mockReturnValue('test-account-id'),
    ...overrides,
  };
}

/**
 * Create mock AppSecrets
 */
export function createMockSecrets(overrides?: Partial<AppSecrets>): AppSecrets {
  return {
    clientId: 'test-client-id',
    tenantId: 'test-tenant-id',
    clientSecret: 'test-client-secret',
    cloudType: 'public',
    ...overrides,
  };
}

/**
 * Create a mock calendar event
 */
export function createMockCalendarEvent(
  overrides?: Partial<Record<string, unknown>>
): Record<string, unknown> {
  return {
    id: 'event-123',
    subject: 'Test Meeting',
    start: {
      dateTime: '2026-01-27T10:00:00.0000000',
      timeZone: 'UTC',
    },
    end: {
      dateTime: '2026-01-27T11:00:00.0000000',
      timeZone: 'UTC',
    },
    isAllDay: false,
    location: {
      displayName: 'Conference Room A',
    },
    organizer: {
      emailAddress: {
        name: 'Organizer',
        address: 'organizer@example.com',
      },
    },
    ...overrides,
  };
}

/**
 * Create a mock mail message
 */
export function createMockMailMessage(
  overrides?: Partial<Record<string, unknown>>
): Record<string, unknown> {
  return {
    id: 'message-123',
    subject: 'Test Email',
    receivedDateTime: '2026-01-27T10:00:00.0000000Z',
    sentDateTime: '2026-01-27T09:55:00.0000000Z',
    from: {
      emailAddress: {
        name: 'Sender',
        address: 'sender@example.com',
      },
    },
    toRecipients: [
      {
        emailAddress: {
          name: 'Recipient',
          address: 'recipient@example.com',
        },
      },
    ],
    isRead: false,
    hasAttachments: false,
    importance: 'normal',
    ...overrides,
  };
}

/**
 * Create a mock drive item
 */
export function createMockDriveItem(
  overrides?: Partial<Record<string, unknown>>
): Record<string, unknown> {
  return {
    id: 'item-123',
    name: 'test-file.docx',
    webUrl: 'https://example.sharepoint.com/test-file.docx',
    size: 1024,
    lastModifiedDateTime: '2026-01-27T10:00:00.0000000Z',
    file: {
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    },
    ...overrides,
  };
}

/**
 * Wait for a specified number of milliseconds
 */
export function wait(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Reset all mocks
 */
export function resetAllMocks(): void {
  vi.clearAllMocks();
}
