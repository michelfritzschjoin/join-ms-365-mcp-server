import { describe, expect, it, vi } from 'vitest';
import { registerCompoundTools } from '../src/compound-tools.js';

vi.mock('../src/logger.js', () => {
  return {
    default: {
      info: vi.fn(),
      warn: vi.fn(),
      error: vi.fn(),
      debug: vi.fn(),
    },
  };
});

vi.mock('../src/chat-memory.js', () => {
  return {
    getChatMemoryStore: vi.fn(),
    isChatMemoryEnabled: vi.fn(() => false),
  };
});

vi.mock('../src/request-context.js', () => {
  return {
    getChatId: vi.fn(() => 'test-chat-id'),
    getUserId: vi.fn(() => 'test-user-id'),
  };
});

type ToolHandler = (params: Record<string, unknown>) => Promise<{
  content: Array<{ type: 'text'; text: string }>;
  isError?: boolean;
}>;

describe('Compound Tool: find-upcoming-meetings', () => {
  it('should roll windows until it finds events', async () => {
    const handlers = new Map<string, ToolHandler>();

    const mockServer = {
      tool: vi.fn(
        (
          name: string,
          _description: string,
          _schema: unknown,
          _meta: unknown,
          handler: ToolHandler
        ) => {
          handlers.set(name, handler);
        }
      ),
    };

    const mockGraphClient = {
      makeRequest: vi
        .fn()
        .mockResolvedValueOnce({ value: [] })
        .mockResolvedValueOnce({
          value: [
            {
              id: 'event-1',
              subject: 'Test Meeting',
              start: { dateTime: '2026-01-27T12:00:00.0000000', timeZone: 'UTC' },
              end: { dateTime: '2026-01-27T12:30:00.0000000', timeZone: 'UTC' },
              isAllDay: false,
              isCancelled: false,
              isOnlineMeeting: false,
            },
          ],
        }),
    };

    registerCompoundTools(mockServer as any, mockGraphClient as any, true);

    const handler = handlers.get('find-upcoming-meetings');
    expect(handler).toBeDefined();

    const result = await handler!({
      windowDays: 30,
      maxWindows: 2,
      limit: 50,
      includeAllDay: true,
      includeCancelled: false,
    });

    expect(result.isError).not.toBe(true);
    expect(mockGraphClient.makeRequest).toHaveBeenCalledTimes(2);

    const payload = JSON.parse(result.content[0].text) as {
      success: boolean;
      events: unknown[];
    };

    expect(payload.success).toBe(true);
    expect(Array.isArray(payload.events)).toBe(true);
    expect(payload.events).toHaveLength(1);
  });
});
