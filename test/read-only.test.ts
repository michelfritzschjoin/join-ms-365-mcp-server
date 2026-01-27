import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { parseArgs } from '../src/cli.js';
import { registerGraphTools } from '../src/graph-tools.js';
import type { GraphClient } from '../src/graph-client.js';

vi.mock('../src/cli.js', () => {
  const parseArgsMock = vi.fn();
  return {
    parseArgs: parseArgsMock,
  };
});

vi.mock('../src/generated/client.js', () => {
  return {
    api: {
      endpoints: [
        {
          alias: 'list-mail-messages',
          method: 'get',
          path: '/me/messages',
          parameters: [],
        },
        {
          alias: 'send-mail',
          method: 'post',
          path: '/me/sendMail',
          parameters: [],
        },
        {
          alias: 'delete-mail-message',
          method: 'delete',
          path: '/me/messages/{message-id}',
          parameters: [],
        },
      ],
    },
  };
});

vi.mock('../src/logger.js', () => {
  return {
    default: {
      info: vi.fn(),
      error: vi.fn(),
    },
  };
});

describe('Read-Only Mode', () => {
  let mockServer: {
    tool: ReturnType<typeof vi.fn>;
    registerTool: ReturnType<typeof vi.fn>;
  };
  let toolCalls: string[];

  beforeEach(() => {
    vi.clearAllMocks();

    delete process.env.READ_ONLY;

    toolCalls = [];
    mockServer = {
      tool: vi.fn((name: string) => {
        toolCalls.push(name);
      }),
      registerTool: vi.fn((name: string) => {
        toolCalls.push(name);
      }),
    };
  });

  afterEach(() => {
    vi.resetAllMocks();
  });

  it('should respect --read-only flag from CLI', () => {
    vi.mocked(parseArgs).mockReturnValue({ readOnly: true } as ReturnType<typeof parseArgs>);

    const options = parseArgs();
    expect(options.readOnly).toBe(true);

    registerGraphTools(mockServer, {} as GraphClient, options.readOnly);

    // Check combined calls from both tool() and registerTool()
    expect(toolCalls.length).toBe(1);
    expect(toolCalls).toContain('list-mail-messages');
    expect(toolCalls).not.toContain('send-mail');
    expect(toolCalls).not.toContain('delete-mail-message');
  });

  it('should register all endpoints when not in read-only mode', () => {
    vi.mocked(parseArgs).mockReturnValue({ readOnly: false } as ReturnType<typeof parseArgs>);

    const options = parseArgs();
    expect(options.readOnly).toBe(false);

    registerGraphTools(mockServer, {} as GraphClient, options.readOnly);

    // Check combined calls from both tool() and registerTool()
    expect(toolCalls.length).toBe(3);
    expect(toolCalls).toContain('list-mail-messages');
    expect(toolCalls).toContain('send-mail');
    expect(toolCalls).toContain('delete-mail-message');
  });
});
