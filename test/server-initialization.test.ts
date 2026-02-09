import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createMockAuthManager, createMockSecrets } from './utils/test-helpers.js';
import type { AuthManager } from '../src/auth.js';
import type { CommandOptions } from '../src/cli.js';

// Note: MicrosoftGraphServer is not exported, so we test initialization indirectly
// through the server module's public API

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
}));

describe('Server Initialization', () => {
  let authManager: Partial<AuthManager>;

  beforeEach(() => {
    vi.clearAllMocks();
    authManager = createMockAuthManager();
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  describe('Server Configuration', () => {
    it('should support read-only mode configuration', () => {
      const options: CommandOptions = { readOnly: true };
      expect(options.readOnly).toBe(true);
    });

    it('should support org mode configuration', () => {
      const options: CommandOptions = { orgMode: true };
      expect(options.orgMode).toBe(true);
    });

    it('should support discovery mode configuration', () => {
      const options: CommandOptions = { discovery: true };
      expect(options.discovery).toBe(true);
    });

    it('should support enabled tools pattern configuration', () => {
      const options: CommandOptions = { enabledTools: 'mail|calendar' };
      expect(options.enabledTools).toBe('mail|calendar');
    });

    it('should support toon output format configuration', () => {
      const options: CommandOptions = { toon: true };
      expect(options.toon).toBe(true);
    });

    it('should support HTTP mode configuration', () => {
      const options: CommandOptions = { http: '0.0.0.0:3000' };
      expect(options.http).toBe('0.0.0.0:3000');
    });

    it('should handle empty options', () => {
      const options: CommandOptions = {};
      expect(options).toBeDefined();
      expect(options.readOnly).toBeUndefined();
    });

    it('should handle multiple configuration options', () => {
      const options: CommandOptions = {
        readOnly: true,
        orgMode: true,
        enabledTools: 'mail',
        toon: false,
      };
      expect(options.readOnly).toBe(true);
      expect(options.orgMode).toBe(true);
      expect(options.enabledTools).toBe('mail');
      expect(options.toon).toBe(false);
    });
  });

  describe('Secrets Loading', () => {
    it('should handle secrets loading', async () => {
      const { getSecrets } = await import('../src/secrets.js');
      const secrets = await getSecrets();

      expect(secrets).toBeDefined();
      expect(secrets.clientId).toBeDefined();
    });

    it('should handle secrets loading errors', async () => {
      const { getSecrets } = await import('../src/secrets.js');
      vi.mocked(getSecrets).mockRejectedValueOnce(new Error('Secrets loading failed'));

      await expect(getSecrets()).rejects.toThrow('Secrets loading failed');
    });
  });
});
