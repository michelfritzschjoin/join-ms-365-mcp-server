import { describe, expect, it } from 'vitest';
import GraphClient, { type GraphBatchRequest } from '../src/graph-client.js';
import type AuthManager from '../src/auth.js';
import type { AppSecrets } from '../src/secrets.js';

describe('super-tools batching guard', () => {
  it('rejects batches larger than Graph maximum (20)', async () => {
    const mockAuthManager = {
      getToken: async () => 'token',
    } as unknown as AuthManager;

    const secrets = {
      clientId: 'test-client-id',
      tenantId: 'common',
      cloudType: 'global',
    } as AppSecrets;

    const graphClient = new GraphClient(mockAuthManager, secrets);
    const requests: GraphBatchRequest[] = Array.from({ length: 21 }, (_, index) => ({
      id: String(index + 1),
      method: 'GET',
      url: '/me',
    }));

    await expect(graphClient.performBatch(requests)).rejects.toThrow(
      'Graph batch supports at most 20 requests per call'
    );
  });
});
