import { describe, it, expect, vi } from 'vitest';
import { fetchAllODataPages, type GraphClientLike } from '../../src/utils/graph-pagination.js';

describe('fetchAllODataPages', () => {
  it('returns parsed object as-is when no value array', async () => {
    const client: GraphClientLike = {
      makeRequest: vi.fn(),
    };
    const input = { foo: 'bar' };
    const result = await fetchAllODataPages(client, input);
    expect(result).toEqual({ foo: 'bar' });
    expect(client.makeRequest).not.toHaveBeenCalled();
  });

  it('returns as-is when value is not array', async () => {
    const client: GraphClientLike = { makeRequest: vi.fn() };
    const input = { value: 'not-array' };
    const result = await fetchAllODataPages(client, input);
    expect(result.value).toBe('not-array');
    expect(client.makeRequest).not.toHaveBeenCalled();
  });

  it('returns combined result when nextLink is present and mock returns more items', async () => {
    const client: GraphClientLike = {
      makeRequest: vi.fn().mockResolvedValue({
        value: [{ id: '2' }],
        '@odata.nextLink': undefined,
      }),
    };
    const input = {
      value: [{ id: '1' }],
      '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/drive/items?$skiptoken=abc',
    };
    const thinking: string[] = [];
    const result = await fetchAllODataPages(client, input, 5, thinking);
    expect(result.value).toHaveLength(2);
    expect((result.value as { id: string }[]).map((x) => x.id)).toEqual(['1', '2']);
    expect(result['@odata.nextLink']).toBeUndefined();
    expect(thinking.length).toBeGreaterThan(0);
    expect(client.makeRequest).toHaveBeenCalledTimes(1);
  });

  it('parses string response when first response is string', async () => {
    const client: GraphClientLike = { makeRequest: vi.fn() };
    const result = await fetchAllODataPages(client, JSON.stringify({ value: [] }));
    expect(result.value).toEqual([]);
    expect(client.makeRequest).not.toHaveBeenCalled();
  });
});
