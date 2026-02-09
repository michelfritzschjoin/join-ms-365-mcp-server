/**
 * Data Aggregator Tests
 *
 * Tests for data aggregation, deduplication, and formatting
 */

import { describe, it, expect, beforeEach } from 'vitest';
import {
  DataAggregator,
  type AggregationOptions,
  type AggregatedItem,
} from '../src/data-aggregator.js';

describe('DataAggregator', () => {
  let aggregator: DataAggregator;

  beforeEach(() => {
    aggregator = new DataAggregator();
  });

  describe('aggregate', () => {
    it('should aggregate items from multiple sources', () => {
      const dataArrays = [
        {
          source: 'source1',
          items: [
            { id: '1', name: 'Item 1' },
            { id: '2', name: 'Item 2' },
          ],
        },
        {
          source: 'source2',
          items: [{ id: '3', name: 'Item 3' }],
        },
      ];

      const result = aggregator.aggregate(dataArrays);

      expect(result.items.length).toBe(3);
      expect(result.totalItems).toBe(3);
      expect(result.uniqueItems).toBe(3);
      expect(result.sources).toContain('source1');
      expect(result.sources).toContain('source2');
    });

    it('should deduplicate items with same ID', () => {
      const dataArrays = [
        {
          source: 'source1',
          items: [
            { id: '1', name: 'Item 1' },
            { id: '2', name: 'Item 2' },
          ],
        },
        {
          source: 'source2',
          items: [
            { id: '1', name: 'Item 1 Duplicate' },
            { id: '3', name: 'Item 3' },
          ],
        },
      ];

      const result = aggregator.aggregate(dataArrays, { deduplicate: true });

      expect(result.items.length).toBe(3);
      expect(result.totalItems).toBe(3); // After deduplication, only unique items are counted
      expect(result.uniqueItems).toBe(3);
    });

    it('should not deduplicate when deduplicate option is false', () => {
      const dataArrays = [
        {
          source: 'source1',
          items: [{ id: '1', name: 'Item 1' }],
        },
        {
          source: 'source2',
          items: [{ id: '1', name: 'Item 1 Duplicate' }],
        },
      ];

      const result = aggregator.aggregate(dataArrays, { deduplicate: false });

      expect(result.items.length).toBe(2);
      expect(result.totalItems).toBe(2);
    });

    it('should sort by relevance score (descending)', () => {
      const dataArrays = [
        {
          source: 'source1',
          items: [
            { id: '1', name: 'Item 1' },
            { id: '2', name: 'Item 2' },
          ],
        },
      ];

      const result = aggregator.aggregate(dataArrays, {
        sortBy: 'relevance',
        sortOrder: 'desc',
      });

      expect(result.items.length).toBe(2);
      // Items should be sorted by relevance (descending)
      expect(result.items[0].relevanceScore).toBeGreaterThanOrEqual(result.items[1].relevanceScore);
    });

    it('should sort by timestamp (ascending)', () => {
      const dataArrays = [
        {
          source: 'source1',
          items: [
            { id: '1', name: 'Item 1', createdDateTime: '2026-01-27T12:00:00Z' },
            { id: '2', name: 'Item 2', createdDateTime: '2026-01-27T10:00:00Z' },
          ],
        },
      ];

      const result = aggregator.aggregate(dataArrays, {
        sortBy: 'timestamp',
        sortOrder: 'asc',
      });

      expect(result.items.length).toBe(2);
      const time1 = result.items[0].timestamp?.getTime() || 0;
      const time2 = result.items[1].timestamp?.getTime() || 0;
      expect(time1).toBeLessThanOrEqual(time2);
    });

    it('should sort by source (alphabetical)', () => {
      const dataArrays = [
        {
          source: 'source2',
          items: [{ id: '1', name: 'Item 1' }],
        },
        {
          source: 'source1',
          items: [{ id: '2', name: 'Item 2' }],
        },
      ];

      const result = aggregator.aggregate(dataArrays, {
        sortBy: 'source',
        sortOrder: 'asc',
      });

      expect(result.items.length).toBe(2);
      expect(result.items[0].source).toBe('source1');
      expect(result.items[1].source).toBe('source2');
    });

    it('should limit items to maxItems', () => {
      const dataArrays = [
        {
          source: 'source1',
          items: Array.from({ length: 10 }, (_, i) => ({
            id: `${i}`,
            name: `Item ${i}`,
          })),
        },
      ];

      const result = aggregator.aggregate(dataArrays, {
        maxItems: 5,
      });

      expect(result.items.length).toBe(5);
      expect(result.totalItems).toBe(10);
    });

    it('should format for LLM when requested', () => {
      const dataArrays = [
        {
          source: 'source1',
          items: [
            { id: '1', name: 'Item 1', subject: 'Test Subject' },
            { id: '2', name: 'Item 2', subject: 'Another Subject' },
          ],
        },
      ];

      const result = aggregator.aggregate(dataArrays, {
        formatForLLM: true,
      });

      expect(result.formattedForLLM).toBeDefined();
      expect(typeof result.formattedForLLM).toBe('string');
      expect(result.formattedForLLM?.length).toBeGreaterThan(0);
    });

    it('should handle items without ID fields', () => {
      const dataArrays = [
        {
          source: 'source1',
          items: [
            { name: 'Item 1', value: 'test' },
            { name: 'Item 2', value: 'test2' },
          ],
        },
      ];

      const result = aggregator.aggregate(dataArrays);

      expect(result.items.length).toBe(2);
      expect(result.items[0].id).toBeDefined();
      expect(result.items[1].id).toBeDefined();
    });

    it('should handle empty arrays', () => {
      const dataArrays: Array<{ source: string; items: unknown[] }> = [];

      const result = aggregator.aggregate(dataArrays);

      expect(result.items.length).toBe(0);
      expect(result.totalItems).toBe(0);
      expect(result.uniqueItems).toBe(0);
      expect(result.sources.length).toBe(0);
    });

    it('should handle arrays with empty items', () => {
      const dataArrays = [
        {
          source: 'source1',
          items: [],
        },
        {
          source: 'source2',
          items: [{ id: '1', name: 'Item 1' }],
        },
      ];

      const result = aggregator.aggregate(dataArrays);

      expect(result.items.length).toBe(1);
      expect(result.sources).toContain('source1');
      expect(result.sources).toContain('source2');
    });

    it('should calculate relevance scores', () => {
      const dataArrays = [
        {
          source: 'source1',
          items: [
            { id: '1', name: 'Item 1', importance: 'high' },
            { id: '2', name: 'Item 2', importance: 'normal' },
          ],
        },
      ];

      const result = aggregator.aggregate(dataArrays);

      expect(result.items[0].relevanceScore).toBeDefined();
      expect(result.items[1].relevanceScore).toBeDefined();
      expect(typeof result.items[0].relevanceScore).toBe('number');
    });

    it('should extract timestamps from various fields', () => {
      const dataArrays = [
        {
          source: 'source1',
          items: [
            { id: '1', name: 'Item 1', createdDateTime: '2026-01-27T10:00:00Z' },
            { id: '2', name: 'Item 2', lastModifiedDateTime: '2026-01-27T11:00:00Z' },
            { id: '3', name: 'Item 3', receivedDateTime: '2026-01-27T12:00:00Z' },
          ],
        },
      ];

      const result = aggregator.aggregate(dataArrays);

      expect(result.items[0].timestamp).toBeDefined();
      expect(result.items[1].timestamp).toBeDefined();
      expect(result.items[2].timestamp).toBeDefined();
    });
  });
});
