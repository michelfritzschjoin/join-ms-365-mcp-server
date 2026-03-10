/**
 * Data Aggregator for deduplication, sorting, relevance scoring, and LLM-friendly formatting
 */

import logger from './logger.js';
import { getMaxAggregateItems } from './perf-config.js';

export interface AggregatedItem {
  id: string;
  data: unknown;
  relevanceScore: number;
  source: string;
  timestamp?: Date;
  duplicateOf?: string;
}

export interface AggregationOptions {
  sortBy?: 'relevance' | 'timestamp' | 'source';
  sortOrder?: 'asc' | 'desc';
  maxItems?: number;
  deduplicate?: boolean;
  formatForLLM?: boolean;
}

export interface AggregationResult {
  items: AggregatedItem[];
  totalItems: number;
  uniqueItems: number;
  sources: string[];
  formattedForLLM?: string;
}

export class DataAggregator {
  /**
   * Aggregate data from multiple sources
   */
  aggregate(
    dataArrays: Array<{ source: string; items: unknown[] }>,
    options: AggregationOptions = {}
  ): AggregationResult {
    const {
      sortBy = 'relevance',
      sortOrder = 'desc',
      maxItems = getMaxAggregateItems(),
      deduplicate = true,
      formatForLLM = false,
    } = options;

    logger.info(`Aggregating ${dataArrays.length} data sources`);

    // 1. Flatten and deduplicate
    const aggregated: AggregatedItem[] = [];
    const seenIds = new Set<string>();
    const sources = new Set<string>();

    for (const { source, items } of dataArrays) {
      sources.add(source);

      for (const item of items) {
        const id = this.extractId(item);
        const normalizedId = this.normalizeId(id, item);

        // Deduplicate
        if (deduplicate && seenIds.has(normalizedId)) {
          // Mark as duplicate
          const existing = aggregated.find((a) => a.id === normalizedId);
          if (existing) {
            existing.duplicateOf = normalizedId;
          }
          continue;
        }

        seenIds.add(normalizedId);

        // Calculate relevance score
        const relevanceScore = this.calculateRelevanceScore(item, source);

        // Extract timestamp
        const timestamp = this.extractTimestamp(item);

        aggregated.push({
          id: normalizedId,
          data: item,
          relevanceScore,
          source,
          timestamp,
        });
      }
    }

    // 2. Sort
    aggregated.sort((a, b) => {
      let comparison = 0;

      switch (sortBy) {
        case 'relevance':
          comparison = a.relevanceScore - b.relevanceScore;
          break;
        case 'timestamp': {
          const aTime = a.timestamp?.getTime() || 0;
          const bTime = b.timestamp?.getTime() || 0;
          comparison = aTime - bTime;
          break;
        }
        case 'source':
          comparison = a.source.localeCompare(b.source);
          break;
      }

      return sortOrder === 'asc' ? comparison : -comparison;
    });

    // 3. Limit items
    const limited = aggregated.slice(0, maxItems);

    // 4. Format for LLM if requested
    let formattedForLLM: string | undefined;
    if (formatForLLM) {
      formattedForLLM = this.formatForLLM(limited);
    }

    logger.info(
      `Aggregation complete: ${limited.length} items from ${sources.size} sources (${aggregated.length - limited.length} duplicates removed)`
    );

    return {
      items: limited,
      totalItems: aggregated.length,
      uniqueItems: limited.length,
      sources: Array.from(sources),
      formattedForLLM,
    };
  }

  /**
   * Extract ID from item
   */
  private extractId(item: unknown): string {
    if (typeof item === 'object' && item !== null) {
      const obj = item as Record<string, unknown>;
      const idFields = ['id', 'objectId', 'itemId', 'messageId', 'eventId', 'siteId', 'driveId'];
      for (const field of idFields) {
        if (typeof obj[field] === 'string') {
          return obj[field] as string;
        }
      }
    }
    return JSON.stringify(item).substring(0, 50);
  }

  /**
   * Normalize ID for deduplication
   */
  private normalizeId(id: string, item: unknown): string {
    // Use ID if available, otherwise create hash from key fields
    if (id && id.length > 0) {
      return id.toLowerCase();
    }

    // Create normalized ID from key fields
    if (typeof item === 'object' && item !== null) {
      const obj = item as Record<string, unknown>;
      const keyFields = ['name', 'title', 'subject', 'displayName', 'webUrl'];
      const keyValues: string[] = [];

      for (const field of keyFields) {
        if (typeof obj[field] === 'string') {
          keyValues.push((obj[field] as string).toLowerCase().trim());
        }
      }

      if (keyValues.length > 0) {
        return keyValues.join('|');
      }
    }

    // Fallback: hash of JSON
    return JSON.stringify(item).substring(0, 100);
  }

  /**
   * Calculate relevance score
   */
  private calculateRelevanceScore(item: unknown, source: string): number {
    let score = 0.5; // Base score

    if (typeof item === 'object' && item !== null) {
      const obj = item as Record<string, unknown>;

      // Higher score if has name/title
      if (obj['name'] || obj['title'] || obj['displayName']) {
        score += 0.2;
      }

      // Higher score if has content/body
      if (obj['content'] || obj['body']) {
        score += 0.1;
      }

      // Higher score if has webUrl (accessible)
      if (obj['webUrl']) {
        score += 0.1;
      }

      // Higher score if recent (within last 30 days)
      const timestamp = this.extractTimestamp(item);
      if (timestamp) {
        const daysSince = (Date.now() - timestamp.getTime()) / (1000 * 60 * 60 * 24);
        if (daysSince < 30) {
          score += 0.1;
        }
      }

      // Source-specific scoring
      if (source === 'sharepoint' || source === 'onedrive') {
        score += 0.1; // Files are often more relevant
      }
    }

    return Math.min(score, 1.0);
  }

  /**
   * Extract timestamp from item
   */
  private extractTimestamp(item: unknown): Date | undefined {
    if (typeof item === 'object' && item !== null) {
      const obj = item as Record<string, unknown>;
      const timeFields = [
        'createdDateTime',
        'lastModifiedDateTime',
        'start',
        'end',
        'sentDateTime',
        'receivedDateTime',
      ];

      for (const field of timeFields) {
        if (typeof obj[field] === 'string') {
          const date = new Date(obj[field] as string);
          if (!isNaN(date.getTime())) {
            return date;
          }
        }
      }
    }

    return undefined;
  }

  /**
   * Format data for LLM consumption
   */
  private formatForLLM(items: AggregatedItem[]): string {
    const sections: string[] = [];

    // Group by source
    const bySource = new Map<string, AggregatedItem[]>();
    for (const item of items) {
      if (!bySource.has(item.source)) {
        bySource.set(item.source, []);
      }
      bySource.get(item.source)!.push(item);
    }

    // Format each source
    for (const [source, sourceItems] of bySource.entries()) {
      sections.push(`## ${source.toUpperCase()} (${sourceItems.length} items)`);

      for (const item of sourceItems.slice(0, 20)) {
        // Limit to 20 items per source
        const summary = this.summarizeItem(item.data);
        sections.push(`- ${summary}`);
      }

      if (sourceItems.length > 20) {
        sections.push(`... and ${sourceItems.length - 20} more items`);
      }
    }

    return sections.join('\n\n');
  }

  /**
   * Summarize item for LLM
   */
  private summarizeItem(item: unknown): string {
    if (typeof item === 'string') {
      return item.substring(0, 200);
    }

    if (typeof item === 'object' && item !== null) {
      const obj = item as Record<string, unknown>;
      const summaryFields = ['name', 'title', 'subject', 'displayName', 'content', 'body'];

      for (const field of summaryFields) {
        if (typeof obj[field] === 'string') {
          const value = obj[field] as string;
          return value.substring(0, 200);
        }
      }

      // Fallback: first few fields
      const keys = Object.keys(obj).slice(0, 3);
      return keys.map((k) => `${k}: ${String(obj[k]).substring(0, 50)}`).join(', ');
    }

    return String(item).substring(0, 200);
  }

  /**
   * Identify important documents based on relevance score
   * Returns top N documents sorted by relevance
   */
  identifyImportantDocuments(
    items: AggregatedItem[],
    threshold: number = 0.7,
    maxItems: number = 10
  ): AggregatedItem[] {
    // Filter items that are documents (files, driveItems, listItems)
    const documentItems = items.filter((item) => {
      if (typeof item.data === 'object' && item.data !== null) {
        const obj = item.data as Record<string, unknown>;
        const entityType = obj['@odata.type'] as string | undefined;
        return (
          entityType?.includes('driveItem') ||
          entityType?.includes('listItem') ||
          obj['file'] !== undefined ||
          obj['webUrl']?.toString().includes('/sites/') ||
          obj['webUrl']?.toString().includes('/drives/')
        );
      }
      return false;
    });

    // Sort by relevance score (descending)
    documentItems.sort((a, b) => b.relevanceScore - a.relevanceScore);

    // Filter by threshold and limit to maxItems
    return documentItems.filter((item) => item.relevanceScore >= threshold).slice(0, maxItems);
  }

  /**
   * Remove duplicates based on similarity
   */
  deduplicateBySimilarity(items: AggregatedItem[], similarityThreshold = 0.8): AggregatedItem[] {
    const unique: AggregatedItem[] = [];
    const seen = new Set<string>();

    for (const item of items) {
      const key = this.getSimilarityKey(item);
      if (!seen.has(key)) {
        seen.add(key);
        unique.push(item);
      }
    }

    return unique;
  }

  /**
   * Get similarity key for deduplication
   */
  private getSimilarityKey(item: AggregatedItem): string {
    if (typeof item.data === 'object' && item.data !== null) {
      const obj = item.data as Record<string, unknown>;
      const keyFields = ['name', 'title', 'subject', 'displayName', 'webUrl'];
      const values: string[] = [];

      for (const field of keyFields) {
        if (typeof obj[field] === 'string') {
          values.push((obj[field] as string).toLowerCase().trim());
        }
      }

      return values.join('|');
    }

    return item.id;
  }
}

export default DataAggregator;
