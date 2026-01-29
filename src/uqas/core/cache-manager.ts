/**
 * Cache Manager
 *
 * Language-aware caching for UQAS responses.
 * Reduces API calls for repeated or similar queries.
 */

import crypto from 'crypto';

/**
 * Cache entry
 */
export interface CacheEntry<T = unknown> {
  key: string;
  value: T;
  language: 'de' | 'en';
  createdAt: Date;
  expiresAt: Date;
  hits: number;
  queryHash: string;
}

/**
 * Cache configuration
 */
export interface CacheConfig {
  /** Enable caching */
  enabled?: boolean;
  /** Default TTL in seconds */
  defaultTTL?: number;
  /** Maximum entries */
  maxEntries?: number;
  /** Enable similar query matching */
  similarityMatching?: boolean;
  /** Similarity threshold (0-1) */
  similarityThreshold?: number;
}

/**
 * Cache statistics
 */
export interface CacheStats {
  hits: number;
  misses: number;
  entries: number;
  hitRate: number;
  oldestEntry: Date | null;
  newestEntry: Date | null;
}

/**
 * CacheManager - Language-aware response caching
 */
export class CacheManager<T = unknown> {
  private cache: Map<string, CacheEntry<T>> = new Map();
  private config: Required<CacheConfig>;
  private stats: { hits: number; misses: number } = { hits: 0, misses: 0 };

  constructor(config: CacheConfig = {}) {
    this.config = {
      enabled: config.enabled ?? true,
      defaultTTL: config.defaultTTL ?? 300, // 5 minutes
      maxEntries: config.maxEntries ?? 1000,
      similarityMatching: config.similarityMatching ?? true,
      similarityThreshold: config.similarityThreshold ?? 0.85,
    };
  }

  /**
   * Generate cache key from query and context
   */
  generateKey(query: string, context?: Record<string, unknown>): string {
    const normalized = this.normalizeQuery(query);
    const contextStr = context ? JSON.stringify(this.sortObject(context)) : '';
    const combined = `${normalized}|${contextStr}`;
    return crypto.createHash('md5').update(combined).digest('hex');
  }

  /**
   * Get entry from cache
   */
  get(key: string): T | undefined {
    if (!this.config.enabled) return undefined;

    const entry = this.cache.get(key);
    if (!entry) {
      this.stats.misses++;
      return undefined;
    }

    // Check expiration
    if (new Date() > entry.expiresAt) {
      this.cache.delete(key);
      this.stats.misses++;
      return undefined;
    }

    entry.hits++;
    this.stats.hits++;
    return entry.value;
  }

  /**
   * Get with similar query matching
   */
  getWithSimilarity(query: string, language: 'de' | 'en'): T | undefined {
    if (!this.config.enabled || !this.config.similarityMatching) {
      return undefined;
    }

    const normalized = this.normalizeQuery(query);

    for (const [key, entry] of this.cache) {
      // Skip expired entries
      if (new Date() > entry.expiresAt) {
        this.cache.delete(key);
        continue;
      }

      // Must match language
      if (entry.language !== language) continue;

      // Calculate similarity
      const similarity = this.calculateSimilarity(normalized, entry.queryHash);
      if (similarity >= this.config.similarityThreshold) {
        entry.hits++;
        this.stats.hits++;
        return entry.value;
      }
    }

    this.stats.misses++;
    return undefined;
  }

  /**
   * Set entry in cache
   */
  set(
    key: string,
    value: T,
    options?: { ttl?: number; language?: 'de' | 'en'; query?: string }
  ): void {
    if (!this.config.enabled) return;

    // Evict if at capacity
    if (this.cache.size >= this.config.maxEntries) {
      this.evict();
    }

    const now = new Date();
    const ttl = options?.ttl ?? this.config.defaultTTL;

    this.cache.set(key, {
      key,
      value,
      language: options?.language ?? 'en',
      createdAt: now,
      expiresAt: new Date(now.getTime() + ttl * 1000),
      hits: 0,
      queryHash: this.normalizeQuery(options?.query ?? ''),
    });
  }

  /**
   * Delete entry
   */
  delete(key: string): boolean {
    return this.cache.delete(key);
  }

  /**
   * Clear all entries
   */
  clear(): void {
    this.cache.clear();
    this.stats = { hits: 0, misses: 0 };
  }

  /**
   * Clear expired entries
   */
  clearExpired(): number {
    const now = new Date();
    let cleared = 0;

    for (const [key, entry] of this.cache) {
      if (now > entry.expiresAt) {
        this.cache.delete(key);
        cleared++;
      }
    }

    return cleared;
  }

  /**
   * Get cache statistics
   */
  getStats(): CacheStats {
    const entries = Array.from(this.cache.values());
    const total = this.stats.hits + this.stats.misses;

    let oldest: Date | null = null;
    let newest: Date | null = null;

    for (const entry of entries) {
      if (!oldest || entry.createdAt < oldest) {
        oldest = entry.createdAt;
      }
      if (!newest || entry.createdAt > newest) {
        newest = entry.createdAt;
      }
    }

    return {
      hits: this.stats.hits,
      misses: this.stats.misses,
      entries: this.cache.size,
      hitRate: total > 0 ? this.stats.hits / total : 0,
      oldestEntry: oldest,
      newestEntry: newest,
    };
  }

  /**
   * Evict least recently used/least hits entries
   */
  private evict(): void {
    // Clear expired first
    this.clearExpired();

    // If still at capacity, remove lowest hit entries
    if (this.cache.size >= this.config.maxEntries) {
      const entries = Array.from(this.cache.entries()).sort((a, b) => a[1].hits - b[1].hits);

      // Remove bottom 10%
      const toRemove = Math.max(1, Math.floor(entries.length * 0.1));
      for (let i = 0; i < toRemove; i++) {
        this.cache.delete(entries[i][0]);
      }
    }
  }

  /**
   * Normalize query for comparison
   */
  private normalizeQuery(query: string): string {
    return query
      .toLowerCase()
      .trim()
      .replace(/[^\w\säöüß]/g, ' ')
      .replace(/\s+/g, ' ');
  }

  /**
   * Calculate similarity between two normalized queries
   */
  private calculateSimilarity(query1: string, query2: string): number {
    if (query1 === query2) return 1;
    if (!query1 || !query2) return 0;

    const words1 = new Set(query1.split(' '));
    const words2 = new Set(query2.split(' '));

    // Jaccard similarity
    const intersection = new Set([...words1].filter((x) => words2.has(x)));
    const union = new Set([...words1, ...words2]);

    return intersection.size / union.size;
  }

  /**
   * Sort object keys for consistent hashing
   */
  private sortObject(obj: Record<string, unknown>): Record<string, unknown> {
    const sorted: Record<string, unknown> = {};
    for (const key of Object.keys(obj).sort()) {
      sorted[key] = obj[key];
    }
    return sorted;
  }

  /**
   * Check if caching is enabled
   */
  isEnabled(): boolean {
    return this.config.enabled;
  }

  /**
   * Enable/disable caching
   */
  setEnabled(enabled: boolean): void {
    this.config.enabled = enabled;
  }

  /**
   * Get configuration
   */
  getConfig(): Required<CacheConfig> {
    return { ...this.config };
  }

  /**
   * Update configuration
   */
  updateConfig(config: Partial<CacheConfig>): void {
    this.config = {
      ...this.config,
      ...config,
    };
  }

  /**
   * Get all entries for a language
   */
  getEntriesByLanguage(language: 'de' | 'en'): CacheEntry<T>[] {
    const entries: CacheEntry<T>[] = [];
    const now = new Date();

    for (const entry of this.cache.values()) {
      if (entry.language === language && now <= entry.expiresAt) {
        entries.push(entry);
      }
    }

    return entries;
  }

  /**
   * Get entry count
   */
  get size(): number {
    return this.cache.size;
  }

  /**
   * Check if cache has entry
   */
  has(key: string): boolean {
    const entry = this.cache.get(key);
    if (!entry) return false;

    // Check expiration
    if (new Date() > entry.expiresAt) {
      this.cache.delete(key);
      return false;
    }

    return true;
  }
}

export default CacheManager;
