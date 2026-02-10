/**
 * Query Store - Persistent storage for all user queries
 *
 * Stores all MCP tool calls and queries from users for auditing,
 * analytics, and debugging purposes.
 *
 * GDPR/DSGVO Compliance:
 * - User IDs are hashed for pseudonymization
 * - Sensitive data (tokens, passwords) are never stored
 * - Data retention period is configurable
 * - Right to erasure is supported via deleteUserQueries()
 *
 * ISO 27001 Compliance:
 * - All queries are logged with timestamps
 * - Access to query data requires authentication
 * - Audit trail is maintained
 */

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { createHash } from 'crypto';
import logger from './logger.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

/**
 * Stored query record
 */
export interface StoredQuery {
  /** Unique query ID */
  id: string;
  /** Timestamp of the query (ISO 8601) */
  timestamp: string;
  /** Hashed user ID for pseudonymization */
  userIdHash: string;
  /** Chat/Session ID */
  chatId?: string;
  /** Tool name that was called */
  toolName: string;
  /** Tool parameters (sanitized) */
  parameters: Record<string, unknown>;
  /** Response summary (truncated for storage) */
  responseSummary?: string;
  /** Whether the call was successful */
  success: boolean;
  /** Error message if failed */
  errorMessage?: string;
  /** Duration in milliseconds */
  durationMs?: number;
  /** IP address (anonymized - only first 2 octets) */
  ipAnonymized?: string;
  /** User agent (browser/client info) */
  userAgent?: string;
}

/**
 * Query statistics
 */
export interface QueryStats {
  totalQueries: number;
  uniqueUsers: number;
  successRate: number;
  averageDuration: number;
  topTools: { tool: string; count: number }[];
  queriesPerHour: { hour: string; count: number }[];
  errorRate: number;
}

/**
 * Query filter options
 */
export interface QueryFilter {
  userIdHash?: string;
  toolName?: string;
  startDate?: Date;
  endDate?: Date;
  success?: boolean;
  limit?: number;
  offset?: number;
}

/**
 * Query pattern for learning from user history
 */
export interface QueryPattern {
  /** Normalized query pattern */
  pattern: string;
  /** Number of times this pattern was used */
  count: number;
  /** Success rate (0.0 - 1.0) */
  successRate: number;
  /** Optimal entity types learned from this pattern */
  optimalEntityTypes: string[];
  /** Average duration in milliseconds */
  avgDuration: number;
  /** Last used timestamp */
  lastUsed: string;
}

/**
 * Entity type recommendation based on user history
 */
export interface EntityTypeRecommendation {
  /** Recommended entity types */
  entityTypes: string[];
  /** Confidence score (0.0 - 1.0) */
  confidence: number;
  /** Reason for recommendation */
  reason: string;
}

/**
 * Stored query pattern for persistent learning
 */
interface StoredQueryPattern {
  /** Normalized query pattern */
  pattern: string;
  /** User ID hash for isolation */
  userIdHash: string;
  /** Entity types used */
  entityTypes: string[];
  /** Success count */
  successCount: number;
  /** Failure count */
  failureCount: number;
  /** Total duration in milliseconds */
  totalDuration: number;
  /** Query count */
  queryCount: number;
  /** Last used timestamp */
  lastUsed: string;
  /** First used timestamp */
  firstUsed: string;
}

/**
 * Stored query transformation for automatic optimization learning
 */
export interface StoredQueryTransformation {
  /** Original query pattern (normalized) */
  originalPattern: string;
  /** Optimized query pattern that worked */
  optimizedPattern: string;
  /** User ID hash for isolation */
  userIdHash: string;
  /** Tool context (e.g., 'email', 'files', 'search') */
  toolContext: string;
  /** Success count */
  successCount: number;
  /** Failure count */
  failureCount: number;
  /** Last used timestamp */
  lastUsed: string;
  /** First used timestamp */
  firstUsed: string;
}

/**
 * Query Store class - manages persistent query storage
 */
export class QueryStore {
  private dataDir: string;
  private queriesFile: string;
  private patternsFile: string;
  private transformationsFile: string;
  private queries: StoredQuery[] = [];
  private patterns: StoredQueryPattern[] = [];
  private transformations: StoredQueryTransformation[] = [];
  private maxQueries: number;
  private retentionDays: number;
  private saveDebounceTimer: ReturnType<typeof setTimeout> | null = null;
  private patternSaveDebounceTimer: ReturnType<typeof setTimeout> | null = null;
  private transformationSaveDebounceTimer: ReturnType<typeof setTimeout> | null = null;
  private readonly minPatternCount: number;

  constructor() {
    this.dataDir = process.env.QUERY_STORE_DIR || path.join(__dirname, '..', 'data');
    this.queriesFile = path.join(this.dataDir, 'queries.json');
    this.patternsFile = path.join(this.dataDir, 'query-patterns.json');
    this.transformationsFile = path.join(this.dataDir, 'query-transformations.json');
    this.maxQueries = parseInt(process.env.QUERY_STORE_MAX_QUERIES || '100000', 10);
    this.retentionDays = parseInt(process.env.QUERY_STORE_RETENTION_DAYS || '90', 10);
    this.minPatternCount = parseInt(process.env.MS365_MCP_PATTERN_MIN_COUNT || '3', 10);

    this.ensureDataDir();
    this.loadQueries();
    this.loadPatterns();
    this.loadTransformations();
    this.startRetentionCleanup();
  }

  /**
   * Ensure data directory exists
   */
  private ensureDataDir(): void {
    if (!fs.existsSync(this.dataDir)) {
      fs.mkdirSync(this.dataDir, { recursive: true });
      logger.info('Query store data directory created', { path: this.dataDir });
    }
  }

  /**
   * Load queries from disk
   */
  private loadQueries(): void {
    try {
      if (fs.existsSync(this.queriesFile)) {
        const data = fs.readFileSync(this.queriesFile, 'utf-8');
        const parsed = JSON.parse(data);
        if (Array.isArray(parsed)) {
          this.queries = parsed;
          logger.info('Query store loaded', { count: this.queries.length });
        }
      }
    } catch (error) {
      logger.error('Failed to load query store:', error);
      this.queries = [];
    }
  }

  /**
   * Load query patterns from disk
   */
  private loadPatterns(): void {
    try {
      if (fs.existsSync(this.patternsFile)) {
        const data = fs.readFileSync(this.patternsFile, 'utf-8');
        const parsed = JSON.parse(data);
        if (Array.isArray(parsed)) {
          this.patterns = parsed;
          logger.info('Query patterns loaded', { count: this.patterns.length });
        }
      }
    } catch (error) {
      logger.error('Failed to load query patterns:', error);
      this.patterns = [];
    }
  }

  /**
   * Save queries to disk (debounced)
   */
  private saveQueries(): void {
    if (this.saveDebounceTimer) {
      clearTimeout(this.saveDebounceTimer);
    }

    this.saveDebounceTimer = setTimeout(() => {
      try {
        fs.writeFileSync(this.queriesFile, JSON.stringify(this.queries, null, 2), 'utf-8');
        logger.debug('Query store saved', { count: this.queries.length });
      } catch (error) {
        logger.error('Failed to save query store:', error);
      }
    }, 1000); // Debounce for 1 second
  }

  /**
   * Save query patterns to disk (debounced)
   */
  private savePatterns(): void {
    if (this.patternSaveDebounceTimer) {
      clearTimeout(this.patternSaveDebounceTimer);
    }

    this.patternSaveDebounceTimer = setTimeout(() => {
      try {
        fs.writeFileSync(this.patternsFile, JSON.stringify(this.patterns, null, 2), 'utf-8');
        logger.debug('Query patterns saved', { count: this.patterns.length });
      } catch (error) {
        logger.error('Failed to save query patterns:', error);
      }
    }, 1000); // Debounce for 1 second
  }

  /**
   * Start periodic retention cleanup
   */
  private startRetentionCleanup(): void {
    // Run cleanup every hour
    setInterval(
      () => {
        this.cleanupOldQueries();
      },
      60 * 60 * 1000
    );

    // Also run on startup
    this.cleanupOldQueries();
  }

  /**
   * Remove queries older than retention period
   */
  private cleanupOldQueries(): void {
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - this.retentionDays);

    const beforeCount = this.queries.length;
    this.queries = this.queries.filter((q) => new Date(q.timestamp) > cutoffDate);

    if (this.queries.length < beforeCount) {
      logger.info('Query store cleanup completed', {
        removed: beforeCount - this.queries.length,
        remaining: this.queries.length,
        retentionDays: this.retentionDays,
      });
      this.saveQueries();
    }
  }

  /**
   * Hash user ID for pseudonymization (GDPR compliance)
   */
  public hashUserId(userId: string): string {
    if (!userId) return 'anonymous';
    return createHash('sha256').update(userId).digest('hex').substring(0, 16);
  }

  /**
   * Anonymize IP address (only keep first 2 octets)
   */
  public anonymizeIp(ip: string): string {
    if (!ip) return 'unknown';
    const parts = ip.split('.');
    if (parts.length >= 2) {
      return `${parts[0]}.${parts[1]}.x.x`;
    }
    // IPv6 or other format
    return ip.substring(0, 10) + '...';
  }

  /**
   * Sanitize parameters - remove sensitive data
   */
  private sanitizeParameters(params: Record<string, unknown>): Record<string, unknown> {
    const sensitiveKeys = [
      'password',
      'token',
      'secret',
      'key',
      'authorization',
      'bearer',
      'credential',
      'accessToken',
      'refreshToken',
    ];

    const sanitized: Record<string, unknown> = {};

    for (const [key, value] of Object.entries(params)) {
      const lowerKey = key.toLowerCase();
      if (sensitiveKeys.some((sk) => lowerKey.includes(sk))) {
        sanitized[key] = '[REDACTED]';
      } else if (typeof value === 'string' && value.length > 500) {
        sanitized[key] = value.substring(0, 500) + '... [truncated]';
      } else if (typeof value === 'object' && value !== null) {
        sanitized[key] = this.sanitizeParameters(value as Record<string, unknown>);
      } else {
        sanitized[key] = value;
      }
    }

    return sanitized;
  }

  /**
   * Generate unique query ID
   */
  private generateId(): string {
    return `q_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`;
  }

  /**
   * Store a new query
   */
  public storeQuery(query: Omit<StoredQuery, 'id' | 'timestamp'>): StoredQuery {
    const storedQuery: StoredQuery = {
      id: this.generateId(),
      timestamp: new Date().toISOString(),
      ...query,
      parameters: this.sanitizeParameters(query.parameters),
      responseSummary: query.responseSummary?.substring(0, 1000),
    };

    this.queries.push(storedQuery);

    // Trim if over limit
    if (this.queries.length > this.maxQueries) {
      this.queries = this.queries.slice(-this.maxQueries);
    }

    this.saveQueries();

    logger.debug('Query stored', {
      id: storedQuery.id,
      tool: storedQuery.toolName,
      userIdHash: storedQuery.userIdHash,
    });

    return storedQuery;
  }

  /**
   * Get queries with optional filtering
   */
  public getQueries(filter: QueryFilter = {}): StoredQuery[] {
    let result = [...this.queries];

    if (filter.userIdHash) {
      result = result.filter((q) => q.userIdHash === filter.userIdHash);
    }

    if (filter.toolName) {
      result = result.filter((q) => q.toolName === filter.toolName);
    }

    if (filter.startDate) {
      result = result.filter((q) => new Date(q.timestamp) >= filter.startDate!);
    }

    if (filter.endDate) {
      result = result.filter((q) => new Date(q.timestamp) <= filter.endDate!);
    }

    if (filter.success !== undefined) {
      result = result.filter((q) => q.success === filter.success);
    }

    // Sort by timestamp descending (newest first)
    result.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());

    // Apply pagination
    const offset = filter.offset || 0;
    const limit = filter.limit || 100;
    result = result.slice(offset, offset + limit);

    return result;
  }

  /**
   * Get total count of queries (for pagination)
   */
  public getQueryCount(filter: QueryFilter = {}): number {
    let result = [...this.queries];

    if (filter.userIdHash) {
      result = result.filter((q) => q.userIdHash === filter.userIdHash);
    }

    if (filter.toolName) {
      result = result.filter((q) => q.toolName === filter.toolName);
    }

    if (filter.startDate) {
      result = result.filter((q) => new Date(q.timestamp) >= filter.startDate!);
    }

    if (filter.endDate) {
      result = result.filter((q) => new Date(q.timestamp) <= filter.endDate!);
    }

    if (filter.success !== undefined) {
      result = result.filter((q) => q.success === filter.success);
    }

    return result.length;
  }

  /**
   * Get query statistics
   */
  public getStats(): QueryStats {
    const total = this.queries.length;
    const uniqueUsers = new Set(this.queries.map((q) => q.userIdHash)).size;
    const successful = this.queries.filter((q) => q.success).length;

    // Calculate average duration
    const durationsWithValues = this.queries
      .filter((q) => q.durationMs !== undefined)
      .map((q) => q.durationMs!);
    const avgDuration =
      durationsWithValues.length > 0
        ? durationsWithValues.reduce((a, b) => a + b, 0) / durationsWithValues.length
        : 0;

    // Top tools
    const toolCounts: Record<string, number> = {};
    for (const q of this.queries) {
      toolCounts[q.toolName] = (toolCounts[q.toolName] || 0) + 1;
    }
    const topTools = Object.entries(toolCounts)
      .map(([tool, count]) => ({ tool, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 10);

    // Queries per hour (last 24 hours)
    const now = new Date();
    const hourlyData: Record<string, number> = {};
    for (let i = 23; i >= 0; i--) {
      const hour = new Date(now.getTime() - i * 60 * 60 * 1000);
      const hourKey = hour.toISOString().substring(0, 13) + ':00';
      hourlyData[hourKey] = 0;
    }

    for (const q of this.queries) {
      const hourKey = q.timestamp.substring(0, 13) + ':00';
      if (hourKey in hourlyData) {
        hourlyData[hourKey]++;
      }
    }

    const queriesPerHour = Object.entries(hourlyData).map(([hour, count]) => ({ hour, count }));

    return {
      totalQueries: total,
      uniqueUsers,
      successRate: total > 0 ? (successful / total) * 100 : 0,
      averageDuration: Math.round(avgDuration),
      topTools,
      queriesPerHour,
      errorRate: total > 0 ? ((total - successful) / total) * 100 : 0,
    };
  }

  /**
   * Get unique tool names
   */
  public getToolNames(): string[] {
    return [...new Set(this.queries.map((q) => q.toolName))].sort();
  }

  /**
   * Delete all queries for a user (GDPR Right to Erasure)
   */
  public deleteUserQueries(userIdHash: string): number {
    const beforeCount = this.queries.length;
    this.queries = this.queries.filter((q) => q.userIdHash !== userIdHash);
    const deleted = beforeCount - this.queries.length;

    if (deleted > 0) {
      this.saveQueries();
      logger.info('User queries deleted (GDPR erasure)', {
        userIdHash,
        deleted,
      });
    }

    return deleted;
  }

  /**
   * Export queries for a user (GDPR Data Portability)
   */
  public exportUserQueries(userIdHash: string): StoredQuery[] {
    return this.queries.filter((q) => q.userIdHash === userIdHash);
  }

  /**
   * Clear all queries (admin function)
   */
  public clearAll(): void {
    this.queries = [];
    this.saveQueries();
    logger.warn('All queries cleared from store');
  }

  // =========================================================================
  // QUERY PATTERN LEARNING (USER-SPECIFIC)
  // =========================================================================

  /**
   * Normalize a query to a pattern for matching
   * Removes specific values but keeps structure
   */
  public normalizeQueryPattern(query: string): string {
    return (
      query
        .toLowerCase()
        .trim()
        // Normalize whitespace
        .replace(/\s+/g, ' ')
        // Remove specific dates/times but keep temporal markers
        .replace(/\d{1,2}[./-]\d{1,2}[./-]\d{2,4}/g, '<DATE>')
        .replace(/\d{1,2}:\d{2}(:\d{2})?/g, '<TIME>')
        // Remove specific numbers but keep numeric placeholders
        .replace(/\b\d{4,}\b/g, '<NUM>')
        // Remove email addresses
        .replace(/\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/gi, '<EMAIL>')
        // Normalize common stop words
        .replace(/\b(der|die|das|ein|eine|für|von|mit|zu|and|the|a|for|with|to)\b/gi, '')
        // Collapse multiple spaces
        .replace(/\s+/g, ' ')
        .trim()
    );
  }

  /**
   * Record a query pattern for learning (USER-SPECIFIC)
   * @param userIdHash - Hashed user ID for isolation
   * @param query - Original query string
   * @param entityTypes - Entity types used in the search
   * @param success - Whether the query was successful (found results)
   * @param duration - Query duration in milliseconds
   */
  public recordQueryPattern(
    userIdHash: string,
    query: string,
    entityTypes: string[],
    success: boolean,
    duration: number
  ): void {
    if (!userIdHash || !query || entityTypes.length === 0) {
      return;
    }

    const pattern = this.normalizeQueryPattern(query);
    if (!pattern) {
      return;
    }

    const now = new Date().toISOString();

    // Find existing pattern for this user
    const existingIndex = this.patterns.findIndex(
      (p) => p.userIdHash === userIdHash && p.pattern === pattern
    );

    if (existingIndex >= 0) {
      // Update existing pattern
      const existing = this.patterns[existingIndex];
      existing.queryCount++;
      if (success) {
        existing.successCount++;
      } else {
        existing.failureCount++;
      }
      existing.totalDuration += duration;
      existing.lastUsed = now;

      // Update entity types if this was a successful query
      if (success) {
        // Merge entity types, keeping most frequent ones
        const allTypes = [...existing.entityTypes, ...entityTypes];
        const typeCounts = new Map<string, number>();
        for (const t of allTypes) {
          typeCounts.set(t, (typeCounts.get(t) || 0) + 1);
        }
        existing.entityTypes = [...typeCounts.entries()]
          .sort((a, b) => b[1] - a[1])
          .slice(0, 5)
          .map(([type]) => type);
      }
    } else {
      // Create new pattern
      const newPattern: StoredQueryPattern = {
        pattern,
        userIdHash,
        entityTypes: [...new Set(entityTypes)].slice(0, 5),
        successCount: success ? 1 : 0,
        failureCount: success ? 0 : 1,
        totalDuration: duration,
        queryCount: 1,
        lastUsed: now,
        firstUsed: now,
      };
      this.patterns.push(newPattern);
    }

    this.savePatterns();

    logger.debug('Query pattern recorded', {
      pattern,
      userIdHash: userIdHash.substring(0, 8) + '...',
      success,
      entityTypes,
    });
  }

  /**
   * Get query patterns for a user (USER-SPECIFIC)
   * @param userIdHash - Hashed user ID
   * @param limit - Maximum number of patterns to return
   * @returns Array of query patterns sorted by frequency
   */
  public getQueryPatterns(userIdHash: string, limit = 20): QueryPattern[] {
    if (!userIdHash) {
      return [];
    }

    // Filter patterns for this user with minimum count
    const userPatterns = this.patterns.filter(
      (p) => p.userIdHash === userIdHash && p.queryCount >= this.minPatternCount
    );

    // Convert to QueryPattern format and sort by count
    return userPatterns
      .map((p) => ({
        pattern: p.pattern,
        count: p.queryCount,
        successRate: p.queryCount > 0 ? p.successCount / p.queryCount : 0,
        optimalEntityTypes: p.entityTypes,
        avgDuration: p.queryCount > 0 ? Math.round(p.totalDuration / p.queryCount) : 0,
        lastUsed: p.lastUsed,
      }))
      .sort((a, b) => b.count - a.count)
      .slice(0, limit);
  }

  /**
   * Get optimal entity types for a query based on user history (USER-SPECIFIC)
   * @param query - Query string to find recommendations for
   * @param userIdHash - Hashed user ID for isolation
   * @returns Entity type recommendation or null if no match found
   */
  public getOptimalEntityTypes(query: string, userIdHash: string): EntityTypeRecommendation | null {
    if (!query || !userIdHash) {
      return null;
    }

    const pattern = this.normalizeQueryPattern(query);
    if (!pattern) {
      return null;
    }

    // Find exact pattern match for this user
    const exactMatch = this.patterns.find(
      (p) =>
        p.userIdHash === userIdHash && p.pattern === pattern && p.queryCount >= this.minPatternCount
    );

    if (exactMatch && exactMatch.entityTypes.length > 0) {
      const successRate =
        exactMatch.queryCount > 0 ? exactMatch.successCount / exactMatch.queryCount : 0;

      // Only recommend if success rate is above 50%
      if (successRate >= 0.5) {
        return {
          entityTypes: exactMatch.entityTypes,
          confidence: Math.min(0.95, successRate * (exactMatch.queryCount / 10)),
          reason: `Based on ${exactMatch.queryCount} similar queries (${Math.round(successRate * 100)}% success rate)`,
        };
      }
    }

    // Try partial matching - find patterns that share significant words
    const patternWords = pattern.split(' ').filter((w) => w.length > 3);
    if (patternWords.length === 0) {
      return null;
    }

    // Find patterns with overlapping words
    const matchingPatterns = this.patterns.filter((p) => {
      if (p.userIdHash !== userIdHash || p.queryCount < this.minPatternCount) {
        return false;
      }
      const pWords = p.pattern.split(' ').filter((w) => w.length > 3);
      const overlap = patternWords.filter((w) => pWords.includes(w)).length;
      return overlap >= Math.min(2, Math.ceil(patternWords.length * 0.5));
    });

    if (matchingPatterns.length === 0) {
      return null;
    }

    // Calculate weighted entity types from matching patterns
    const entityTypeScores = new Map<string, number>();
    let totalWeight = 0;

    for (const match of matchingPatterns) {
      const successRate = match.queryCount > 0 ? match.successCount / match.queryCount : 0;
      const weight = match.queryCount * successRate;
      totalWeight += weight;

      for (const et of match.entityTypes) {
        entityTypeScores.set(et, (entityTypeScores.get(et) || 0) + weight);
      }
    }

    if (entityTypeScores.size === 0 || totalWeight === 0) {
      return null;
    }

    // Sort by score and take top entity types
    const sortedTypes = [...entityTypeScores.entries()]
      .sort((a, b) => b[1] - a[1])
      .slice(0, 4)
      .map(([type]) => type);

    // Calculate confidence based on match quality
    const avgScore = totalWeight / matchingPatterns.length;
    const confidence = Math.min(0.8, avgScore / 5); // Cap at 0.8 for partial matches

    if (confidence < 0.3) {
      return null; // Too low confidence
    }

    return {
      entityTypes: sortedTypes,
      confidence,
      reason: `Based on ${matchingPatterns.length} similar patterns with ${Math.round(confidence * 100)}% confidence`,
    };
  }

  // =========================================================================
  // QUERY TRANSFORMATION LEARNING (FOR AUTOMATIC QUERY OPTIMIZATION)
  // =========================================================================

  /**
   * Load query transformations from disk
   */
  private loadTransformations(): void {
    try {
      if (fs.existsSync(this.transformationsFile)) {
        const data = fs.readFileSync(this.transformationsFile, 'utf-8');
        const parsed = JSON.parse(data);
        if (Array.isArray(parsed)) {
          this.transformations = parsed;
          logger.info('Query transformations loaded', { count: this.transformations.length });
        }
      }
    } catch (error) {
      logger.error('Failed to load query transformations:', error);
      this.transformations = [];
    }
  }

  /**
   * Save query transformations to disk (debounced)
   */
  private saveTransformations(): void {
    if (this.transformationSaveDebounceTimer) {
      clearTimeout(this.transformationSaveDebounceTimer);
    }

    this.transformationSaveDebounceTimer = setTimeout(() => {
      try {
        fs.writeFileSync(
          this.transformationsFile,
          JSON.stringify(this.transformations, null, 2),
          'utf-8'
        );
        logger.debug('Query transformations saved', { count: this.transformations.length });
      } catch (error) {
        logger.error('Failed to save query transformations:', error);
      }
    }, 1000);
  }

  /**
   * Record a query variant result for learning which optimizations work
   * @param originalQuery - The original unmodified query
   * @param optimizedQuery - The optimized query that was executed
   * @param success - Whether the optimized query returned results
   * @param userIdHash - Hashed user ID for isolation
   * @param toolContext - Which tool executed the query (e.g., 'email', 'search', 'files')
   */
  public recordQueryVariant(
    originalQuery: string,
    optimizedQuery: string,
    success: boolean,
    userIdHash: string,
    toolContext: string = 'search'
  ): void {
    if (!originalQuery || !optimizedQuery || !userIdHash) {
      return;
    }

    // Don't record if they're the same
    if (originalQuery.toLowerCase().trim() === optimizedQuery.toLowerCase().trim()) {
      return;
    }

    const originalNormalized = originalQuery.toLowerCase().trim();
    const optimizedNormalized = optimizedQuery.toLowerCase().trim();
    const now = new Date().toISOString();

    // Find existing transformation
    const existingIndex = this.transformations.findIndex(
      (t) =>
        t.userIdHash === userIdHash &&
        t.originalPattern === originalNormalized &&
        t.optimizedPattern === optimizedNormalized &&
        t.toolContext === toolContext
    );

    if (existingIndex >= 0) {
      // Update existing transformation
      const existing = this.transformations[existingIndex];
      if (success) {
        existing.successCount++;
      } else {
        existing.failureCount++;
      }
      existing.lastUsed = now;
    } else {
      // Create new transformation record
      const newTransformation: StoredQueryTransformation = {
        originalPattern: originalNormalized,
        optimizedPattern: optimizedNormalized,
        userIdHash,
        toolContext,
        successCount: success ? 1 : 0,
        failureCount: success ? 0 : 1,
        lastUsed: now,
        firstUsed: now,
      };
      this.transformations.push(newTransformation);
    }

    this.saveTransformations();

    logger.debug('Query variant recorded', {
      original: originalNormalized.substring(0, 50),
      optimized: optimizedNormalized.substring(0, 50),
      success,
      toolContext,
    });
  }

  /**
   * Get successful query variants for a pattern
   * @param query - Query to find successful variants for
   * @param userIdHash - Hashed user ID
   * @param toolContext - Optional tool context filter
   * @returns Array of successful query variants sorted by success rate
   */
  public getSuccessfulQueryVariants(
    query: string,
    userIdHash: string,
    toolContext?: string
  ): string[] {
    if (!query || !userIdHash) {
      return [];
    }

    const normalized = query.toLowerCase().trim();

    return this.transformations
      .filter((t) => {
        const isMatch =
          t.userIdHash === userIdHash &&
          (t.originalPattern === normalized || t.optimizedPattern === normalized);
        const isToolMatch = !toolContext || t.toolContext === toolContext;
        const total = t.successCount + t.failureCount;
        const hasMinCount = total >= this.minPatternCount;
        const isSuccessful = total > 0 && t.successCount / total >= 0.5;
        return isMatch && isToolMatch && hasMinCount && isSuccessful;
      })
      .sort((a, b) => {
        const rateA = a.successCount / (a.successCount + a.failureCount);
        const rateB = b.successCount / (b.successCount + b.failureCount);
        return rateB - rateA;
      })
      .map((t) => (t.originalPattern === normalized ? t.optimizedPattern : t.originalPattern));
  }

  /**
   * Get learned query transformation patterns for automatic optimization
   * @param userIdHash - Hashed user ID
   * @param toolContext - Optional tool context filter
   * @returns Array of transformations sorted by success rate
   */
  public getQueryTransformationPatterns(
    userIdHash: string,
    toolContext?: string
  ): StoredQueryTransformation[] {
    if (!userIdHash) {
      return [];
    }

    return this.transformations
      .filter((t) => {
        const isUser = t.userIdHash === userIdHash;
        const isToolMatch = !toolContext || t.toolContext === toolContext;
        const total = t.successCount + t.failureCount;
        const hasMinCount = total >= this.minPatternCount;
        const isSuccessful = total > 0 && t.successCount / total >= 0.5;
        return isUser && isToolMatch && hasMinCount && isSuccessful;
      })
      .sort((a, b) => {
        const rateA = a.successCount / (a.successCount + a.failureCount);
        const rateB = b.successCount / (b.successCount + b.failureCount);
        return rateB - rateA;
      });
  }

  /**
   * Delete transformations for a user (GDPR Right to Erasure)
   */
  public deleteUserTransformations(userIdHash: string): number {
    const beforeCount = this.transformations.length;
    this.transformations = this.transformations.filter((t) => t.userIdHash !== userIdHash);
    const deleted = beforeCount - this.transformations.length;

    if (deleted > 0) {
      this.saveTransformations();
      logger.info('User query transformations deleted (GDPR erasure)', {
        userIdHash,
        deleted,
      });
    }

    return deleted;
  }

  /**
   * Clear all transformations (admin function)
   */
  public clearAllTransformations(): void {
    this.transformations = [];
    this.saveTransformations();
    logger.warn('All query transformations cleared from store');
  }

  /**
   * Clear patterns for a user (GDPR Right to Erasure)
   */
  public deleteUserPatterns(userIdHash: string): number {
    const beforeCount = this.patterns.length;
    this.patterns = this.patterns.filter((p) => p.userIdHash !== userIdHash);
    const deleted = beforeCount - this.patterns.length;

    if (deleted > 0) {
      this.savePatterns();
      logger.info('User patterns deleted (GDPR erasure)', {
        userIdHash,
        deleted,
      });
    }

    return deleted;
  }

  /**
   * Clear all patterns (admin function)
   */
  public clearAllPatterns(): void {
    this.patterns = [];
    this.savePatterns();
    logger.warn('All query patterns cleared from store');
  }
}

// Singleton instance
let queryStoreInstance: QueryStore | null = null;

/**
 * Get the query store singleton
 */
export function getQueryStore(): QueryStore {
  if (!queryStoreInstance) {
    queryStoreInstance = new QueryStore();
  }
  return queryStoreInstance;
}

export default getQueryStore;
