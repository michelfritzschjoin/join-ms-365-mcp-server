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
 * Query Store class - manages persistent query storage
 */
export class QueryStore {
  private dataDir: string;
  private queriesFile: string;
  private queries: StoredQuery[] = [];
  private maxQueries: number;
  private retentionDays: number;
  private saveDebounceTimer: ReturnType<typeof setTimeout> | null = null;

  constructor() {
    this.dataDir = process.env.QUERY_STORE_DIR || path.join(__dirname, '..', 'data');
    this.queriesFile = path.join(this.dataDir, 'queries.json');
    this.maxQueries = parseInt(process.env.QUERY_STORE_MAX_QUERIES || '100000', 10);
    this.retentionDays = parseInt(process.env.QUERY_STORE_RETENTION_DAYS || '90', 10);

    this.ensureDataDir();
    this.loadQueries();
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
