/**
 * Super-Tools: Consolidated tool interface for Microsoft 365 MCP Server
 *
 * Instead of 126+ individual tools, we provide 10 "Super-Tools" that group
 * related functionality together. Each tool accepts an `action` parameter
 * to specify the operation.
 *
 * Benefits:
 * - Easier for LLMs to choose the right tool
 * - Cleaner UI in MCP clients
 * - Reduced cognitive load
 * - Same underlying functionality
 */

import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import GraphClient, { type GraphBatchRequest } from './graph-client.js';
import logger from './logger.js';
import { addThinkingToResponse, isThinkingEnabled } from './thinking-process.js';
import {
  formatCalendarResponse,
  formatCalendarEvent,
  calendarResponseToText,
  formatMailResponse,
  mailResponseToText,
  isCalendarResponse,
  isMailResponse,
} from './response-formatter.js';
import NLPEnhancer, { type DecomposedQuery, type ExtractedEntity } from './nlp-enhancer.js';
import DataAggregator, { type AggregationResult, type AggregatedItem } from './data-aggregator.js';
import {
  GraphApiError,
  RateLimitError,
  ServiceUnavailableError,
  isRetryableError,
  getRetryAfter,
} from './errors.js';
// UQAS Pro - Bilingual Support (DE/EN)
import { getUQAS } from './uqas/integration/index.js';
import DownloadLinkGenerator from './download-link-generator.js';
import type { AppSecrets } from './secrets.js';
import {
  getRequestTokens,
  getProfessionProfile,
  getUserProfile,
  getUserId,
} from './request-context.js';
import { validateEntityTypeCombinations } from './utils/entity-type-validator.js';
import {
  isLoopFile,
  detectLoopFile,
  parseLoopContent,
  formatLoopFileInfo,
} from './utils/loop-detector.js';
// Query Pattern Learning
import { getQueryStore } from './query-store.js';
// Automatic Query Optimization
import {
  getQueryOptimizer,
  type OptimizedQuery,
  type OptimizationContext,
} from './query-optimizer.js';
import type { ProfessionProfile } from './user-profile.js';
import {
  calendarResponseToTextByProfession,
  mailResponseToTextByProfession,
  formatDataByProfession,
  getProfessionGreeting,
} from './response-formatter.js';
import { findUser, findChatsWithUser, type GraphUser, type GraphChat } from './compound-tools.js';
import { getLearningSystem } from './discovery-tools.js';
import { encode as toonEncode } from '@toon-format/toon';
import {
  getPatternBasedExtractor,
  getExtractorRegistry,
  EntityExtractorRegistry,
  MetadataExtractor,
  SummaryGenerator,
  getExtractionCache,
  generateCacheKey as generateExtractionCacheKey,
  type BusinessContentExtraction,
  type DocumentType,
  type ProjectContent,
  type CustomerContent,
  type MeetingContent,
  type DocumentContent,
  type SalesContent,
  type HRContent,
  type ExtractorOptions,
} from './utils/content-extractor.js';

// Initialize NLP Enhancer for intelligent query processing
const nlpEnhancer = new NLPEnhancer();

// Initialize Data Aggregator for consistent data processing
const dataAggregator = new DataAggregator();

// Initialize UQAS for bilingual support
const uqas = getUQAS();

/**
 * Format search query for Microsoft Graph API endpoints that require property:value format
 * @param searchValue - The search query string
 * @param defaultProperty - The default property to use if not already specified (e.g., 'displayName')
 * @param searchType - Type of search: 'email', 'event', 'contact', or 'general'
 * @returns Formatted search query in property:value format, wrapped in double quotes
 */
function formatSearchQuery(
  searchValue: string | undefined | null,
  defaultProperty = 'displayName',
  searchType: 'email' | 'event' | 'contact' | 'general' = 'general'
): string {
  // Handle null, undefined, or non-string values
  if (!searchValue) return '';
  if (typeof searchValue !== 'string') {
    // Convert to string if possible, otherwise return empty
    const stringValue = String(searchValue);
    if (stringValue === 'null' || stringValue === 'undefined') return '';
    searchValue = stringValue;
  }

  // Check if search already contains a property prefix (e.g., "displayName:John")
  const propertyValuePattern = /^([a-zA-Z]+):(.+)$/i;
  const trimmedValue = searchValue.trim();

  // If already has property prefix, format it properly
  const propertyMatch = trimmedValue.match(propertyValuePattern);
  if (propertyMatch) {
    const property = propertyMatch[1];
    let value = propertyMatch[2];

    // Remove any existing escaped quotes or regular quotes from the value
    // Handle both \" and " formats - strip quotes from start and end
    value = value.trim();
    if (value.startsWith('\\"') && value.endsWith('\\"')) {
      value = value.slice(2, -2);
    } else if (value.startsWith('"') && value.endsWith('"')) {
      value = value.slice(1, -1);
    }

    // If value contains spaces, quote and escape it for KQL syntax
    // Format: property:"value with spaces" -> "property:\"value with spaces\""
    if (value.includes(' ')) {
      // Format as property:"value" then escape inner quotes when wrapping
      const innerQuery = `${property}:"${value}"`;
      // Escape the inner quotes for the outer wrapper
      return `"${innerQuery.replace(/"/g, '\\"')}"`;
    } else {
      // Simple value without spaces: property:value -> "property:value"
      return `"${property}:${value}"`;
    }
  }

  // Format based on search type
  // Note: Microsoft Graph $search searches across all fields by default when no property is specified
  // For better results, we can use simple text search or specific properties
  let formattedSearch: string;
  if (searchType === 'email') {
    // For emails, use simple text search (searches in subject, body, from, to, etc.)
    // Or use specific property if needed: subject:, body:, from:, to:
    formattedSearch = trimmedValue; // Simple text search searches all email fields
  } else if (searchType === 'event') {
    // For events, use simple text search (searches in subject, body, attendees, etc.)
    formattedSearch = trimmedValue; // Simple text search searches all event fields
  } else if (searchType === 'contact') {
    // For contacts, search in companyName specifically
    formattedSearch = `companyName:${trimmedValue}`;
  } else {
    // General search with default property
    formattedSearch = `${defaultProperty}:${trimmedValue}`;
  }

  return `"${formattedSearch}"`;
}

/**
 * Format KQL query for Microsoft Graph Search API
 * Handles property filters with spaces, OR/AND operators, and ensures proper KQL syntax
 * @param query - The KQL query string
 * @returns Formatted KQL query string
 */
function formatKQLQuery(query: string): string {
  if (!query) return '';

  const trimmedQuery = query.trim();

  // Normalize whitespace around operators
  let formattedQuery = trimmedQuery
    .replace(/\s+\b(OR|AND|NOT)\b\s+/gi, (match) => ` ${match.trim().toUpperCase()} `)
    .replace(/\s+/g, ' ')
    .trim();

  // Check if query contains property filters (e.g., from:, subject:, body:)
  const propertyFilterPattern = /\b(\w+):/;
  const hasPropertyFilters = propertyFilterPattern.test(formattedQuery);

  if (!hasPropertyFilters) {
    // No property filters, return as-is (but with normalized operators)
    return formattedQuery;
  }

  // Step 1: Quote property values that contain spaces and aren't already quoted
  // Pattern: property:value where value contains spaces
  formattedQuery = formattedQuery.replace(
    /(\w+):([^\s"()]+(?:\s+[^\s"()]+)+)(?=\s|$|AND|OR|NOT|\))/g,
    (match, property, value) => {
      // If value contains spaces and isn't already quoted, quote it
      if (value.includes(' ') && !value.startsWith('"') && !value.endsWith('"')) {
        return `${property}:"${value}"`;
      }
      return match;
    }
  );

  // Step 2: Handle property filters followed by free text with OR/AND operators
  // Pattern: property:"value" text1 OR text2 -> property:"value" AND (text1 OR text2)
  // Pattern: property:value text1 OR text2 -> property:value AND (text1 OR text2)
  const propertyAndFreeTextPattern = /^(\w+:"[^"]+"|\w+:[^\s"()]+)\s+([^()]+(?:OR|AND)[^()]+)$/i;
  const propertyAndFreeTextMatch = formattedQuery.match(propertyAndFreeTextPattern);

  if (propertyAndFreeTextMatch) {
    const propertyPart = propertyAndFreeTextMatch[1];
    let textPart = propertyAndFreeTextMatch[2].trim();

    // Normalize operators in text part
    textPart = textPart.replace(/\b(OR|AND|NOT)\b/gi, (match) => ` ${match.toUpperCase()} `).trim();

    // If text part contains OR/AND but isn't already grouped, wrap it in parentheses
    if (
      (textPart.includes(' OR ') || textPart.includes(' AND ')) &&
      !textPart.startsWith('(') &&
      !textPart.endsWith(')')
    ) {
      formattedQuery = `${propertyPart} AND (${textPart})`;
    } else if (!textPart.match(/^\w+:\S+$/)) {
      // If text part is not a property filter, add AND between property and text
      formattedQuery = `${propertyPart} AND ${textPart}`;
    } else {
      formattedQuery = `${propertyPart} ${textPart}`;
    }
  }

  return formattedQuery.trim();
}

// ============================================================================
// API CACHE - LRU Cache for API Responses
// ============================================================================

interface CacheEntry {
  data: unknown;
  timestamp: number;
  ttl: number;
}

interface CacheNode {
  key: string;
  value: CacheEntry;
  prev: CacheNode | null;
  next: CacheNode | null;
}

/**
 * Simple LRU Cache implementation for API responses
 */
class APICache {
  private cache = new Map<string, CacheNode>();
  private head: CacheNode | null = null;
  private tail: CacheNode | null = null;
  private maxSize: number;

  constructor(maxSize = 1000) {
    this.maxSize = maxSize;
  }

  /**
   * Get cached value if not expired
   */
  get(key: string): unknown | null {
    const node = this.cache.get(key);
    if (!node) {
      return null;
    }

    const entry = node.value;
    const now = Date.now();

    // Check if expired
    if (now - entry.timestamp > entry.ttl) {
      this.delete(key);
      return null;
    }

    // Move to front (most recently used)
    this.moveToFront(node);
    return entry.data;
  }

  /**
   * Set cache value with TTL
   */
  set(key: string, data: unknown, ttl: number): void {
    const now = Date.now();
    const entry: CacheEntry = { data, timestamp: now, ttl };

    let node = this.cache.get(key);

    if (node) {
      // Update existing node
      node.value = entry;
      this.moveToFront(node);
    } else {
      // Create new node
      node = { key, value: entry, prev: null, next: null };

      if (this.cache.size >= this.maxSize) {
        // Remove least recently used
        if (this.tail) {
          this.delete(this.tail.key);
        }
      }

      this.cache.set(key, node);
      this.moveToFront(node);
    }
  }

  /**
   * Delete cache entry
   */
  delete(key: string): void {
    const node = this.cache.get(key);
    if (!node) {
      return;
    }

    // Remove from linked list
    if (node.prev) {
      node.prev.next = node.next;
    } else {
      this.head = node.next;
    }

    if (node.next) {
      node.next.prev = node.prev;
    } else {
      this.tail = node.prev;
    }

    this.cache.delete(key);
  }

  /**
   * Invalidate cache entries matching pattern
   */
  invalidate(pattern: string): void {
    const regex = new RegExp(pattern);
    const keysToDelete: string[] = [];

    for (const key of this.cache.keys()) {
      if (regex.test(key)) {
        keysToDelete.push(key);
      }
    }

    for (const key of keysToDelete) {
      this.delete(key);
    }
  }

  /**
   * Clear all cache
   */
  clear(): void {
    this.cache.clear();
    this.head = null;
    this.tail = null;
  }

  /**
   * Get cache statistics
   */
  getStats(): { size: number; maxSize: number; hitRate: number } {
    return {
      size: this.cache.size,
      maxSize: this.maxSize,
      hitRate: 0, // Would need to track hits/misses for accurate rate
    };
  }

  private moveToFront(node: CacheNode): void {
    // Remove from current position
    if (node.prev) {
      node.prev.next = node.next;
    } else if (this.head === node) {
      return; // Already at front
    }

    if (node.next) {
      node.next.prev = node.prev;
    } else if (this.tail === node) {
      this.tail = node.prev;
    }

    // Add to front
    node.prev = null;
    node.next = this.head;

    if (this.head) {
      this.head.prev = node;
    } else {
      this.tail = node;
    }

    this.head = node;
  }
}

// Global API cache instance
const apiCache = new APICache(1000);

/**
 * Generate cache key from request parameters
 */
function generateCacheKey(
  method: string,
  endpoint: string,
  queryParams?: Record<string, string>,
  body?: unknown
): string {
  const parts = [method, endpoint];

  if (queryParams && Object.keys(queryParams).length > 0) {
    const sortedParams = Object.keys(queryParams)
      .sort()
      .map((key) => `${key}=${queryParams[key]}`)
      .join('&');
    parts.push(sortedParams);
  }

  if (body) {
    parts.push(JSON.stringify(body));
  }

  return parts.join('|');
}

/**
 * Determine if request should be cached
 */
function shouldCache(method: string, endpoint: string, body?: unknown): boolean {
  // Only cache GET requests (read-only)
  if (method !== 'GET') {
    return false;
  }

  // Don't cache search queries (too dynamic)
  if (endpoint.includes('/search/query')) {
    return false;
  }

  // Don't cache if body is present (POST with body)
  if (body) {
    return false;
  }

  return true;
}

/**
 * Determine TTL for cached response
 */
function getCacheTTL(endpoint: string): number {
  // Search queries: 30 seconds
  if (endpoint.includes('/search')) {
    return 30 * 1000;
  }

  // Calendar events: 2 minutes (can change frequently)
  if (endpoint.includes('/calendar') || endpoint.includes('/events')) {
    return 2 * 60 * 1000;
  }

  // Messages: 1 minute
  if (endpoint.includes('/messages')) {
    return 60 * 1000;
  }

  // Default: 5 minutes for read-only queries
  return 5 * 60 * 1000;
}

// ============================================================================
// RETRY LOGIC
// ============================================================================

interface RetryConfig {
  maxRetries: number;
  baseDelay: number;
  maxDelay: number;
  retryableStatusCodes: number[];
}

const defaultRetryConfig: RetryConfig = {
  maxRetries: 3,
  baseDelay: 1000, // 1 second
  maxDelay: 30000, // 30 seconds
  retryableStatusCodes: [429, 500, 502, 503, 504],
};

/**
 * Calculate delay with exponential backoff and jitter
 */
function calculateRetryDelay(attempt: number, baseDelay: number, maxDelay: number): number {
  const exponentialDelay = Math.min(baseDelay * Math.pow(2, attempt), maxDelay);
  const jitter = Math.random() * 0.3 * exponentialDelay; // 0-30% jitter
  return Math.floor(exponentialDelay + jitter);
}

/**
 * Call Graph API with retry logic and caching
 */
async function callGraphWithRetry(
  graphClient: GraphClient,
  method: 'GET' | 'POST' | 'PATCH' | 'DELETE' | '',
  endpoint: string,
  queryParams?: Record<string, string>,
  body?: unknown,
  headers?: Record<string, string>,
  retryConfig: RetryConfig = defaultRetryConfig,
  useCache = true
): Promise<string> {
  // Handle empty method (shouldn't happen but fail gracefully)
  if (!method || !endpoint) {
    throw new Error('Invalid callGraph: method and endpoint are required');
  }

  // Check cache for GET requests
  if (useCache && shouldCache(method, endpoint, body)) {
    const cacheKey = generateCacheKey(method, endpoint, queryParams, body);
    const cached = apiCache.get(cacheKey);
    if (cached !== null) {
      logger.debug(`Cache hit for ${method} ${endpoint}`);
      return typeof cached === 'string' ? cached : JSON.stringify(cached, null, 2);
    }
  }

  const options: {
    method: string;
    queryParams?: Record<string, string>;
    body?: string;
    headers?: Record<string, string>;
  } = { method };

  if (queryParams && Object.keys(queryParams).length > 0) {
    options.queryParams = queryParams;
  }

  if (body) {
    options.body = JSON.stringify(body);
  }

  if (headers && Object.keys(headers).length > 0) {
    options.headers = headers;
  }

  let lastError: Error | null = null;

  // Retry loop
  for (let attempt = 0; attempt <= retryConfig.maxRetries; attempt++) {
    try {
      const result = await graphClient.makeRequest(endpoint, options);
      const resultString = typeof result === 'string' ? result : JSON.stringify(result, null, 2);

      // Cache successful GET responses
      if (useCache && shouldCache(method, endpoint, body)) {
        const cacheKey = generateCacheKey(method, endpoint, queryParams, body);
        const ttl = getCacheTTL(endpoint);
        apiCache.set(cacheKey, result, ttl);
      }

      return resultString;
    } catch (error) {
      lastError = error instanceof Error ? error : new Error(String(error));

      // Check if error is retryable
      const statusCode = (error as GraphApiError).statusCode;
      const isRetryable = statusCode && retryConfig.retryableStatusCodes.includes(statusCode);

      // Don't retry on last attempt or if error is not retryable
      if (attempt >= retryConfig.maxRetries || !isRetryable) {
        break;
      }

      // Get retry-after header if available (for 429 errors)
      let delay = calculateRetryDelay(attempt, retryConfig.baseDelay, retryConfig.maxDelay);

      if (error instanceof RateLimitError) {
        const retryAfter = getRetryAfter(error);
        if (retryAfter) {
          delay = retryAfter * 1000; // Convert to milliseconds
        }
      }

      logger.warn(
        `Graph API request failed (attempt ${attempt + 1}/${retryConfig.maxRetries + 1}), retrying in ${delay}ms: ${lastError.message}`
      );

      // Wait before retry
      await new Promise((resolve) => setTimeout(resolve, delay));
    }
  }

  // All retries exhausted
  if (lastError) {
    throw lastError;
  }

  throw new Error('Unknown error in callGraphWithRetry');
}

/**
 * Invalidate cache for write operations
 */
function invalidateCacheForWrite(endpoint: string): void {
  // Invalidate related caches based on endpoint
  if (endpoint.includes('/messages') || endpoint.includes('/sendMail')) {
    // Invalidate email-related caches
    apiCache.invalidate('GET.*messages');
    apiCache.invalidate('GET.*mailFolders');
    logger.debug('Invalidated email cache due to write operation', { endpoint });
  } else if (endpoint.includes('/events') || endpoint.includes('/calendar')) {
    // Invalidate calendar-related caches
    apiCache.invalidate('GET.*events');
    apiCache.invalidate('GET.*calendar');
    logger.debug('Invalidated calendar cache due to write operation', { endpoint });
  } else if (endpoint.includes('/drive') || endpoint.includes('/items')) {
    // Invalidate file-related caches
    apiCache.invalidate('GET.*drive');
    apiCache.invalidate('GET.*items');
    logger.debug('Invalidated files cache due to write operation', { endpoint });
  } else if (endpoint.includes('/contacts')) {
    // Invalidate contact-related caches
    apiCache.invalidate('GET.*contacts');
    apiCache.invalidate('GET.*users');
    logger.debug('Invalidated contacts cache due to write operation', { endpoint });
  } else {
    // Broad invalidation for unknown endpoints
    apiCache.clear();
    logger.debug('Cleared all cache due to write operation', { endpoint });
  }
}

/**
 * Helper function to call Graph API endpoints (backward compatible)
 * Wraps callGraphWithRetry with default settings
 */
async function callGraph(
  graphClient: GraphClient,
  method: 'GET' | 'POST' | 'PATCH' | 'DELETE' | '',
  endpoint: string,
  queryParams?: Record<string, string>,
  body?: unknown,
  headers?: Record<string, string>
): Promise<string> {
  // Invalidate cache for write operations
  if (method !== 'GET' && method !== '') {
    invalidateCacheForWrite(endpoint);
  }

  return callGraphWithRetry(graphClient, method, endpoint, queryParams, body, headers);
}

// ============================================================================
// STANDARDIZED RESPONSE FORMAT
// ============================================================================

interface PaginationInfo {
  currentPage?: number;
  pageSize?: number;
  totalItems?: number;
  totalPages?: number;
  hasNext?: boolean;
  hasPrevious?: boolean;
  nextCursor?: string;
  previousCursor?: string;
  skip?: number;
  top?: number;
}

interface ErrorInfo {
  code?: string;
  message: string;
  details?: unknown;
  retryable?: boolean;
}

interface NLPAnalysis {
  intent?: string;
  service?: string;
  entities?: Array<{ value: string; type: string; confidence?: number }>;
  temporal?: {
    expression: string;
    type: string;
    relativeDays?: number;
  } | null;
  confidence?: number;
}

interface StandardResponseMetadata {
  timestamp: string;
  executionTime: number;
  sources: string[];
  cacheHit: boolean;
  pagination?: PaginationInfo;
  requestId?: string;
}

interface StandardResponse<T> {
  success: boolean;
  data?: T;
  metadata: StandardResponseMetadata;
  errors?: ErrorInfo[];
  suggestions?: string[];
  nlpAnalysis?: NLPAnalysis;
  thinking?: string[];
}

/**
 * Format standard response with metadata
 * Supports profession-based personalization via context
 */
function formatStandardResponse<T>(
  data: T | undefined,
  options: {
    success?: boolean;
    executionTime?: number;
    sources?: string[];
    cacheHit?: boolean;
    pagination?: PaginationInfo;
    errors?: ErrorInfo[];
    suggestions?: string[];
    nlpAnalysis?: NLPAnalysis;
    thinking?: string[];
    requestId?: string;
    responseType?: 'calendar' | 'mail' | 'search' | 'general' | 'business-content';
    professionProfile?: ProfessionProfile;
  } = {}
): StandardResponse<T> {
  const {
    success = true,
    executionTime = 0,
    sources = [],
    cacheHit = false,
    pagination,
    errors,
    suggestions,
    nlpAnalysis,
    thinking,
    requestId,
    responseType,
    professionProfile,
  } = options;

  // Get active profession profile from context or options
  const activeProfile = professionProfile || getProfessionProfile();

  const metadata: StandardResponseMetadata = {
    timestamp: new Date().toISOString(),
    executionTime,
    sources,
    cacheHit,
    ...(pagination && { pagination }),
    ...(requestId && { requestId }),
  };

  // Add profession info to metadata if available
  if (activeProfile) {
    (
      metadata as StandardResponseMetadata & { professionProfile?: { id: string; name: string } }
    ).professionProfile = {
      id: activeProfile.id,
      name: activeProfile.name,
    };
  }

  const response: StandardResponse<T> = {
    success,
    ...(data !== undefined && { data }),
    metadata,
    ...(errors && errors.length > 0 && { errors }),
    ...(suggestions && suggestions.length > 0 && { suggestions }),
    ...(nlpAnalysis && { nlpAnalysis }),
    ...(thinking && thinking.length > 0 && { thinking }),
  };

  // Add profession greeting if available
  if (activeProfile && responseType) {
    const greeting = getProfessionGreeting(responseType, { professionProfile: activeProfile });
    if (greeting) {
      (response as StandardResponse<T> & { _professionGreeting?: string })._professionGreeting =
        greeting;
    }
  }

  // Format as Tool output and remove unnecessary information
  return formatToolResponse<StandardResponse<T>>(response);
}

/**
 * Format response as Tool output and remove unnecessary information
 * Automatically removes:
 * - OData properties (@odata.*)
 * - Internal metadata (executionTime, cacheHit, requestId, professionProfile, etc.)
 * - Debug information
 * - Redundant fields
 */
export function formatToolResponse<T>(response: unknown): T {
  if (response === null || response === undefined) {
    return response as T;
  }

  // If it's already a string (JSON stringified), parse it first
  if (typeof response === 'string') {
    try {
      response = JSON.parse(response);
    } catch {
      // Not JSON, return as is
      return response as T;
    }
  }

  // Deep clone to avoid mutating original
  const clone = JSON.parse(JSON.stringify(response));

  /**
   * Remove unnecessary properties from object
   */
  const removeUnnecessaryProps = (obj: unknown): unknown => {
    if (obj === null || obj === undefined) {
      return obj;
    }

    if (Array.isArray(obj)) {
      return obj.map(removeUnnecessaryProps);
    }

    if (typeof obj !== 'object') {
      return obj;
    }

    const cleaned: Record<string, unknown> = {};

    for (const [key, value] of Object.entries(obj)) {
      // Remove OData properties
      if (key.startsWith('@odata.')) {
        continue;
      }

      // Remove internal metadata fields
      if (
        key === 'executionTime' ||
        key === 'cacheHit' ||
        key === 'requestId' ||
        key === 'professionProfile' ||
        key === '_professionGreeting' ||
        key === '_headers' ||
        key === '_etag' ||
        key === '_meta'
      ) {
        continue;
      }

      // Remove metadata object if it only contains unnecessary fields
      if (key === 'metadata') {
        if (value && typeof value === 'object' && !Array.isArray(value)) {
          const meta = value as Record<string, unknown>;
          const cleanedMeta: Record<string, unknown> = {};

          // Keep only essential metadata
          if (meta.timestamp) cleanedMeta.timestamp = meta.timestamp;
          if (meta.pagination) cleanedMeta.pagination = removeUnnecessaryProps(meta.pagination);
          if (meta.sources && Array.isArray(meta.sources) && meta.sources.length > 0) {
            cleanedMeta.sources = meta.sources;
          }

          // Only add metadata if it has meaningful content
          if (Object.keys(cleanedMeta).length > 0) {
            cleaned[key] = cleanedMeta;
          }
        }
        continue;
      }

      // Remove thinking process if it's empty or only contains debug info
      if (key === 'thinking') {
        if (Array.isArray(value) && value.length > 0) {
          // Keep only meaningful thinking steps (filter out debug info)
          const meaningfulThinking = (value as string[]).filter(
            (step) =>
              step &&
              !step.includes('processing') &&
              !step.includes('Response received') &&
              !step.includes('Response contains') &&
              !step.includes('Response has')
          );
          if (meaningfulThinking.length > 0) {
            cleaned[key] = meaningfulThinking;
          }
        }
        continue;
      }

      // Remove nlpAnalysis if it's too verbose
      if (key === 'nlpAnalysis') {
        if (value && typeof value === 'object' && !Array.isArray(value)) {
          const nlp = value as Record<string, unknown>;
          const cleanedNlp: Record<string, unknown> = {};

          // Keep only essential NLP info
          if (nlp.intent) cleanedNlp.intent = nlp.intent;
          if (nlp.confidence !== undefined) cleanedNlp.confidence = nlp.confidence;
          if (nlp.entities && Array.isArray(nlp.entities) && nlp.entities.length > 0) {
            cleanedNlp.entities = removeUnnecessaryProps(nlp.entities);
          }

          if (Object.keys(cleanedNlp).length > 0) {
            cleaned[key] = cleanedNlp;
          }
        }
        continue;
      }

      // Recursively clean nested objects
      cleaned[key] = removeUnnecessaryProps(value);
    }

    return cleaned;
  };

  return removeUnnecessaryProps(clone) as T;
}

/**
 * Get output format from environment or default to JSON
 */
function getOutputFormat(): 'json' | 'toon' {
  const envFormat = process.env.MS365_MCP_OUTPUT_FORMAT;
  return envFormat === 'toon' ? 'toon' : 'json';
}

/**
 * Format and return response as Tool output
 * Wrapper function that formats any response (string or object) and removes unnecessary information
 * Uses TOON format if enabled for 30-60% token reduction
 */
function formatAndReturnToolResponse(response: string | unknown, thinking?: string[]): string {
  // If response is already a string, try to parse it
  let parsed: unknown;
  if (typeof response === 'string') {
    try {
      parsed = JSON.parse(response);
    } catch {
      // Not JSON, return as is (but still format if it's a JSON string in text)
      return response;
    }
  } else {
    parsed = response;
  }

  // Format as Tool output
  const formatted = formatToolResponse(parsed);

  // Convert to string using TOON format if enabled, otherwise JSON
  const outputFormat = getOutputFormat();
  let formattedString: string;

  if (outputFormat === 'toon') {
    try {
      formattedString = toonEncode(formatted);
    } catch (error) {
      logger.warn(`Failed to encode as TOON, falling back to JSON: ${error}`);
      formattedString = JSON.stringify(formatted, null, 2);
    }
  } else {
    formattedString = JSON.stringify(formatted, null, 2);
  }

  // Add thinking if provided
  if (thinking && thinking.length > 0) {
    return addThinkingToResponse(formattedString, thinking);
  }

  return formattedString;
}

/**
 * Extract pagination info from Graph API response
 */
function extractPaginationInfo(
  response: unknown,
  skip?: number,
  top?: number
): PaginationInfo | undefined {
  if (typeof response !== 'object' || response === null) {
    return undefined;
  }

  const obj = response as Record<string, unknown>;
  const value = obj.value;

  if (!Array.isArray(value)) {
    return undefined;
  }

  const totalItems = value.length;
  const pageSize = top || 25;
  const currentPage = skip ? Math.floor(skip / pageSize) + 1 : 1;
  const totalPages = Math.ceil(totalItems / pageSize);

  // Check for @odata.nextLink for cursor-based pagination
  const nextLink = obj['@odata.nextLink'] as string | undefined;
  const hasNext = !!nextLink || totalItems >= pageSize;

  return {
    currentPage,
    pageSize,
    totalItems,
    totalPages,
    hasNext,
    hasPrevious: currentPage > 1,
    ...(skip !== undefined && { skip }),
    ...(top !== undefined && { top }),
  };
}

// ============================================================================
// RATE LIMITING
// ============================================================================

interface RequestQueue {
  queue: Array<{
    operation: () => Promise<unknown>;
    resolve: (value: unknown) => void;
    reject: (error: Error) => void;
    priority: 'high' | 'medium' | 'low';
    timestamp: number;
  }>;
  processing: boolean;
  lastRequestTime: number;
  minDelay: number; // Minimum delay between requests (ms)
}

class RateLimiter {
  private queues = new Map<string, RequestQueue>();
  private defaultMinDelay = 100; // 100ms default delay

  /**
   * Execute operation with rate limiting
   */
  async execute<T>(
    key: string,
    operation: () => Promise<T>,
    priority: 'high' | 'medium' | 'low' = 'medium'
  ): Promise<T> {
    let queue = this.queues.get(key);

    if (!queue) {
      queue = {
        queue: [],
        processing: false,
        lastRequestTime: 0,
        minDelay: this.defaultMinDelay,
      };
      this.queues.set(key, queue);
    }

    return new Promise<T>((resolve, reject) => {
      // Add to queue with priority
      queue!.queue.push({
        operation: operation as () => Promise<unknown>,
        resolve: resolve as (value: unknown) => void,
        reject: reject,
        priority,
        timestamp: Date.now(),
      });

      // Sort queue by priority (high first, then by timestamp)
      queue!.queue.sort((a, b) => {
        const priorityOrder = { high: 0, medium: 1, low: 2 };
        const priorityDiff = priorityOrder[a.priority] - priorityOrder[b.priority];
        if (priorityDiff !== 0) return priorityDiff;
        return a.timestamp - b.timestamp;
      });

      // Process queue
      this.processQueue(key);
    });
  }

  /**
   * Process queue for a given key
   */
  private async processQueue(key: string): Promise<void> {
    const queue = this.queues.get(key);
    if (!queue || queue.processing || queue.queue.length === 0) {
      return;
    }

    queue.processing = true;

    while (queue.queue.length > 0) {
      const item = queue.queue.shift()!;

      // Calculate delay based on last request time
      const now = Date.now();
      const timeSinceLastRequest = now - queue.lastRequestTime;
      const delay = Math.max(0, queue.minDelay - timeSinceLastRequest);

      if (delay > 0) {
        await new Promise((resolve) => setTimeout(resolve, delay));
      }

      try {
        queue.lastRequestTime = Date.now();
        const result = await item.operation();
        item.resolve(result);
      } catch (error) {
        item.reject(error instanceof Error ? error : new Error(String(error)));
      }
    }

    queue.processing = false;
  }

  /**
   * Update rate limit for a key (e.g., after receiving 429)
   */
  updateRateLimit(key: string, minDelay: number): void {
    const queue = this.queues.get(key);
    if (queue) {
      queue.minDelay = minDelay;
      logger.warn(`Updated rate limit for ${key}: ${minDelay}ms delay`);
    }
  }

  /**
   * Clear rate limit for a key
   */
  clear(key: string): void {
    this.queues.delete(key);
  }

  /**
   * Clear all rate limits
   */
  clearAll(): void {
    this.queues.clear();
  }
}

// Global rate limiter instance
const rateLimiter = new RateLimiter();

// ============================================================================
// BATCH OPERATIONS & REQUEST DEDUPLICATION
// ============================================================================

interface PendingRequest {
  key: string;
  promise: Promise<unknown>;
  resolve: (value: unknown) => void;
  reject: (error: Error) => void;
  timestamp: number;
}

/**
 * Request deduplication - prevents duplicate API calls
 */
class RequestDeduplicator {
  private pendingRequests = new Map<string, PendingRequest>();
  private readonly maxAge = 5000; // 5 seconds max age for pending requests

  /**
   * Execute request with deduplication
   */
  async execute<T>(key: string, operation: () => Promise<T>): Promise<T> {
    // Check if same request is already pending
    const pending = this.pendingRequests.get(key);
    if (pending) {
      const age = Date.now() - pending.timestamp;
      if (age < this.maxAge) {
        logger.debug(`Deduplicating request: ${key} (age: ${age}ms)`);
        return pending.promise as Promise<T>;
      } else {
        // Request is too old, remove it
        this.pendingRequests.delete(key);
      }
    }

    // Create new request
    let resolve!: (value: T) => void;
    let reject!: (error: Error) => void;
    const promise = new Promise<T>((res, rej) => {
      resolve = res;
      reject = rej;
    });

    const pendingRequest: PendingRequest = {
      key,
      promise: promise as Promise<unknown>,
      resolve: resolve as (value: unknown) => void,
      reject: reject as (error: Error) => void,
      timestamp: Date.now(),
    };

    this.pendingRequests.set(key, pendingRequest);

    // Execute operation
    try {
      const result = await operation();
      resolve!(result);
      this.pendingRequests.delete(key);
      return result;
    } catch (error) {
      reject!(error instanceof Error ? error : new Error(String(error)));
      this.pendingRequests.delete(key);
      throw error;
    }
  }

  /**
   * Clean up old pending requests
   */
  cleanup(): void {
    const now = Date.now();
    for (const [key, request] of this.pendingRequests.entries()) {
      if (now - request.timestamp > this.maxAge) {
        this.pendingRequests.delete(key);
      }
    }
  }

  /**
   * Clear all pending requests
   */
  clear(): void {
    this.pendingRequests.clear();
  }
}

// Global request deduplicator
const requestDeduplicator = new RequestDeduplicator();

/**
 * Build a relative Graph URL with optional query params (for batch requests).
 */
function buildGraphBatchUrl(path: string, queryParams?: Record<string, string>): string {
  if (!queryParams || Object.keys(queryParams).length === 0) return path;
  const qs = Object.entries(queryParams)
    .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
    .join('&');
  return `${path}?${qs}`;
}

/**
 * Batch API calls with deduplication and error handling (parallel per-request).
 * Use graphClient.performBatch() for true Graph $batch when all calls are GETs.
 */
async function batchAPICalls<T>(
  calls: Array<{
    key: string;
    operation: () => Promise<T>;
    description?: string;
  }>
): Promise<Array<{ key: string; result?: T; error?: Error; description?: string }>> {
  // Clean up old requests
  requestDeduplicator.cleanup();

  // Execute all calls in parallel with deduplication
  const promises = calls.map(async ({ key, operation, description }) => {
    try {
      const result = await requestDeduplicator.execute(key, operation);
      return { key, result, description };
    } catch (error) {
      return {
        key,
        error: error instanceof Error ? error : new Error(String(error)),
        description,
      };
    }
  });

  return Promise.all(promises);
}

// ============================================================================
// NLP QUERY OPTIMIZATION
// ============================================================================

/**
 * Optimize query using NLP analysis
 */
function optimizeQueryWithNLP(query: string): {
  optimizedQuery: string;
  filters?: Record<string, unknown>;
  nlpAnalysis: NLPAnalysis;
} {
  const decomposed = nlpEnhancer.decomposeQuery(query);

  let optimizedQuery = query;

  // Extract temporal filters
  const filters: Record<string, unknown> = {};
  if (decomposed.temporal) {
    const now = new Date();
    if (decomposed.temporal.relativeDays) {
      const days = decomposed.temporal.relativeDays;
      const date = new Date(now);
      date.setDate(date.getDate() + days);
      filters.dateFilter = date.toISOString();
    }
  }

  // Optimize query string (remove stopwords, expand synonyms)
  const normalized = nlpEnhancer.normalizeQuery(query);
  if (normalized.normalized !== query) {
    optimizedQuery = normalized.normalized;
  }

  const nlpAnalysis: NLPAnalysis = {
    intent: decomposed.intent.type,
    service: decomposed.ms365Context?.service,
    entities: decomposed.entities.map((e) => ({
      value: e.value,
      type: e.type,
      confidence: e.confidence,
    })),
    temporal: decomposed.temporal
      ? {
          expression: decomposed.temporal.expression,
          type: decomposed.temporal.type,
          relativeDays: decomposed.temporal.relativeDays,
        }
      : null,
    confidence: decomposed.confidence,
  };

  return {
    optimizedQuery,
    ...(Object.keys(filters).length > 0 && { filters }),
    nlpAnalysis,
  };
}

// ============================================================================
// GENERIC QUERY OPTIMIZATION HELPER (USED BY ALL HANDLERS)
// ============================================================================

/**
 * Optimize query for search using the QueryOptimizer
 *
 * This is the single entry point for all handlers to optimize queries.
 * It combines NLP analysis, history-based learning, and pattern detection.
 *
 * @param query - Original query string
 * @param context - Tool context with tool name, entity types, and userIdHash
 * @returns Optimized query with metadata, optimization steps, and variants
 */
function optimizeQueryForSearch(query: string, context: OptimizationContext): OptimizedQuery {
  const optimizer = getQueryOptimizer();
  return optimizer.optimizeQuery(query, context);
}

/**
 * Record the result of a query optimization for learning
 *
 * Call this after executing a query to record whether the optimization was successful.
 * This feeds the learning system so future queries benefit from past experience.
 *
 * @param originalQuery - Original unmodified query
 * @param optimizedQuery - The optimized query that was executed
 * @param success - Whether the query returned results
 * @param userIdHash - Hashed user ID for isolation
 * @param toolContext - Which tool executed the query
 */
function recordQueryOptimizationResult(
  originalQuery: string,
  optimizedQuery: string,
  success: boolean,
  userIdHash: string,
  toolContext: string = 'search'
): void {
  try {
    const queryStore = getQueryStore();
    queryStore.recordQueryVariant(originalQuery, optimizedQuery, success, userIdHash, toolContext);
  } catch (error) {
    logger.debug('Failed to record query optimization result', { error });
  }
}

/**
 * Add optimization steps to thinking output
 * @param thinking - Array of thinking strings
 * @param result - The optimization result
 */
function addOptimizationThinking(thinking: string[], result: OptimizedQuery): void {
  if (result.optimizations.length > 0) {
    thinking.push(`🧠 Auto Query Optimization (${result.optimizations.length} step(s)):`);
    for (const step of result.optimizations) {
      thinking.push(
        `  → ${step.type}: ${step.description} (confidence: ${Math.round(step.confidence * 100)}%)`
      );
    }
    if (result.optimizedQuery !== result.originalQuery) {
      thinking.push(`  📝 Optimized: "${result.originalQuery}" → "${result.optimizedQuery}"`);
    }
    if (result.variants.length > 0) {
      thinking.push(
        `  🔀 Variants: ${result.variants
          .slice(0, 3)
          .map((v) => `"${v}"`)
          .join(', ')}`
      );
    }
    if (result.learnedFromHistory) {
      thinking.push(`  📚 Used learned patterns from history`);
    }
  }
}

/**
 * Build a short optimization summary for the LLM (included in search response when optimization ran).
 */
function buildOptimizationSummary(result: OptimizedQuery): string {
  const parts: string[] = [];
  if (result.originalQuery !== result.optimizedQuery) {
    parts.push(`Query optimized: "${result.originalQuery}" → "${result.optimizedQuery}"`);
  }
  if (result.learnedFromHistory) {
    parts.push('learned pattern from history');
  }
  if (result.optimizations.length > 0) {
    parts.push(result.optimizations.map((s) => `${s.type}: ${s.description}`).join('; '));
  }
  if (result.variants.length > 0) {
    parts.push(`${result.variants.length} variant(s) tried if no results`);
  }
  return parts.join('. ');
}

// Common schemas
const paginationSchema = {
  top: z.number().optional().default(25).describe('Maximum number of items to return'),
  skip: z.number().optional().default(0).describe('Number of items to skip for pagination'),
};

const filterSchema = {
  filter: z.string().optional().describe('OData filter expression'),
  search: z.string().optional().describe('Search query string'),
  orderby: z.string().optional().describe('OData orderby expression'),
};

// Date helper functions for calendar queries
function setStartOfDay(date: Date): Date {
  const result = new Date(date);
  result.setHours(0, 0, 0, 0);
  return result;
}

function setEndOfDay(date: Date): Date {
  const result = new Date(date);
  result.setHours(23, 59, 59, 999);
  return result;
}

// Read-only mode check helper
function checkReadOnly(readOnly: boolean, action: string): void {
  if (readOnly) {
    throw new Error(
      `Action "${action}" is a write operation and is blocked in read-only mode. ` +
        'Set READ_ONLY=0 or MS365_MCP_READ_ONLY=false to enable write operations.'
    );
  }
}

// Write actions that require readOnly check
const WRITE_ACTIONS = new Set([
  // Email
  'send',
  'delete',
  'move',
  'reply',
  'forward',
  // Calendar
  'create-event',
  'update-event',
  'delete-event',
  // Tasks
  'create-task',
  'update-task',
  'delete-task',
  // Contacts
  'create-contact',
  'update-contact',
  'delete-contact',
]);

// ============================================================================
// 1. EMAIL SUPER-TOOL
// ============================================================================
const emailActionsRead = [
  'list', // List messages from inbox or folder
  'get', // Get a specific message
  'folders', // List mail folders
  'child-folders', // List child folders of a mail folder
  'attachments', // List/get attachments
  'search', // Search messages
] as const;

const emailActionsWrite = [
  'send', // Send a new email
  'reply', // Reply to an email
  'delete', // Delete an email
  'move', // Move email to folder
] as const;

// Build schema dynamically based on readOnly mode
function getEmailActions(readOnly: boolean) {
  if (readOnly) {
    return z.enum(emailActionsRead);
  }
  return z.enum([...emailActionsRead, ...emailActionsWrite]);
}

const emailActions = z.enum([
  // Read operations
  'list',
  'get',
  'folders',
  'child-folders',
  'attachments',
  'search',
  // Business content extraction
  'extract-business-content',
  'extract-project',
  'extract-customer',
  'extract-meeting',
  'extract-document',
  'extract-sales',
  'extract-hr',
  // Write operations (blocked in read-only mode)
  'send',
  'reply',
  'delete',
  'move',
]);

const emailSchema = z.object({
  action: emailActions.describe(
    'The email operation: list, get, folders, child-folders, attachments, search, extract-business-content, extract-project, extract-customer, extract-meeting, extract-document, extract-sales, extract-hr (read) | send, reply, delete, move (write)'
  ),
  // Identifiers
  messageId: z
    .string()
    .optional()
    .describe('Message ID (required for get, attachments, reply, delete, move, extract-* actions)'),
  folderId: z.string().optional().describe('Folder ID to list messages from or move to'),
  attachmentId: z.string().optional().describe('Attachment ID (for getting specific attachment)'),
  // For send/reply
  to: z.string().optional().describe('Recipient email address(es), comma-separated (for send)'),
  subject: z.string().optional().default('(No subject)').describe('Email subject (for send)'),
  body: z.string().optional().default('').describe('Email body content (for send/reply)'),
  // For extract actions
  extractType: z
    .string()
    .optional()
    .describe(
      'Specific extract type or "auto" for automatic detection (for extract-business-content)'
    ),
  includeMetadata: z
    .boolean()
    .optional()
    .default(true)
    .describe('Include metadata in extraction result'),
  includeEntities: z
    .boolean()
    .optional()
    .default(true)
    .describe('Include entities in extraction result'),
  includeSummary: z
    .boolean()
    .optional()
    .default(true)
    .describe('Include summary in extraction result'),
  // Filters
  ...filterSchema,
  ...paginationSchema,
});

type EmailInput = z.infer<typeof emailSchema>;

async function handleEmail(
  input: EmailInput,
  graphClient: GraphClient,
  readOnly: boolean
): Promise<string> {
  const thinking: string[] = [];

  // Check write operations against readOnly mode
  if (['send', 'reply', 'delete', 'move'].includes(input.action)) {
    checkReadOnly(readOnly, input.action);
  }

  switch (input.action) {
    case 'list': {
      const startTime = Date.now();
      thinking.push(
        `Listing emails${input.folderId ? ` from folder ${input.folderId}` : ' from inbox'}`
      );
      const endpoint = input.folderId
        ? `/me/mailFolders/${input.folderId}/messages`
        : '/me/messages';
      const params: Record<string, string> = { $top: String(input.top || 25) };
      if (input.filter) params.$filter = input.filter;
      if (input.search) params.$search = formatSearchQuery(input.search);
      if (input.orderby) params.$orderby = input.orderby;
      if (input.skip) params.$skip = String(input.skip);

      const result = await callGraph(graphClient, 'GET', endpoint, params);
      const parsedResult = JSON.parse(result);
      const executionTime = Date.now() - startTime;

      // Extract pagination info
      const pagination = extractPaginationInfo(parsedResult, input.skip, input.top);

      // Format mail response with profession-specific formatting
      if (isMailResponse(parsedResult)) {
        const formatted = formatMailResponse(parsedResult);
        // Use profession-specific text formatting
        const formattedText = mailResponseToTextByProfession(formatted);

        // Add metadata to response
        const responseWithMetadata = formatStandardResponse(
          { formatted: formattedText, raw: parsedResult },
          {
            executionTime,
            sources: ['email'],
            cacheHit: false,
            pagination,
            responseType: 'mail',
            suggestions: [
              '💡 Use "email" tool with action "get" to view email details',
              '💡 Use "email" tool with action "search" for advanced search',
            ],
          }
        );

        return formatAndReturnToolResponse(responseWithMetadata, thinking);
      }

      // Fallback: return with metadata
      const responseWithMetadata = formatStandardResponse(parsedResult, {
        executionTime,
        sources: ['email'],
        cacheHit: false,
        pagination,
      });

      return formatAndReturnToolResponse(responseWithMetadata, thinking);
    }

    case 'get': {
      if (!input.messageId) throw new Error('messageId is required for action "get"');
      thinking.push(`Getting email with ID: ${input.messageId}`);
      const result = await callGraph(graphClient, 'GET', `/me/messages/${input.messageId}`);
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'extract-business-content':
    case 'extract-project':
    case 'extract-customer':
    case 'extract-meeting':
    case 'extract-document':
    case 'extract-sales':
    case 'extract-hr': {
      if (!input.messageId) throw new Error('messageId is required for extract actions');
      thinking.push(`Extracting business content from email: ${input.messageId}`);

      // Get email content
      const emailResult = await callGraph(graphClient, 'GET', `/me/messages/${input.messageId}`, {
        $select: 'id,subject,body,bodyPreview,receivedDateTime,from,toRecipients,ccRecipients',
      });
      const emailData = JSON.parse(emailResult);

      // Extract body content
      const bodyContent = emailData.body?.content || emailData.bodyPreview || '';
      const contentType = emailData.body?.contentType || 'text';

      // Sanitize HTML if needed
      let textContent = bodyContent;
      if (contentType === 'html') {
        // Use sanitizeHtml from BaseExtractor
        const { ProjectExtractor } = await import('./utils/content-extractor.js');
        const tempExtractor = new ProjectExtractor();
        // Access protected method via type assertion
        textContent = (
          tempExtractor as unknown as { sanitizeHtml: (html: string) => string }
        ).sanitizeHtml(bodyContent);
      }

      // Check cache first
      const cache = getExtractionCache();
      const extractOptions: ExtractorOptions = {
        includeMetadata: input.includeMetadata !== false,
        includeEntities: input.includeEntities !== false,
        includeSummary: input.includeSummary !== false,
      };
      const cacheKey = generateExtractionCacheKey(input.messageId, extractOptions);
      const cachedResult = cache.get(cacheKey);

      if (cachedResult) {
        thinking.push(`✅ Using cached extraction result`);
        const responseWithMetadata = formatStandardResponse(
          { extracted: cachedResult, raw: emailData },
          {
            executionTime: 0,
            sources: ['email', 'content-extraction', 'cache'],
            cacheHit: true,
            responseType: 'business-content',
            suggestions: [
              '💡 Use "email" tool with action "get" to view full email',
              '💡 Use specific extract actions for targeted extraction',
            ],
          }
        );
        return formatAndReturnToolResponse(responseWithMetadata, thinking);
      }

      // Determine extraction type
      let extractType: DocumentType | undefined;
      if (input.action === 'extract-project') extractType = 'project_plan';
      else if (input.action === 'extract-customer') extractType = 'customer_info';
      else if (input.action === 'extract-meeting') extractType = 'meeting_notes';
      else if (input.action === 'extract-document') extractType = 'invoice';
      else if (input.action === 'extract-sales') extractType = 'offer';
      else if (input.action === 'extract-hr') extractType = 'onboarding';
      else if (input.extractType && input.extractType !== 'auto') {
        extractType = input.extractType as DocumentType;
      }

      // Perform extraction
      const patternExtractor = getPatternBasedExtractor();
      const extracted = patternExtractor.extract(textContent, extractType, extractOptions);

      // Build complete business content extraction result
      const entityRegistry = new EntityExtractorRegistry();
      const metadataExtractor = new MetadataExtractor();
      const summaryGenerator = new SummaryGenerator();

      const businessExtraction: BusinessContentExtraction = {
        detectedType: extracted.type,
        confidence: extracted.confidence,
        extracted: {
          project:
            extracted.type.includes('project') || extracted.type === 'roadmap'
              ? (extracted.content as ProjectContent)
              : undefined,
          customer:
            extracted.type.includes('customer') ||
            extracted.type === 'contract' ||
            extracted.type === 'proposal'
              ? (extracted.content as CustomerContent)
              : undefined,
          meeting:
            extracted.type.includes('meeting') || extracted.type === 'action_items'
              ? (extracted.content as MeetingContent)
              : undefined,
          document:
            extracted.type === 'invoice' || extracted.type === 'report'
              ? (extracted.content as DocumentContent)
              : undefined,
          sales:
            extracted.type === 'offer' ||
            extracted.type === 'budget' ||
            extracted.type === 'forecast'
              ? (extracted.content as SalesContent)
              : undefined,
          hr:
            extracted.type === 'onboarding' ||
            extracted.type === 'review' ||
            extracted.type === 'application'
              ? (extracted.content as HRContent)
              : undefined,
        },
        metadata: {
          priorities:
            input.includeMetadata !== false
              ? metadataExtractor.extractPriorities(textContent)
              : undefined,
          statuses:
            input.includeMetadata !== false
              ? metadataExtractor.extractStatuses(textContent)
              : undefined,
          deadlines:
            input.includeMetadata !== false
              ? metadataExtractor.extractDeadlines(textContent)
              : undefined,
          tags:
            input.includeMetadata !== false
              ? metadataExtractor.extractTags(textContent)
              : undefined,
        },
        entities: input.includeEntities !== false ? entityRegistry.extractAll(textContent) : {},
        summary:
          input.includeSummary !== false
            ? {
                actionItems: summaryGenerator.generateActionItems(textContent),
                decisions: summaryGenerator.generateDecisions(textContent),
                keyPoints: summaryGenerator.generateKeyPoints(textContent),
              }
            : undefined,
      };

      thinking.push(
        `✅ Extracted business content: ${extracted.type} (confidence: ${(extracted.confidence * 100).toFixed(1)}%)`
      );

      // Cache the result
      cache.set(cacheKey, businessExtraction);

      const responseWithMetadata = formatStandardResponse(
        { extracted: businessExtraction, raw: emailData },
        {
          executionTime: 0,
          sources: ['email', 'content-extraction'],
          cacheHit: false,
          responseType: 'business-content',
          suggestions: [
            '💡 Use "email" tool with action "get" to view full email',
            '💡 Use specific extract actions for targeted extraction',
          ],
        }
      );

      return formatAndReturnToolResponse(responseWithMetadata, thinking);
    }

    case 'folders': {
      thinking.push('Listing mail folders');
      const params: Record<string, string> = { $top: String(input.top || 50) };
      const result = await callGraph(graphClient, 'GET', '/me/mailFolders', params);
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'child-folders': {
      if (!input.folderId) throw new Error('folderId is required for action "child-folders"');
      thinking.push(`Listing child folders of folder: ${input.folderId}`);
      const params: Record<string, string> = { $top: String(input.top || 50) };
      const result = await callGraph(
        graphClient,
        'GET',
        `/me/mailFolders/${input.folderId}/childFolders`,
        params
      );
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'attachments': {
      if (!input.messageId) throw new Error('messageId is required for action "attachments"');
      thinking.push(`Getting attachments for message: ${input.messageId}`);
      const endpoint = input.attachmentId
        ? `/me/messages/${input.messageId}/attachments/${input.attachmentId}`
        : `/me/messages/${input.messageId}/attachments`;
      const result = await callGraph(graphClient, 'GET', endpoint);
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'search': {
      if (!input.search) throw new Error('search query is required for action "search"');

      // Automatic query optimization for email search
      const startTime = Date.now();
      const currentUserId = getUserId();
      const userIdHash = currentUserId ? getQueryStore().hashUserId(currentUserId) : undefined;
      const optimized = optimizeQueryForSearch(input.search, {
        tool: 'email',
        entityTypes: ['message'],
        userIdHash,
      });
      thinking.push(`🔍 Searching emails for: "${input.search}"`);
      addOptimizationThinking(thinking, optimized);
      if (optimized.nlpAnalysis.intent) {
        thinking.push(`📊 Detected intent: ${optimized.nlpAnalysis.intent}`);
      }

      const params: Record<string, string> = {
        $search: formatSearchQuery(optimized.optimizedQuery),
        $top: String(input.top || 25),
      };

      // Microsoft Graph API limitation: $orderby is NOT supported with $search
      // Search results are automatically ordered by relevance
      // Note: input.orderby is intentionally not used here

      // Apply temporal filters if NLP detected them
      if (optimized.filters?.dateFilter) {
        const dateFilter = optimized.filters.dateFilter as string;
        params.$filter = `receivedDateTime ge ${dateFilter}`;
        thinking.push(`📅 Applying date filter: after ${dateFilter}`);
      }

      const result = await callGraph(graphClient, 'GET', '/me/messages', params);
      const parsedResult = JSON.parse(result);
      const executionTime = Date.now() - startTime;

      // Record optimization result for learning
      const emailSearchSuccess = parsedResult?.value?.length > 0;
      if (userIdHash) {
        recordQueryOptimizationResult(
          input.search,
          optimized.optimizedQuery,
          emailSearchSuccess,
          userIdHash,
          'email'
        );
      }

      // Format mail response with profession-specific formatting
      if (isMailResponse(parsedResult)) {
        const formatted = formatMailResponse(parsedResult);
        // Use profession-specific text formatting
        const formattedText = mailResponseToTextByProfession(formatted);

        // Add NLP insights to response
        const responseWithMetadata = formatStandardResponse(
          { formatted: formattedText, raw: parsedResult },
          {
            executionTime,
            sources: ['email'],
            cacheHit: false,
            responseType: 'mail',
            nlpAnalysis: optimized.nlpAnalysis,
            suggestions: [
              '💡 Use "email" tool with action "get" to view full email details',
              '💡 Use "email" tool with action "list" to browse more emails',
            ],
          }
        );

        return formatAndReturnToolResponse(responseWithMetadata, thinking);
      }

      return formatAndReturnToolResponse(result, thinking);
    }

    // Write operations (blocked in read-only mode - check happens at function start)
    case 'send': {
      const to = input.to?.trim();
      if (!to) throw new Error('to (recipient) is required for action "send"');
      const subject = input.subject ?? '(No subject)';
      const body = input.body ?? '';
      thinking.push(`Sending email to: ${to}`);
      const recipients = to.split(',').map((email) => ({
        emailAddress: { address: email.trim() },
      }));
      const message = {
        subject,
        body: { contentType: 'Text', content: body },
        toRecipients: recipients,
      };
      const result = await callGraph(graphClient, 'POST', '/me/sendMail', undefined, {
        message,
        saveToSentItems: true,
      });
      return addThinkingToResponse(
        result || JSON.stringify({ success: true, message: 'Email sent' }),
        thinking
      );
    }

    case 'reply': {
      if (!input.messageId) throw new Error('messageId is required for action "reply"');
      const replyBody = input.body ?? '';
      thinking.push(`Replying to email: ${input.messageId}`);
      const result = await callGraph(
        graphClient,
        'POST',
        `/me/messages/${input.messageId}/reply`,
        undefined,
        { comment: replyBody }
      );
      return addThinkingToResponse(
        result || JSON.stringify({ success: true, message: 'Reply sent' }),
        thinking
      );
    }

    case 'delete': {
      if (!input.messageId) throw new Error('messageId is required for action "delete"');
      thinking.push(`Deleting email: ${input.messageId}`);
      const result = await callGraph(graphClient, 'DELETE', `/me/messages/${input.messageId}`);
      return addThinkingToResponse(
        result || JSON.stringify({ success: true, message: 'Email deleted' }),
        thinking
      );
    }

    case 'move': {
      if (!input.messageId) throw new Error('messageId is required for action "move"');
      if (!input.folderId) throw new Error('folderId (destination) is required for action "move"');
      thinking.push(`Moving email ${input.messageId} to folder ${input.folderId}`);
      const result = await callGraph(
        graphClient,
        'POST',
        `/me/messages/${input.messageId}/move`,
        undefined,
        { destinationId: input.folderId }
      );
      return formatAndReturnToolResponse(result, thinking);
    }

    default:
      throw new Error(`Unknown email action: ${input.action}`);
  }
}

// ============================================================================
// 2. CALENDAR SUPER-TOOL
// ============================================================================
const calendarActions = z.enum([
  // Read operations
  'list', // List events from primary calendar
  'get', // Get specific event
  'view', // Get calendar view (date range)
  'calendars', // List all calendars
  'specific-calendar', // List events from specific calendar
  // Write operations (blocked in read-only mode)
  'create-event', // Create new event
  'update-event', // Update existing event
  'delete-event', // Delete event
]);

const calendarSchema = z.object({
  action: calendarActions.describe(
    'Calendar operation: list, get, view, calendars (read) | create-event, update-event, delete-event (write)'
  ),
  // Identifiers
  eventId: z.string().optional().describe('Event ID (required for get)'),
  calendarId: z.string().optional().describe('Calendar ID (for specific-calendar action)'),
  // Date range for view and create-event
  startDateTime: z.string().optional().describe('Start date/time (ISO format)'),
  endDateTime: z.string().optional().describe('End date/time (ISO format)'),
  // Timezone
  timezone: z
    .string()
    .optional()
    .default('UTC')
    .describe('Timezone for date/time values (e.g., "Europe/Berlin")'),
  // For create/update event
  subject: z
    .string()
    .optional()
    .default('Untitled Event')
    .describe('Event subject/title (for create-event, update-event)'),
  body: z.string().optional().describe('Event body/description (for create-event, update-event)'),
  location: z.string().optional().describe('Event location (for create-event, update-event)'),
  attendees: z.string().optional().describe('Attendee emails, comma-separated (for create-event)'),
  isOnline: z.boolean().optional().describe('Create as online meeting (for create-event)'),
  // Filters
  ...filterSchema,
  ...paginationSchema,
});

/** Input type for calendar tool; use z.input so partial objects (e.g. auto-execution) are accepted. */
type CalendarInput = z.input<typeof calendarSchema>;

async function handleCalendar(
  input: CalendarInput,
  graphClient: GraphClient,
  readOnly: boolean
): Promise<string> {
  const thinking: string[] = [];
  const headers: Record<string, string> = {};
  if (input.timezone) {
    headers['Prefer'] = `outlook.timezone="${input.timezone}"`;
  }

  // Check write operations against readOnly mode
  if (['create-event', 'update-event', 'delete-event'].includes(input.action)) {
    checkReadOnly(readOnly, input.action);
  }

  switch (input.action) {
    case 'list': {
      const startTime = Date.now();
      thinking.push('Listing calendar events');
      const params: Record<string, string> = { $top: String(input.top || 25) };
      if (input.filter) params.$filter = input.filter;
      if (input.orderby) params.$orderby = input.orderby;
      if (input.skip) params.$skip = String(input.skip);
      const result = await callGraph(graphClient, 'GET', '/me/events', params, undefined, headers);
      const parsedResult = JSON.parse(result);
      const executionTime = Date.now() - startTime;

      // Extract pagination info
      const pagination = extractPaginationInfo(parsedResult, input.skip, input.top);

      // Format calendar response with quick summary
      // Use profession-based formatting if available
      if (isCalendarResponse(parsedResult)) {
        const formatted = formatCalendarResponse(parsedResult);
        // Use profession-specific text formatting
        const formattedText = calendarResponseToTextByProfession(formatted);

        // Add metadata to response
        const responseWithMetadata = formatStandardResponse(
          { formatted: formattedText, raw: parsedResult },
          {
            executionTime,
            sources: ['calendar'],
            cacheHit: false,
            pagination,
            responseType: 'calendar',
            suggestions: [
              '💡 Use "calendar" tool with action "get" to view event details',
              '💡 Use "calendar" tool with action "view" for date range queries',
            ],
          }
        );

        return formatAndReturnToolResponse(responseWithMetadata, thinking);
      }

      // Fallback: return with metadata
      const responseWithMetadata = formatStandardResponse(parsedResult, {
        executionTime,
        sources: ['calendar'],
        cacheHit: false,
        pagination,
      });

      return formatAndReturnToolResponse(responseWithMetadata, thinking);
    }

    case 'get': {
      if (!input.eventId) throw new Error('eventId is required for action "get"');
      thinking.push(`Getting event: ${input.eventId}`);
      const result = await callGraph(
        graphClient,
        'GET',
        `/me/events/${input.eventId}`,
        undefined,
        undefined,
        headers
      );
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'view': {
      const now = new Date();
      const defaultEnd = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
      const startDateTime = input.startDateTime ?? now.toISOString();
      const endDateTime = input.endDateTime ?? defaultEnd.toISOString();
      thinking.push(`Getting calendar view from ${startDateTime} to ${endDateTime}`);
      const params: Record<string, string> = {
        startDateTime,
        endDateTime,
        $top: String(input.top || 50),
      };
      if (input.orderby) params.$orderby = input.orderby;
      const result = await callGraph(
        graphClient,
        'GET',
        '/me/calendarView',
        params,
        undefined,
        headers
      );
      const parsedResult = JSON.parse(result);

      // Format calendar response with profession-specific formatting
      if (isCalendarResponse(parsedResult)) {
        const formatted = formatCalendarResponse(parsedResult, startDateTime, endDateTime);
        // Use profession-specific text formatting
        const formattedText = calendarResponseToTextByProfession(formatted);
        return addThinkingToResponse(formattedText, thinking);
      }

      return formatAndReturnToolResponse(result, thinking);
    }

    case 'calendars': {
      thinking.push('Listing all calendars');
      const result = await callGraph(graphClient, 'GET', '/me/calendars');
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'specific-calendar': {
      if (!input.calendarId)
        throw new Error('calendarId is required for action "specific-calendar"');
      thinking.push(`Listing events from calendar: ${input.calendarId}`);
      const params: Record<string, string> = { $top: String(input.top || 25) };
      if (input.filter) params.$filter = input.filter;
      const result = await callGraph(
        graphClient,
        'GET',
        `/me/calendars/${input.calendarId}/events`,
        params,
        undefined,
        headers
      );
      const parsedResult = JSON.parse(result);

      // Format calendar response with profession-specific formatting
      if (isCalendarResponse(parsedResult)) {
        const formatted = formatCalendarResponse(parsedResult);
        // Use profession-specific text formatting
        const formattedText = calendarResponseToTextByProfession(formatted);
        return addThinkingToResponse(formattedText, thinking);
      }

      return formatAndReturnToolResponse(result, thinking);
    }

    // Write operations (blocked in read-only mode - check happens at function start)
    case 'create-event': {
      const now = new Date();
      const defaultEnd = new Date(now.getTime() + 60 * 60 * 1000);
      const subject = input.subject ?? 'Untitled Event';
      const startDateTime = input.startDateTime ?? now.toISOString();
      const endDateTime = input.endDateTime ?? defaultEnd.toISOString();
      const tz = input.timezone ?? 'UTC';
      thinking.push(`Creating event: ${subject}`);
      const event: Record<string, unknown> = {
        subject,
        start: { dateTime: startDateTime, timeZone: tz },
        end: { dateTime: endDateTime, timeZone: tz },
      };
      if (input.body) event.body = { contentType: 'Text', content: input.body };
      if (input.location) event.location = { displayName: input.location };
      if (input.isOnline) event.isOnlineMeeting = true;
      if (input.attendees) {
        event.attendees = input.attendees.split(',').map((email) => ({
          emailAddress: { address: email.trim() },
          type: 'required',
        }));
      }
      const result = await callGraph(graphClient, 'POST', '/me/events', undefined, event, headers);
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'update-event': {
      if (!input.eventId) throw new Error('eventId is required for update-event');
      thinking.push(`Updating event: ${input.eventId}`);
      const updates: Record<string, unknown> = {};
      if (input.subject) updates.subject = input.subject;
      if (input.body) updates.body = { contentType: 'Text', content: input.body };
      if (input.location) updates.location = { displayName: input.location };
      if (input.startDateTime)
        updates.start = { dateTime: input.startDateTime, timeZone: input.timezone || 'UTC' };
      if (input.endDateTime)
        updates.end = { dateTime: input.endDateTime, timeZone: input.timezone || 'UTC' };
      const result = await callGraph(
        graphClient,
        'PATCH',
        `/me/events/${input.eventId}`,
        undefined,
        updates,
        headers
      );
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'delete-event': {
      if (!input.eventId) throw new Error('eventId is required for delete-event');
      thinking.push(`Deleting event: ${input.eventId}`);
      const result = await callGraph(graphClient, 'DELETE', `/me/events/${input.eventId}`);
      return addThinkingToResponse(
        result || JSON.stringify({ success: true, message: 'Event deleted' }),
        thinking
      );
    }

    default:
      throw new Error(`Unknown calendar action: ${input.action}`);
  }
}

// ============================================================================
// 3. TEAMS SUPER-TOOL
// ============================================================================
const teamsActions = z.enum([
  'list-teams', // List joined teams
  'get-team', // Get team details
  'channels', // List team channels
  'channel-messages', // List channel messages
  'chats', // List chats
  'chat-messages', // List chat messages
]);

const teamsSchema = z.object({
  action: teamsActions.describe('The Teams operation to perform'),
  // Identifiers
  teamId: z
    .string()
    .optional()
    .describe('Team ID (required for get-team, channels, channel-messages)'),
  channelId: z.string().optional().describe('Channel ID (required for channel-messages)'),
  chatId: z
    .string()
    .optional()
    .describe(
      'Chat ID (REQUIRED for chat-messages action - use "chats" action first to get chat IDs)'
    ),
  messageId: z.string().optional().describe('Message ID'),
  // Options
  includeMessages: z
    .boolean()
    .optional()
    .describe('Include last messages for each chat (default: true for chats action)'),
  person: z
    .string()
    .optional()
    .describe(
      'Filter chats by person name or email. When provided, the system will first resolve the person, then fetch only chats with that person.'
    ),
  // Filters
  ...filterSchema,
  ...paginationSchema,
});

type TeamsInput = z.infer<typeof teamsSchema>;

async function handleTeams(
  input: TeamsInput,
  graphClient: GraphClient,
  _readOnly: boolean
): Promise<string> {
  const thinking: string[] = [];
  // Teams operations are read-only in this version

  switch (input.action) {
    case 'list-teams': {
      thinking.push('Listing joined teams');
      const result = await callGraph(graphClient, 'GET', '/me/joinedTeams');
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'get-team': {
      if (!input.teamId) throw new Error('teamId is required');
      thinking.push(`Getting team: ${input.teamId}`);
      const result = await callGraph(graphClient, 'GET', `/teams/${input.teamId}`);
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'channels': {
      if (!input.teamId) throw new Error('teamId is required for channels');
      thinking.push(`Listing channels for team: ${input.teamId}`);
      const result = await callGraph(graphClient, 'GET', `/teams/${input.teamId}/channels`);
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'channel-messages': {
      if (!input.teamId || !input.channelId) {
        throw new Error('teamId and channelId are required for channel-messages');
      }
      thinking.push(`Listing messages in channel: ${input.channelId}`);
      const params: Record<string, string> = {
        $top: String(input.top || 25),
        $orderby: 'createdDateTime desc',
      };
      const result = await callGraph(
        graphClient,
        'GET',
        `/teams/${input.teamId}/channels/${input.channelId}/messages`,
        params
      );
      const messagesData = JSON.parse(result);

      // Format messages for better readability
      if (messagesData.value && Array.isArray(messagesData.value)) {
        const formattedMessages = messagesData.value.map((msg: any) => ({
          id: msg.id,
          from: msg.from?.user?.displayName || msg.from?.application?.displayName || 'Unknown',
          fromEmail: msg.from?.user?.userPrincipalName,
          content: msg.body?.content || '',
          contentType: msg.body?.contentType || 'text',
          createdDateTime: msg.createdDateTime,
          importance: msg.importance,
          subject: msg.subject,
          attachments: msg.attachments?.map((att: any) => ({
            id: att.id,
            name: att.name,
            contentType: att.contentType,
            size: att.size,
          })),
        }));

        const output = {
          teamId: input.teamId,
          channelId: input.channelId,
          totalMessages: formattedMessages.length,
          messages: formattedMessages,
        };

        thinking.push(`Retrieved ${formattedMessages.length} channel messages with content`);
        return addThinkingToResponse(JSON.stringify(output, null, 2), thinking);
      }

      return formatAndReturnToolResponse(result, thinking);
    }

    case 'chats': {
      let chatsData: { value?: any[] };
      let chatsToProcess: any[] = [];

      // If person filter is provided, resolve the person first, then filter chats
      if (input.person) {
        thinking.push(`Resolving person: ${input.person}`);
        const user = await findUser(graphClient, input.person);
        if (!user) {
          const errorResponse = {
            error: 'User not found',
            message: `Could not find a user matching "${input.person}". Try using their full name, email address, or check the spelling.`,
            searchedFor: input.person,
            suggestion: 'Use list-users tool with a search query to find the correct user.',
          };
          thinking.push(`User not found: ${input.person}`);
          return addThinkingToResponse(JSON.stringify(errorResponse, null, 2), thinking);
        }

        thinking.push(`Found user: ${user.displayName} (${user.id})`);
        thinking.push(`Finding chats with ${user.displayName}...`);

        const matchingChats = await findChatsWithUser(
          graphClient,
          user.id,
          user.mail || user.userPrincipalName,
          user.displayName
        );

        if (matchingChats.length === 0) {
          const noChatsResponse = {
            success: true,
            person: {
              name: user.displayName,
              email: user.mail || user.userPrincipalName,
              id: user.id,
            },
            message: `No Teams chats found with ${user.displayName}. You may not have any direct chats with this person.`,
            chatsFound: 0,
            messagesFound: 0,
          };
          thinking.push(`No chats found with ${user.displayName}`);
          return addThinkingToResponse(JSON.stringify(noChatsResponse, null, 2), thinking);
        }

        thinking.push(`Found ${matchingChats.length} chat(s) with ${user.displayName}`);
        chatsToProcess = matchingChats;
      } else {
        // No person filter - list all chats
        thinking.push('Listing chats');
        const params: Record<string, string> = { $top: String(input.top || 25) };
        const chatsResult = await callGraph(graphClient, 'GET', '/me/chats', params);
        chatsData = JSON.parse(chatsResult);
        chatsToProcess = chatsData.value || [];
      }

      // Default: include messages for chats action (unless explicitly disabled)
      const includeMessages = input.includeMessages !== false;

      if (includeMessages && chatsToProcess.length > 0) {
        thinking.push(`Fetching last messages for ${chatsToProcess.length} chat(s)...`);

        // Fetch last messages for each chat (limit to avoid too many requests)
        const chatsWithMessages = await Promise.allSettled(
          chatsToProcess.slice(0, 10).map(async (chat: GraphChat | any) => {
            try {
              const messagesResult = await callGraph(
                graphClient,
                'GET',
                `/me/chats/${chat.id}/messages`,
                { $top: '5', $orderby: 'createdDateTime desc' }
              );
              const messagesData = JSON.parse(messagesResult);
              return {
                ...chat,
                lastMessages: messagesData.value || [],
              };
            } catch (error) {
              logger.warn(`Failed to fetch messages for chat ${chat.id}: ${error}`);
              return {
                ...chat,
                lastMessages: [],
              };
            }
          })
        );

        // Format results
        const formattedChats = chatsWithMessages
          .map((result) => {
            if (result.status === 'fulfilled') {
              const chat = result.value;
              return {
                id: chat.id,
                topic: chat.topic,
                chatType: chat.chatType,
                createdDateTime: chat.createdDateTime,
                lastUpdatedDateTime: chat.lastUpdatedDateTime,
                webUrl: chat.webUrl,
                lastMessages: (chat.lastMessages || []).map((msg: any) => ({
                  id: msg.id,
                  from: msg.from?.user?.displayName || 'Unknown',
                  content: msg.body?.content || '',
                  contentType: msg.body?.contentType || 'text',
                  createdDateTime: msg.createdDateTime,
                  importance: msg.importance,
                })),
              };
            }
            return null;
          })
          .filter(Boolean);

        const output = {
          totalChats: formattedChats.length,
          chats: formattedChats,
        };

        thinking.push(`Retrieved last messages for ${formattedChats.length} chat(s)`);
        return addThinkingToResponse(JSON.stringify(output, null, 2), thinking);
      }

      // If messages are not included, return the chats list
      if (input.person) {
        const output = {
          totalChats: chatsToProcess.length,
          chats: chatsToProcess.map((chat: GraphChat) => ({
            id: chat.id,
            topic: chat.topic,
            chatType: chat.chatType,
            createdDateTime: chat.createdDateTime,
            lastUpdatedDateTime: chat.lastUpdatedDateTime,
            webUrl: chat.webUrl,
          })),
        };
        return addThinkingToResponse(JSON.stringify(output, null, 2), thinking);
      }

      // Fallback: return raw result if no person filter and no messages
      const params: Record<string, string> = { $top: String(input.top || 25) };
      const chatsResult = await callGraph(graphClient, 'GET', '/me/chats', params);
      return addThinkingToResponse(chatsResult, thinking);
    }

    case 'chat-messages': {
      if (!input.chatId) {
        throw new Error(
          'chatId is required for chat-messages action. ' +
            'Use action "chats" first to list your chats and get chat IDs, ' +
            'then use action "chat-messages" with a specific chatId.'
        );
      }
      thinking.push(`Listing messages in chat: ${input.chatId}`);
      const params: Record<string, string> = {
        $top: String(input.top || 25),
        $orderby: 'createdDateTime desc',
      };
      const result = await callGraph(
        graphClient,
        'GET',
        `/me/chats/${input.chatId}/messages`,
        params
      );
      const messagesData = JSON.parse(result);

      // Format messages for better readability
      if (messagesData.value && Array.isArray(messagesData.value)) {
        const formattedMessages = messagesData.value.map((msg: any) => ({
          id: msg.id,
          from: msg.from?.user?.displayName || msg.from?.user?.userPrincipalName || 'Unknown',
          fromEmail: msg.from?.user?.userPrincipalName,
          content: msg.body?.content || '',
          contentType: msg.body?.contentType || 'text',
          createdDateTime: msg.createdDateTime,
          importance: msg.importance,
          subject: msg.subject,
          attachments: msg.attachments?.map((att: any) => ({
            id: att.id,
            name: att.name,
            contentType: att.contentType,
            size: att.size,
          })),
        }));

        const output = {
          chatId: input.chatId,
          totalMessages: formattedMessages.length,
          messages: formattedMessages,
        };

        thinking.push(`Retrieved ${formattedMessages.length} messages with content`);
        return addThinkingToResponse(JSON.stringify(output, null, 2), thinking);
      }

      return formatAndReturnToolResponse(result, thinking);
    }

    default:
      throw new Error(`Unknown teams action: ${input.action}`);
  }
}

// ============================================================================
// 4. FILES SUPER-TOOL
// ============================================================================
const filesActions = z.enum([
  'drives', // List drives
  'list', // List files in folder
  'get', // Get file metadata
  'download', // Download file content
  'root', // Get drive root
  'search', // Search files
]);

const filesSchema = z.object({
  action: filesActions.describe('The files operation to perform'),
  // Identifiers
  driveId: z.string().optional().describe('Drive ID'),
  itemId: z.string().optional().describe('Item (file/folder) ID'),
  path: z.string().optional().describe('Path to file/folder'),
  // Filters
  ...filterSchema,
  ...paginationSchema,
});

type FilesInput = z.infer<typeof filesSchema>;

async function handleFiles(
  input: FilesInput,
  graphClient: GraphClient,
  _readOnly: boolean
): Promise<string> {
  const thinking: string[] = [];
  // Files operations are read-only in this version

  switch (input.action) {
    case 'drives': {
      thinking.push('Listing drives');
      const result = await callGraph(graphClient, 'GET', '/me/drives');
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'list': {
      const driveId = input.driveId || 'me';
      const itemId = input.itemId || 'root';
      thinking.push(`Listing files in ${driveId}/${itemId}`);
      const endpoint =
        driveId === 'me'
          ? `/me/drive/items/${itemId}/children`
          : `/drives/${driveId}/items/${itemId}/children`;
      const params: Record<string, string> = { $top: String(input.top || 50) };
      const result = await callGraph(graphClient, 'GET', endpoint, params);

      // Check for Loop files in the listing and mark them
      try {
        const parsedResult = JSON.parse(result);
        if (parsedResult.value && Array.isArray(parsedResult.value)) {
          let loopFileCount = 0;
          for (const item of parsedResult.value) {
            const detection = detectLoopFile(item);
            if (detection.isLoopFile) {
              item.isLoopFile = true;
              item.loopDetection = {
                method: detection.detectionMethod,
                confidence: detection.confidence,
              };
              loopFileCount++;
            }
          }
          if (loopFileCount > 0) {
            thinking.push(`📋 Found ${loopFileCount} Loop file(s) in listing`);
          }
          return addThinkingToResponse(JSON.stringify(parsedResult, null, 2), thinking);
        }
      } catch {
        // If parsing fails, just return the original result
      }

      return formatAndReturnToolResponse(result, thinking);
    }

    case 'get': {
      if (!input.itemId) throw new Error('itemId is required for get');
      const driveId = input.driveId || 'me';
      thinking.push(`Getting file metadata: ${input.itemId}`);
      const endpoint =
        driveId === 'me'
          ? `/me/drive/items/${input.itemId}`
          : `/drives/${driveId}/items/${input.itemId}`;
      const result = await callGraph(graphClient, 'GET', endpoint);

      // Check if this is a Loop file and add detection info
      try {
        const parsedResult = JSON.parse(result);
        const loopDetection = detectLoopFile(parsedResult);
        if (loopDetection.isLoopFile) {
          thinking.push(
            `📋 Loop file detected (${loopDetection.detectionMethod}, ${loopDetection.confidence} confidence)`
          );
          parsedResult.isLoopFile = true;
          parsedResult.loopDetection = loopDetection;
          return addThinkingToResponse(JSON.stringify(parsedResult, null, 2), thinking);
        }
      } catch {
        // If parsing fails, just return the original result
      }

      return formatAndReturnToolResponse(result, thinking);
    }

    case 'download': {
      if (!input.itemId) throw new Error('itemId is required for download');
      const driveId = input.driveId || 'me';
      thinking.push(`Downloading file: ${input.itemId}`);

      // First, get file metadata to check if it's a Loop file
      const metadataEndpoint =
        driveId === 'me'
          ? `/me/drive/items/${input.itemId}`
          : `/drives/${driveId}/items/${input.itemId}`;

      let isLoopDetected = false;
      let loopDetectionResult = null;
      try {
        const metadataResult = await callGraph(graphClient, 'GET', metadataEndpoint);
        const metadata = JSON.parse(metadataResult);
        const detection = detectLoopFile(metadata);
        if (detection.isLoopFile) {
          isLoopDetected = true;
          loopDetectionResult = detection;
          thinking.push(
            `📋 Loop file detected (${detection.detectionMethod}, ${detection.confidence} confidence)`
          );
        }
      } catch {
        // Continue with download even if metadata fails
      }

      const contentEndpoint =
        driveId === 'me'
          ? `/me/drive/items/${input.itemId}/content`
          : `/drives/${driveId}/items/${input.itemId}/content`;
      const result = await callGraph(graphClient, 'GET', contentEndpoint);

      // If this is a Loop file, try to parse and extract content
      if (isLoopDetected && typeof result === 'string') {
        thinking.push('📖 Parsing Loop file content');
        const parsedContent = parseLoopContent(result);

        if (parsedContent.success) {
          const response = {
            isLoopFile: true,
            loopDetection: loopDetectionResult,
            contentType: parsedContent.contentType,
            textContent: parsedContent.textContent || null,
            metadata: parsedContent.metadata || null,
            rawContentLength: parsedContent.rawContent.length,
            rawContent:
              parsedContent.rawContent.length <= 10000
                ? parsedContent.rawContent
                : parsedContent.rawContent.substring(0, 10000) + '... (truncated)',
          };

          if (parsedContent.textContent) {
            thinking.push(
              `✅ Extracted ${parsedContent.textContent.length} characters of text content`
            );
          }

          return formatAndReturnToolResponse(response, thinking);
        }
      }

      return formatAndReturnToolResponse(result, thinking);
    }

    case 'root': {
      const driveId = input.driveId || 'me';
      thinking.push('Getting drive root');
      const endpoint = driveId === 'me' ? '/me/drive/root' : `/drives/${driveId}/root`;
      const result = await callGraph(graphClient, 'GET', endpoint);
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'search': {
      if (!input.search) throw new Error('search query is required');

      // Automatic query optimization for file search
      const startTime = Date.now();
      const currentUserId = getUserId();
      const userIdHash = currentUserId ? getQueryStore().hashUserId(currentUserId) : undefined;
      const optimized = optimizeQueryForSearch(input.search, {
        tool: 'files',
        entityTypes: ['driveItem'],
        userIdHash,
      });
      thinking.push(`🔍 Searching files for: "${input.search}"`);
      addOptimizationThinking(thinking, optimized);
      if (optimized.nlpAnalysis.intent) {
        thinking.push(`📊 Detected intent: ${optimized.nlpAnalysis.intent}`);
      }

      const result = await callGraph(
        graphClient,
        'GET',
        `/me/drive/root/search(q='${encodeURIComponent(optimized.optimizedQuery)}')`
      );
      const parsedResult = JSON.parse(result);
      const executionTime = Date.now() - startTime;

      // Record optimization result for learning
      const fileSearchSuccess = parsedResult?.value?.length > 0;
      if (userIdHash) {
        recordQueryOptimizationResult(
          input.search,
          optimized.optimizedQuery,
          fileSearchSuccess,
          userIdHash,
          'files'
        );
      }

      // Extract pagination info
      const pagination = extractPaginationInfo(parsedResult, input.skip, input.top);

      // Format response with metadata
      const responseWithMetadata = formatStandardResponse(parsedResult, {
        executionTime,
        sources: ['files'],
        cacheHit: false,
        pagination,
        nlpAnalysis: optimized.nlpAnalysis,
        suggestions: [
          '💡 Use "files" tool with action "get" to view file details',
          '💡 Use "files" tool with action "download" to download file content',
          '💡 Use "files" tool with action "list" to browse folders',
        ],
      });

      return formatAndReturnToolResponse(responseWithMetadata, thinking);
    }

    default:
      throw new Error(`Unknown files action: ${input.action}`);
  }
}

// ============================================================================
// 5. TASKS SUPER-TOOL
// ============================================================================
const tasksActions = z.enum([
  // Read operations
  'todo-lists', // List To-Do task lists
  'todo-tasks', // List tasks in a To-Do list
  'todo-get', // Get specific To-Do task
  'planner-tasks', // List Planner tasks assigned to me
  'planner-plans', // Get Planner plan details
  'plan-tasks', // List tasks in a Planner plan
  // Write operations (blocked in read-only mode)
  'create-todo', // Create To-Do task
  'update-todo', // Update To-Do task
  'delete-todo', // Delete To-Do task
]);

const tasksSchema = z.object({
  action: tasksActions.describe(
    'Tasks operation: todo-lists, todo-tasks, planner-tasks (read) | create-todo, update-todo, delete-todo (write)'
  ),
  // Identifiers
  taskListId: z.string().optional().describe('To-Do task list ID'),
  taskId: z.string().optional().describe('Task ID'),
  planId: z.string().optional().describe('Planner plan ID'),
  // For create/update todo
  title: z
    .string()
    .optional()
    .default('Untitled task')
    .describe('Task title (for create-todo, update-todo)'),
  dueDateTime: z
    .string()
    .optional()
    .describe('Due date/time ISO format (for create-todo, update-todo)'),
  isCompleted: z.boolean().optional().describe('Mark as completed (for update-todo)'),
  // Filters
  ...filterSchema,
  ...paginationSchema,
});

type TasksInput = z.infer<typeof tasksSchema>;

async function handleTasks(
  input: TasksInput,
  graphClient: GraphClient,
  readOnly: boolean
): Promise<string> {
  const thinking: string[] = [];

  // Check write operations against readOnly mode
  if (['create-todo', 'update-todo', 'delete-todo'].includes(input.action)) {
    checkReadOnly(readOnly, input.action);
  }

  switch (input.action) {
    case 'todo-lists': {
      thinking.push('Listing To-Do task lists');
      const result = await callGraph(graphClient, 'GET', '/me/todo/lists');
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'todo-tasks': {
      if (!input.taskListId) throw new Error('taskListId is required');
      thinking.push(`Listing tasks in list: ${input.taskListId}`);
      const params: Record<string, string> = { $top: String(input.top || 50) };
      if (input.filter) params.$filter = input.filter;
      const result = await callGraph(
        graphClient,
        'GET',
        `/me/todo/lists/${input.taskListId}/tasks`,
        params
      );
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'todo-get': {
      if (!input.taskListId || !input.taskId) {
        throw new Error('taskListId and taskId are required');
      }
      thinking.push(`Getting task: ${input.taskId}`);
      const result = await callGraph(
        graphClient,
        'GET',
        `/me/todo/lists/${input.taskListId}/tasks/${input.taskId}`
      );
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'planner-tasks': {
      thinking.push('Listing Planner tasks assigned to me');
      const result = await callGraph(graphClient, 'GET', '/me/planner/tasks');
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'planner-plans': {
      if (!input.planId) throw new Error('planId is required');
      thinking.push(`Getting Planner plan: ${input.planId}`);
      const result = await callGraph(graphClient, 'GET', `/planner/plans/${input.planId}`);
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'plan-tasks': {
      if (!input.planId) throw new Error('planId is required');
      thinking.push(`Listing tasks in plan: ${input.planId}`);
      const result = await callGraph(graphClient, 'GET', `/planner/plans/${input.planId}/tasks`);
      return formatAndReturnToolResponse(result, thinking);
    }

    // Write operations (blocked in read-only mode - check happens at function start)
    case 'create-todo': {
      if (!input.taskListId) throw new Error('taskListId is required for create-todo');
      const title = input.title ?? 'Untitled task';
      thinking.push(`Creating To-Do task: ${title}`);
      const task: Record<string, unknown> = { title };
      if (input.dueDateTime) {
        task.dueDateTime = { dateTime: input.dueDateTime, timeZone: 'UTC' };
      }
      const result = await callGraph(
        graphClient,
        'POST',
        `/me/todo/lists/${input.taskListId}/tasks`,
        undefined,
        task
      );
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'update-todo': {
      if (!input.taskListId) throw new Error('taskListId is required for update-todo');
      if (!input.taskId) throw new Error('taskId is required for update-todo');
      thinking.push(`Updating To-Do task: ${input.taskId}`);
      const updates: Record<string, unknown> = {};
      if (input.title) updates.title = input.title;
      if (input.dueDateTime) updates.dueDateTime = { dateTime: input.dueDateTime, timeZone: 'UTC' };
      if (input.isCompleted !== undefined) {
        updates.status = input.isCompleted ? 'completed' : 'notStarted';
      }
      const result = await callGraph(
        graphClient,
        'PATCH',
        `/me/todo/lists/${input.taskListId}/tasks/${input.taskId}`,
        undefined,
        updates
      );
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'delete-todo': {
      if (!input.taskListId) throw new Error('taskListId is required for delete-todo');
      if (!input.taskId) throw new Error('taskId is required for delete-todo');
      thinking.push(`Deleting To-Do task: ${input.taskId}`);
      const result = await callGraph(
        graphClient,
        'DELETE',
        `/me/todo/lists/${input.taskListId}/tasks/${input.taskId}`
      );
      return addThinkingToResponse(
        result || JSON.stringify({ success: true, message: 'Task deleted' }),
        thinking
      );
    }

    default:
      throw new Error(`Unknown tasks action: ${input.action}`);
  }
}

// ============================================================================
// 6. CONTACTS SUPER-TOOL
// ============================================================================
const contactsActions = z.enum([
  'list', // List contacts
  'get', // Get specific contact
  'users', // List organization users
  'current-user', // Get current user info
  'search', // Search contacts/users
]);

const contactsSchema = z
  .object({
    action: contactsActions.describe('The contacts operation to perform'),
    // Identifiers
    contactId: z.string().optional().describe('Contact ID'),
    userId: z.string().optional().describe('User ID'),
    // Filters
    ...filterSchema,
    ...paginationSchema,
  })
  .refine(
    (data) => {
      // When action is 'search', search query is required
      if (data.action === 'search' && !data.search) {
        return false;
      }
      return true;
    },
    {
      message: 'search query is required when action is "search"',
      path: ['search'],
    }
  );

type ContactsInput = z.infer<typeof contactsSchema>;

async function handleContacts(
  input: ContactsInput,
  graphClient: GraphClient,
  _readOnly: boolean
): Promise<string> {
  const thinking: string[] = [];
  // Contacts operations are read-only in this version

  switch (input.action) {
    case 'list': {
      thinking.push('Listing contacts');
      const params: Record<string, string> = { $top: String(input.top || 50) };
      if (input.filter) params.$filter = input.filter;
      if (input.search) params.$search = `"${input.search}"`;
      const result = await callGraph(graphClient, 'GET', '/me/contacts', params);
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'get': {
      if (!input.contactId) throw new Error('contactId is required');
      thinking.push(`Getting contact: ${input.contactId}`);
      const result = await callGraph(graphClient, 'GET', `/me/contacts/${input.contactId}`);
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'users': {
      thinking.push('Listing organization users');
      const params: Record<string, string> = { $top: String(input.top || 50) };
      if (input.filter) params.$filter = input.filter;
      if (input.search) {
        // Microsoft Graph API requires property:value format for $search on /users endpoint
        params.$search = formatSearchQuery(input.search, 'displayName');
      }
      const result = await callGraph(graphClient, 'GET', '/users', params, undefined, {
        ConsistencyLevel: 'eventual',
      });
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'current-user': {
      thinking.push('Getting current user info');
      const result = await callGraph(graphClient, 'GET', '/me');
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'search': {
      if (!input.search) throw new Error('search query is required');

      // Automatic query optimization for contact/user search
      const startTime = Date.now();
      const currentUserId = getUserId();
      const userIdHash = currentUserId ? getQueryStore().hashUserId(currentUserId) : undefined;
      const optimized = optimizeQueryForSearch(input.search, {
        tool: 'contacts',
        entityTypes: ['person'],
        userIdHash,
      });
      thinking.push(`🔍 Searching contacts/users for: "${input.search}"`);
      addOptimizationThinking(thinking, optimized);
      if (optimized.nlpAnalysis.intent) {
        thinking.push(`📊 Detected intent: ${optimized.nlpAnalysis.intent}`);
      }

      // Microsoft Graph API requires property:value format for $search on /users endpoint
      const params: Record<string, string> = {
        $search: formatSearchQuery(optimized.optimizedQuery, 'displayName'),
        $top: String(input.top || 25),
      };
      const result = await callGraph(graphClient, 'GET', '/users', params, undefined, {
        ConsistencyLevel: 'eventual',
      });
      const parsedResult = JSON.parse(result);
      const executionTime = Date.now() - startTime;

      // Record optimization result for learning
      const contactSearchSuccess = parsedResult?.value?.length > 0;
      if (userIdHash) {
        recordQueryOptimizationResult(
          input.search,
          optimized.optimizedQuery,
          contactSearchSuccess,
          userIdHash,
          'contacts'
        );
      }

      // Extract pagination info
      const pagination = extractPaginationInfo(parsedResult, input.skip, input.top);

      // Format response with metadata
      const responseWithMetadata = formatStandardResponse(parsedResult, {
        executionTime,
        sources: ['contacts', 'users'],
        cacheHit: false,
        pagination,
        nlpAnalysis: optimized.nlpAnalysis,
        suggestions: [
          '💡 Use "contacts" tool with action "get" to view contact details',
          '💡 Use "assistant" tool with action "discover-person" for comprehensive person info',
        ],
      });

      return formatAndReturnToolResponse(responseWithMetadata, thinking);
    }

    default:
      throw new Error(`Unknown contacts action: ${input.action}`);
  }
}

// ============================================================================
// 7. MEETINGS SUPER-TOOL
// ============================================================================
const meetingsActions = z.enum([
  'list', // List online meetings
  'get', // Get meeting details
  'recordings', // List/get recordings
  'transcripts', // List/get transcripts
  'transcript-content', // Get transcript content
]);

const meetingsSchema = z.object({
  action: meetingsActions.describe('The meetings operation to perform'),
  // Identifiers
  meetingId: z.string().optional().describe('Online meeting ID'),
  recordingId: z.string().optional().describe('Recording ID'),
  transcriptId: z.string().optional().describe('Transcript ID'),
  // Filters
  ...filterSchema,
  ...paginationSchema,
});

type MeetingsInput = z.infer<typeof meetingsSchema>;

async function handleMeetings(
  input: MeetingsInput,
  graphClient: GraphClient,
  _readOnly: boolean
): Promise<string> {
  const thinking: string[] = [];
  // Meetings operations are read-only in this version

  switch (input.action) {
    case 'list': {
      thinking.push('Listing online meetings');
      const params: Record<string, string> = { $top: String(input.top || 25) };
      if (input.filter) params.$filter = input.filter;
      const result = await callGraph(graphClient, 'GET', '/me/onlineMeetings', params);
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'get': {
      if (!input.meetingId) throw new Error('meetingId is required');
      thinking.push(`Getting meeting: ${input.meetingId}`);
      const result = await callGraph(graphClient, 'GET', `/me/onlineMeetings/${input.meetingId}`);
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'recordings': {
      if (!input.meetingId) throw new Error('meetingId is required');
      thinking.push(`Getting recordings for meeting: ${input.meetingId}`);
      const endpoint = input.recordingId
        ? `/me/onlineMeetings/${input.meetingId}/recordings/${input.recordingId}`
        : `/me/onlineMeetings/${input.meetingId}/recordings`;
      const result = await callGraph(graphClient, 'GET', endpoint);
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'transcripts': {
      if (!input.meetingId) throw new Error('meetingId is required');
      thinking.push(`Getting transcripts for meeting: ${input.meetingId}`);
      const endpoint = input.transcriptId
        ? `/me/onlineMeetings/${input.meetingId}/transcripts/${input.transcriptId}`
        : `/me/onlineMeetings/${input.meetingId}/transcripts`;
      const result = await callGraph(graphClient, 'GET', endpoint);
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'transcript-content': {
      if (!input.meetingId || !input.transcriptId) {
        throw new Error('meetingId and transcriptId are required');
      }
      thinking.push(`Getting transcript content: ${input.transcriptId}`);
      const result = await callGraph(
        graphClient,
        'GET',
        `/me/onlineMeetings/${input.meetingId}/transcripts/${input.transcriptId}/content`
      );
      return formatAndReturnToolResponse(result, thinking);
    }

    default:
      throw new Error(`Unknown meetings action: ${input.action}`);
  }
}

// ============================================================================
// 8. SHAREPOINT SUPER-TOOL
// ============================================================================
/** Returns true if the value looks like a Graph API site ID (GUID or hostname,siteGuid,webGuid). */
function looksLikeSharePointSiteId(value: string): boolean {
  const trimmed = value.trim();
  if (!trimmed) return false;
  const guidPart = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  if (guidPart.test(trimmed)) return true;
  const composite = /^[^,\s]+,([0-9a-f-]+,)?[0-9a-f-]+$/i;
  return composite.test(trimmed) && !trimmed.includes(' ');
}

/**
 * Resolves site identifier to a Graph API site ID.
 * If the value looks like a site ID (GUID or hostname,guid,guid), returns it as-is.
 * Otherwise searches sites by name and returns the first match's ID.
 */
async function resolveSharePointSiteId(
  graphClient: GraphClient,
  siteIdOrName: string,
  thinking: string[]
): Promise<string> {
  const trimmed = siteIdOrName.trim();
  if (looksLikeSharePointSiteId(trimmed)) return trimmed;
  thinking.push(`Resolving site by name: "${trimmed}"`);
  const result = await callGraph(graphClient, 'GET', '/sites', {
    search: trimmed,
    $top: '5',
  });
  const parsed = JSON.parse(result) as { value?: Array<{ id?: string; displayName?: string }> };
  const sites = parsed?.value;
  if (!sites?.length) {
    throw new Error(
      `No SharePoint site found for "${trimmed}". Use action "search-sites" with search parameter to list available sites.`
    );
  }
  const first = sites[0];
  if (!first?.id) {
    throw new Error(`SharePoint search returned no valid site ID for "${trimmed}".`);
  }
  thinking.push(`Using site: ${first.displayName ?? first.id}`);
  return first.id;
}

const sharepointActions = z.enum([
  'search-sites', // Search SharePoint sites
  'get-site', // Get site details
  'site-drives', // List site drives
  'site-lists', // List site lists
  'list-items', // List items in a list
  'site-items', // List items in a site
]);

const sharepointSchema = z.object({
  action: sharepointActions.describe('The SharePoint operation to perform'),
  // Identifiers (siteId can be Graph site ID or site display name for lookup)
  siteId: z
    .string()
    .optional()
    .describe('Site ID (GUID or hostname,id,id) or site display name to resolve'),
  siteld: z.string().optional().describe('Alias for siteId (use siteId when possible)'),
  driveId: z.string().optional().describe('Drive ID'),
  listId: z.string().optional().describe('List ID'),
  itemId: z.string().optional().describe('Item ID'),
  // Filters
  ...filterSchema,
  ...paginationSchema,
});

type SharePointInput = z.infer<typeof sharepointSchema>;

async function handleSharePoint(
  input: SharePointInput,
  graphClient: GraphClient,
  _readOnly: boolean
): Promise<string> {
  const thinking: string[] = [];
  // SharePoint operations are read-only in this version

  switch (input.action) {
    case 'search-sites': {
      thinking.push('Searching SharePoint sites');
      const params: Record<string, string> = { $top: String(input.top || 25) };
      if (input.search) params.search = input.search;
      const result = await callGraph(graphClient, 'GET', '/sites', params);
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'get-site': {
      const siteIdOrName = input.siteId ?? input.siteld;
      if (!siteIdOrName) throw new Error('siteId is required');
      const siteId = await resolveSharePointSiteId(graphClient, siteIdOrName, thinking);
      thinking.push(`Getting site: ${siteId}`);
      const result = await callGraph(graphClient, 'GET', `/sites/${siteId}`);
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'site-drives': {
      const siteIdOrName = input.siteId ?? input.siteld;
      if (!siteIdOrName) throw new Error('siteId is required');
      const siteId = await resolveSharePointSiteId(graphClient, siteIdOrName, thinking);
      thinking.push(`Listing drives for site: ${siteId}`);
      const result = await callGraph(graphClient, 'GET', `/sites/${siteId}/drives`);
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'site-lists': {
      const siteIdOrName = input.siteId ?? input.siteld;
      if (!siteIdOrName) throw new Error('siteId is required');
      const siteId = await resolveSharePointSiteId(graphClient, siteIdOrName, thinking);
      thinking.push(`Listing lists for site: ${siteId}`);
      const result = await callGraph(graphClient, 'GET', `/sites/${siteId}/lists`);
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'list-items': {
      const siteIdOrName = input.siteId ?? input.siteld;
      if (!siteIdOrName || !input.listId) {
        throw new Error('siteId and listId are required');
      }
      const siteId = await resolveSharePointSiteId(graphClient, siteIdOrName, thinking);
      thinking.push(`Listing items in list: ${input.listId}`);
      const params: Record<string, string> = { $top: String(input.top || 50) };
      if (input.filter) params.$filter = input.filter;
      const result = await callGraph(
        graphClient,
        'GET',
        `/sites/${siteId}/lists/${input.listId}/items`,
        params
      );
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'site-items': {
      const siteIdOrName = input.siteId ?? input.siteld;
      if (!siteIdOrName) throw new Error('siteId is required');
      const siteId = await resolveSharePointSiteId(graphClient, siteIdOrName, thinking);
      thinking.push(`Listing items in site: ${siteId}`);
      const params: Record<string, string> = { $top: String(input.top || 50) };
      const result = await callGraph(graphClient, 'GET', `/sites/${siteId}/items`, params);
      return formatAndReturnToolResponse(result, thinking);
    }

    default:
      throw new Error(`Unknown sharepoint action: ${input.action}`);
  }
}

// ============================================================================
// 9. NOTES SUPER-TOOL (OneNote)
// ============================================================================
const notesActions = z.enum([
  'notebooks', // List notebooks
  'sections', // List sections in notebook
  'pages', // List pages in section (requires sectionId) OR search all pages (with search parameter)
  'page-content', // Get page content
  'search-pages', // Search all pages by title
]);

const notesSchema = z.object({
  action: notesActions.describe('The OneNote operation to perform'),
  // Identifiers
  notebookId: z.string().optional().describe('Notebook ID'),
  sectionId: z
    .string()
    .optional()
    .describe('Section ID (required for pages action without search)'),
  pageId: z.string().optional().describe('Page ID'),
  // Filters
  ...filterSchema,
  ...paginationSchema,
});

type NotesInput = z.infer<typeof notesSchema>;

async function handleNotes(
  input: NotesInput,
  graphClient: GraphClient,
  _readOnly: boolean
): Promise<string> {
  const thinking: string[] = [];
  // Notes operations are read-only in this version

  switch (input.action) {
    case 'notebooks': {
      thinking.push('Listing OneNote notebooks');
      const result = await callGraph(graphClient, 'GET', '/me/onenote/notebooks');
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'sections': {
      if (!input.notebookId) throw new Error('notebookId is required');
      thinking.push(`Listing sections in notebook: ${input.notebookId}`);
      const result = await callGraph(
        graphClient,
        'GET',
        `/me/onenote/notebooks/${input.notebookId}/sections`
      );
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'pages': {
      const params: Record<string, string> = { $top: String(input.top || 50) };

      // If search parameter is provided, search across all pages
      if (input.search) {
        thinking.push(`Searching all OneNote pages for: "${input.search}"`);
        // Use $filter to search by title contains
        params.$filter = `contains(title,'${input.search.replace(/'/g, "''")}')`;
        params.$orderby = 'lastModifiedDateTime desc';
        const result = await callGraph(graphClient, 'GET', '/me/onenote/pages', params);
        return formatAndReturnToolResponse(result, thinking);
      }

      // Otherwise, require sectionId to list pages in a specific section
      if (!input.sectionId) {
        throw new Error(
          'Either sectionId or search parameter is required. Use sectionId to list pages in a section, or search to find pages by title.'
        );
      }
      thinking.push(`Listing pages in section: ${input.sectionId}`);
      const result = await callGraph(
        graphClient,
        'GET',
        `/me/onenote/sections/${input.sectionId}/pages`,
        params
      );
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'search-pages': {
      if (!input.search) throw new Error('search parameter is required for search-pages action');
      thinking.push(`Searching all OneNote pages for: "${input.search}"`);
      const params: Record<string, string> = {
        $top: String(input.top || 50),
        $filter: `contains(title,'${input.search.replace(/'/g, "''")}')`,
        $orderby: 'lastModifiedDateTime desc',
      };
      const result = await callGraph(graphClient, 'GET', '/me/onenote/pages', params);
      return formatAndReturnToolResponse(result, thinking);
    }

    case 'page-content': {
      if (!input.pageId) throw new Error('pageId is required');
      thinking.push(`Getting page content: ${input.pageId}`);
      const result = await callGraph(
        graphClient,
        'GET',
        `/me/onenote/pages/${input.pageId}/content`
      );
      return formatAndReturnToolResponse(result, thinking);
    }

    default:
      throw new Error(`Unknown notes action: ${input.action}`);
  }
}

// ============================================================================
// 10. SEARCH SUPER-TOOL (Microsoft 365 Unified Search)
// ============================================================================
/**
 * The Search Super-Tool uses Microsoft Graph Search API to search across
 * all Microsoft 365 content. This is the RECOMMENDED FIRST TOOL to use
 * when exploring data, as it helps identify which specific tools to use next.
 *
 * EntityTypes:
 * - message: Emails
 * - event: Calendar events
 * - driveItem: OneDrive/SharePoint files
 * - site: SharePoint sites
 * - list: SharePoint lists
 * - listItem: SharePoint list items
 * - chatMessage: Teams chat messages
 * - person: People in the organization
 */
const searchEntityTypes = [
  'message',
  'event',
  'driveItem',
  'site',
  'list',
  'listItem',
  'chatMessage',
  'person',
  'acronym',
  'bookmark',
  'qna',
  'externalItem',
] as const;

/**
 * Get available entity types based on token permissions
 * Uses comprehensive default set and filters based on API responses
 */
async function getAvailableEntityTypes(
  graphClient: GraphClient,
  defaultTypes: string[]
): Promise<string[]> {
  // Start with comprehensive set including all possible entity types
  const allPossibleTypes = [
    'message',
    'event',
    'driveItem',
    'site',
    'list',
    'listItem',
    'chatMessage',
    'person',
    'acronym',
    'bookmark',
    'qna',
    'externalItem',
  ];

  // Use default types if provided, otherwise use all possible
  const typesToTry = defaultTypes.length > 0 ? defaultTypes : allPossibleTypes;

  // For now, return the types to try
  // In future, we could test each type with a minimal search query
  // and filter out those that return permission errors
  return typesToTry;
}

const searchSchema = z.object({
  query: z.string().describe('The search query - natural language or keywords'),
  entityTypes: z
    .array(z.enum(searchEntityTypes))
    .optional()
    .describe(
      'Types of entities to search: message (emails), event (calendar), driveItem (files), site, list, listItem, chatMessage, person. Default: all types.'
    ),
  from: z.number().optional().describe('Starting index for pagination (default: 0)'),
  size: z.number().optional().describe('Number of results to return (default: 25, max: 500)'),
  // Advanced options
  fields: z.array(z.string()).optional().describe('Specific fields to return in results'),
  sortBy: z.string().optional().describe('Field to sort results by'),
  trimDuplicates: z.boolean().optional().describe('Remove duplicate results (default: true)'),
});

type SearchInput = z.infer<typeof searchSchema>;

// ============================================================================
// PRODUCT-BASED SEARCH SCHEMA
// ============================================================================

const productSearchActions = z.enum(['search']);

const productSearchSchema = z.object({
  action: productSearchActions,
  query: z.string().describe('Search query string'),
  maxResults: z.number().optional().describe('Maximum initial search results (default: 50)'),
  topPerProduct: z
    .number()
    .optional()
    .describe('Top results per product to summarize (default: 5)'),
});

type ProductSearchInput = z.infer<typeof productSearchSchema>;

// Product mapping configuration
interface ProductMapping {
  product: string;
  entityTypes: string[];
  apiEndpoint?: string;
  requiresDirectApi: boolean;
}

const PRODUCT_MAPPINGS: ProductMapping[] = [
  {
    product: 'Outlook',
    entityTypes: ['message'],
    apiEndpoint: '/me/messages',
    requiresDirectApi: true,
  },
  {
    product: 'Calendar',
    entityTypes: ['event'],
    apiEndpoint: '/me/calendarView',
    requiresDirectApi: true,
  },
  {
    product: 'OneDrive',
    entityTypes: ['driveItem'],
    apiEndpoint: '/me/drive/root/search',
    requiresDirectApi: true,
  },
  {
    product: 'SharePoint',
    entityTypes: ['site', 'list', 'listItem'],
    apiEndpoint: '/sites',
    requiresDirectApi: true,
  },
  {
    product: 'Teams',
    entityTypes: ['chatMessage'],
    apiEndpoint: '/me/chats',
    requiresDirectApi: true,
  },
  {
    product: 'OneNote',
    entityTypes: [],
    apiEndpoint: '/me/onenote/pages',
    requiresDirectApi: true,
  },
  { product: 'Users', entityTypes: ['person'], apiEndpoint: '/users', requiresDirectApi: true },
  { product: 'Groups', entityTypes: [], apiEndpoint: '/groups', requiresDirectApi: true },
  {
    product: 'Planner',
    entityTypes: [],
    apiEndpoint: '/me/planner/plans',
    requiresDirectApi: true,
  },
  { product: 'ToDo', entityTypes: [], apiEndpoint: '/me/todo/lists', requiresDirectApi: true },
];

// Product search result interfaces
interface ProductSearchResult {
  product: string;
  resultCount: number;
  topResults: Array<{
    title: string;
    summary: string;
    relevance: number;
    webUrl?: string;
    metadata: Record<string, unknown>;
  }>;
}

interface ProductSearchResponse {
  query: string;
  initialSearchResults: {
    totalHits: number;
    productsDetected: string[];
  };
  productResults: ProductSearchResult[];
  thinking: string[];
}

// Common German/English words that should NOT be detected as person names
const NON_PERSON_WORDS = new Set([
  // German temporal words
  'letzte',
  'letzten',
  'letzter',
  'nächste',
  'nächsten',
  'nächster',
  'heute',
  'morgen',
  'gestern',
  'woche',
  'monat',
  'jahr',
  'montag',
  'dienstag',
  'mittwoch',
  'donnerstag',
  'freitag',
  'samstag',
  'sonntag',
  // German common words
  'alle',
  'meine',
  'zeige',
  'finde',
  'suche',
  'liste',
  'wichtige',
  'dringende',
  'neue',
  'aktuelle',
  // English temporal words
  'last',
  'next',
  'today',
  'tomorrow',
  'yesterday',
  'week',
  'month',
  'year',
  'monday',
  'tuesday',
  'wednesday',
  'thursday',
  'friday',
  'saturday',
  'sunday',
  // English common words
  'all',
  'my',
  'show',
  'find',
  'search',
  'list',
  'important',
  'urgent',
  'new',
  'current',
  'recent',
]);

// ============================================================================
// INTELLIGENT QUERY DECOMPOSITION
// ============================================================================

/**
 * Sub-query with context for multi-query execution
 */
interface SubQueryWithContext {
  query: string;
  entityTypes: string[];
  priority: number;
  reason: string;
  expectedResultType: 'primary' | 'complementary' | 'fallback';
}

/**
 * Result from executing multiple sub-queries
 */
interface MultiQueryResult {
  results: Array<{
    subQuery: SubQueryWithContext;
    searchResult: SearchHit[] | null;
    error: string | null;
    durationMs: number;
  }>;
  totalDurationMs: number;
  successCount: number;
}

/**
 * Search hit from Graph API
 */
interface SearchHit {
  hitId: string;
  rank: number;
  summary?: string;
  resource: Record<string, unknown>;
}

/**
 * Merged search result
 */
interface MergedSearchResult {
  results: Record<string, MergedHit[]>;
  totalHits: number;
  uniqueHits: number;
  multiMatchHits: number;
  queryBreakdown: Array<{
    query: string;
    hitCount: number;
    contributedUniqueHits: number;
    filteredCount?: number;
  }>;
}

/**
 * Merged hit with combined ranking
 */
interface MergedHit {
  id: string;
  resource: unknown;
  combinedRank: number;
  matchedQueries: string[];
  matchCount: number;
  primaryMatch: boolean;
  relevanceScore?: number;
}

/**
 * Relevance score for a search hit
 */
interface RelevanceScore {
  score: number;
  confidence: number;
  matchedFields: string[];
  queryTermsFound: number;
  isRelevant: boolean;
}

/**
 * Check if a query should be decomposed into sub-queries
 * @param decomposed - Decomposed query from NLP
 * @param originalQuery - Original query string
 * @returns true if query should be decomposed
 */
function shouldDecomposeQuery(decomposed: DecomposedQuery, originalQuery: string): boolean {
  const wordCount = originalQuery.trim().split(/\s+/).length;
  const entityCount = decomposed.entities.length;
  const uniqueEntityTypes = new Set(decomposed.entities.map((e) => e.type)).size;

  // Don't decompose if:
  // - Query is short and simple (< 3 words, 0-1 entity)
  // - Confidence is very high (> 0.9) - query is clearly understood
  if (wordCount < 3 && entityCount <= 1 && decomposed.confidence > 0.85) {
    return false;
  }

  // Don't decompose very short queries
  if (wordCount < 2) {
    return false;
  }

  // Decompose if:
  return (
    entityCount >= 2 || // Multiple entities
    wordCount >= 5 || // Long query
    decomposed.confidence < 0.7 || // Unclear intent
    decomposed.subQueries.length > 3 || // Many sub-queries generated
    uniqueEntityTypes >= 2 || // Different entity types
    decomposed.compoundParts.length > 1 // Compound word
  );
}

/**
 * Extract main context words from a query
 * (e.g., "RathausGPT" from "Angebot für RathausGPT Empfänger extern")
 */
function extractMainContextWords(query: string): string[] {
  const words = query.split(/\s+/);
  // Filter stop words and short words
  const stopWords = new Set([
    'für',
    'von',
    'mit',
    'an',
    'in',
    'auf',
    'zu',
    'der',
    'die',
    'das',
    'und',
    'oder',
    'für',
    'the',
    'a',
    'an',
    'and',
    'or',
    'with',
    'in',
    'on',
    'at',
    'to',
    'for',
    'of',
  ]);
  return words.filter((w) => w.length > 3 && !stopWords.has(w.toLowerCase())).slice(0, 2);
}

/**
 * Get entity types based on entity type
 */
function getEntityTypesForEntityType(
  entityType: string
): Array<(typeof searchEntityTypes)[number]> {
  const mapping: Record<string, Array<(typeof searchEntityTypes)[number]>> = {
    person: ['message', 'chatMessage', 'event'],
    organization: ['message', 'driveItem', 'site'],
    project: ['driveItem', 'site', 'listItem'],
    file: ['driveItem'],
    event: ['event'],
    email: ['message'],
    task: ['message', 'listItem'],
    product: ['message', 'driveItem'],
    location: ['event', 'site'],
    date: ['event', 'message'],
    time: ['event'],
    food: ['driveItem', 'listItem'],
    unknown: ['driveItem', 'site', 'message'],
  };
  return mapping[entityType] || ['driveItem', 'site', 'message'];
}

/**
 * Select the best sub-queries from decomposed query WITH CONTEXT PRESERVATION
 * @param decomposed - Decomposed query from NLP
 * @param userIdHash - User ID hash for history-based optimization
 * @param maxQueries - Maximum number of sub-queries
 * @returns Array of sub-queries with context
 */
function selectBestSubQueries(
  decomposed: DecomposedQuery,
  userIdHash: string,
  maxQueries = 4
): SubQueryWithContext[] {
  const selected: SubQueryWithContext[] = [];

  // Identify main entity (most important context)
  const mainEntity = decomposed.entity || decomposed.entities[0]?.value || '';
  const mainEntityWords = mainEntity.split(/\s+/).filter((w) => w.length > 3);
  const mainEntityValue = mainEntityWords[0] || ''; // First important word (e.g., "RathausGPT")

  // Get user-specific optimal entity types if available
  let primaryEntityTypes: Array<(typeof searchEntityTypes)[number]> = ['message', 'driveItem'];
  if (decomposed.ms365Context?.searchScopes) {
    primaryEntityTypes = decomposed.ms365Context.searchScopes as Array<
      (typeof searchEntityTypes)[number]
    >;
  }

  // Try to get from user history
  if (userIdHash) {
    try {
      const queryStore = getQueryStore();
      const historyRec = queryStore.getOptimalEntityTypes(decomposed.original, userIdHash);
      if (historyRec && historyRec.confidence > 0.6) {
        primaryEntityTypes = historyRec.entityTypes as Array<(typeof searchEntityTypes)[number]>;
      }
    } catch {
      // Ignore errors
    }
  }

  // 1. Primary Query (always included) - keeps original context
  selected.push({
    query: decomposed.entity || decomposed.original,
    entityTypes: validateEntityTypeCombinations(primaryEntityTypes),
    priority: 1,
    reason: 'Primary entity-focused query',
    expectedResultType: 'primary',
  });

  // 2. Entity-specific queries (max 2) - WITH CONTEXT PRESERVATION
  const entityQueries = decomposed.entities
    .filter(
      (e) => e.value !== decomposed.entity && !mainEntityWords.includes(e.value.split(/\s+/)[0])
    )
    .slice(0, 2)
    .map((entity) => {
      // CRITICAL: Combine entity with main context to not lose context
      const queryWithContext = mainEntityValue
        ? `${mainEntityValue} ${entity.value}`.trim()
        : entity.value;

      return {
        query: queryWithContext,
        entityTypes: validateEntityTypeCombinations(getEntityTypesForEntityType(entity.type)),
        priority: 2,
        reason: `Entity: ${entity.type} + context`,
        expectedResultType: 'complementary' as const,
      };
    });
  selected.push(...entityQueries);

  // 3. Intent-specific query (if space) - WITH CONTEXT
  if (selected.length < maxQueries && decomposed.intent.type !== 'unknown') {
    let intentQuery = '';
    let intentEntityTypes: Array<(typeof searchEntityTypes)[number]> = ['message'];

    switch (decomposed.intent.type) {
      case 'who':
        intentQuery = mainEntityValue ? `${mainEntityValue} Empfänger` : decomposed.original;
        intentEntityTypes = ['message', 'chatMessage'];
        break;
      case 'when':
      case 'last_occurrence':
        intentQuery = mainEntityValue ? `${mainEntityValue} letzte` : decomposed.original;
        intentEntityTypes = ['message', 'event'];
        break;
      case 'where':
        intentQuery = mainEntityValue ? `${mainEntityValue} Ort` : decomposed.original;
        intentEntityTypes = ['event', 'site'];
        break;
      default:
        intentQuery = '';
    }

    if (intentQuery && intentQuery !== selected[0]?.query) {
      selected.push({
        query: intentQuery,
        entityTypes: validateEntityTypeCombinations(intentEntityTypes),
        priority: 3,
        reason: `Intent: ${decomposed.intent.type} + context`,
        expectedResultType: 'complementary',
      });
    }
  }

  // 4. Fallback query (if space) - WITH MAIN CONTEXT
  if (selected.length < maxQueries && mainEntityValue) {
    selected.push({
      query: mainEntityValue,
      entityTypes: validateEntityTypeCombinations(['driveItem', 'site', 'listItem']),
      priority: 4,
      reason: 'Fallback broad search with main context',
      expectedResultType: 'fallback',
    });
  }

  return selected.slice(0, maxQueries);
}

/**
 * Validate relevance of a search hit against the original query
 * CRITICAL: Validates against ORIGINAL query, not just sub-query!
 */
function validateRelevance(
  hit: SearchHit,
  originalQuery: string,
  subQuery: string
): RelevanceScore {
  // Extract important context words from original query (e.g., "RathausGPT")
  const originalWords = originalQuery
    .toLowerCase()
    .split(/\s+/)
    .filter((t) => t.length > 3);
  const subQueryWords = subQuery
    .toLowerCase()
    .split(/\s+/)
    .filter((t) => t.length > 3);

  // Identify main context words (in original but not in sub-query)
  const contextWords = originalWords.filter((w) => !subQueryWords.includes(w));

  // Combine: sub-query terms + important context words
  const allRelevantTerms = [...new Set([...subQueryWords, ...contextWords])].filter(
    (t) => t.length > 2
  );

  const resource = hit.resource || {};

  // Key fields for matching
  const keyFields: Record<string, string> = {
    subject: ((resource.subject as string) || '').toLowerCase(),
    name: ((resource.name as string) || '').toLowerCase(),
    displayName: ((resource.displayName as string) || '').toLowerCase(),
    bodyPreview: ((resource.bodyPreview as string) || '').toLowerCase(),
    summary: (hit.summary || '').toLowerCase(),
    content: ((resource.content as string) || '').toLowerCase(),
  };

  const matchedFields: string[] = [];
  let queryTermsFound = 0;
  let contextTermsFound = 0;
  let totalMatches = 0;

  // Check each relevant term in each field
  for (const term of allRelevantTerms) {
    let termFound = false;
    const isContextWord = contextWords.includes(term);

    for (const [field, content] of Object.entries(keyFields)) {
      if (content.includes(term)) {
        if (!matchedFields.includes(field)) {
          matchedFields.push(field);
        }
        termFound = true;
        totalMatches++;
      }
    }
    if (termFound) {
      if (isContextWord) {
        contextTermsFound++;
      } else {
        queryTermsFound++;
      }
    }
  }

  // CRITICAL: If important context words are missing, result is less relevant
  const contextMatchRatio = contextWords.length > 0 ? contextTermsFound / contextWords.length : 1.0;

  // Calculate relevance score
  const termMatchRatio =
    allRelevantTerms.length > 0
      ? (queryTermsFound + contextTermsFound) / allRelevantTerms.length
      : 0;
  const fieldMatchRatio =
    Object.keys(keyFields).length > 0 ? matchedFields.length / Object.keys(keyFields).length : 0;
  const matchDensity = totalMatches / Math.max(1, allRelevantTerms.length);

  // Combined score
  let score = termMatchRatio * 0.4 + fieldMatchRatio * 0.3 + Math.min(matchDensity, 1) * 0.2;

  // CRITICAL: Penalty if important context words are missing
  if (contextWords.length > 0 && contextMatchRatio < 0.5) {
    score *= 0.5; // 50% penalty
  }

  // Confidence based on match quality
  let confidence = score;

  // Boost for exact matches in important fields
  if (
    keyFields.subject.includes(originalQuery.toLowerCase()) ||
    keyFields.name.includes(originalQuery.toLowerCase())
  ) {
    confidence = Math.min(1.0, confidence + 0.2);
  }

  // Penalty for very low rank (late results are less relevant)
  const rankPenalty = hit.rank > 50 ? 0.3 : hit.rank > 25 ? 0.1 : 0;
  confidence = Math.max(0, confidence - rankPenalty);

  // Threshold:
  // - At least 30% of query terms must be found
  // - If context words are present, at least 50% must be found
  const isRelevant =
    termMatchRatio >= 0.3 &&
    confidence >= 0.3 &&
    (contextWords.length === 0 || contextMatchRatio >= 0.5);

  return {
    score,
    confidence,
    matchedFields,
    queryTermsFound: queryTermsFound + contextTermsFound,
    isRelevant,
  };
}

/**
 * Filter irrelevant results from search hits
 */
function filterIrrelevantResults(
  hits: SearchHit[],
  originalQuery: string,
  minRelevance = 0.3
): SearchHit[] {
  return hits
    .map((hit) => ({
      hit,
      relevance: validateRelevance(hit, originalQuery, ''),
    }))
    .filter(({ relevance }) => relevance.isRelevant && relevance.confidence >= minRelevance)
    .sort((a, b) => b.relevance.confidence - a.relevance.confidence)
    .map(({ hit }) => hit);
}

/**
 * Merge search results from multiple sub-queries
 * Includes deduplication, multi-match boost, and relevance validation
 */
function mergeSearchResults(
  multiQueryResult: MultiQueryResult,
  originalQuery: string
): MergedSearchResult {
  const hitMap = new Map<string, MergedHit>();
  const queryBreakdown: MergedSearchResult['queryBreakdown'] = [];
  const minRelevance = parseFloat(process.env.MS365_MCP_MIN_RELEVANCE || '0.3');
  const mainContextWords = extractMainContextWords(originalQuery);

  for (const { subQuery, searchResult } of multiQueryResult.results) {
    if (!searchResult) continue;

    // CRITICAL: Filter irrelevant results BEFORE merging
    const relevantHits = filterIrrelevantResults(searchResult, originalQuery, minRelevance);

    let hitCount = 0;
    let contributedUnique = 0;
    let filteredCount = searchResult.length - relevantHits.length;

    for (const hit of relevantHits) {
      const id = hit.resource?.id as string;
      if (!id) {
        filteredCount++;
        continue;
      }

      // Additional validation: Check relevance against ORIGINAL query
      const relevance = validateRelevance(hit, originalQuery, subQuery.query);
      if (!relevance.isRelevant) {
        filteredCount++;
        continue;
      }

      // Additional context check: Only filter if validateRelevance passed but result
      // clearly lacks important context AND has low confidence
      // This is a lenient check since validateRelevance already validated relevance
      if (mainContextWords.length > 0 && relevance.confidence < 0.35) {
        const resourceText = JSON.stringify(hit.resource || {}).toLowerCase();
        const hasContext = mainContextWords.some((word) =>
          resourceText.includes(word.toLowerCase())
        );
        // Only filter if NO context words found AND confidence is very low
        // This prevents filtering results that validateRelevance deemed relevant
        if (!hasContext) {
          filteredCount++;
          continue;
        }
      }

      // Hit passed all validations
      hitCount++;

      if (hitMap.has(id)) {
        // Multi-match: Boost existing result
        const existing = hitMap.get(id)!;
        existing.matchCount++;
        existing.matchedQueries.push(subQuery.query);

        // Rank combination: Weighted average with boost for multi-match
        const priorityWeight = 5 - subQuery.priority;
        existing.combinedRank = (existing.combinedRank + hit.rank * priorityWeight) / 2;

        // Combine relevance score (highest score)
        existing.relevanceScore = Math.max(existing.relevanceScore || 0, relevance.confidence);

        // Primary match flag
        if (subQuery.expectedResultType === 'primary') {
          existing.primaryMatch = true;
        }
      } else {
        // New result
        contributedUnique++;
        const priorityWeight = 5 - subQuery.priority;
        hitMap.set(id, {
          id,
          resource: hit.resource,
          combinedRank: hit.rank * priorityWeight,
          matchedQueries: [subQuery.query],
          matchCount: 1,
          primaryMatch: subQuery.expectedResultType === 'primary',
          relevanceScore: relevance.confidence,
        });
      }
    }

    queryBreakdown.push({
      query: subQuery.query,
      hitCount,
      contributedUniqueHits: contributedUnique,
      filteredCount,
    });
  }

  // Apply multi-match boost
  for (const hit of hitMap.values()) {
    if (hit.matchCount > 1) {
      // Boost: 20% per additional match
      hit.combinedRank *= 1 + 0.2 * (hit.matchCount - 1);
    }
    if (hit.primaryMatch) {
      // Primary match boost: 30%
      hit.combinedRank *= 1.3;
    }

    // Relevance boost: Higher relevance = higher rank
    if (hit.relevanceScore && hit.relevanceScore > 0.7) {
      hit.combinedRank *= 1.2;
    }
  }

  // Group by entity type and sort
  const results: Record<string, MergedHit[]> = {};
  const allHits = Array.from(hitMap.values());

  for (const hit of allHits) {
    const entityType =
      ((hit.resource as Record<string, unknown>)?.['@odata.type'] as string) || 'unknown';
    if (!results[entityType]) {
      results[entityType] = [];
    }
    results[entityType].push(hit);
  }

  // Sort by combinedRank (higher = better)
  for (const entityType of Object.keys(results)) {
    results[entityType].sort((a, b) => b.combinedRank - a.combinedRank);
  }

  return {
    results,
    totalHits: allHits.reduce((sum, h) => sum + h.matchCount, 0),
    uniqueHits: allHits.length,
    multiMatchHits: allHits.filter((h) => h.matchCount > 1).length,
    queryBreakdown,
  };
}

/**
 * Execute a single search query against Graph API
 */
async function executeSingleSearch(
  graphClient: GraphClient,
  query: string,
  entityTypes: string[],
  size = 25
): Promise<SearchHit[]> {
  try {
    const searchBody = {
      requests: [
        {
          entityTypes: validateEntityTypeCombinations(entityTypes),
          query: { queryString: query },
          from: 0,
          size,
        },
      ],
    };

    const response = await callGraph(graphClient, 'POST', '/search/query', undefined, searchBody);

    const parsed = JSON.parse(response);
    const hits: SearchHit[] = [];

    if (parsed.value?.[0]?.hitsContainers?.[0]?.hits) {
      for (const hit of parsed.value[0].hitsContainers[0].hits) {
        hits.push({
          hitId: hit.hitId || hit.resource?.id || '',
          rank: hit.rank || 0,
          summary: hit.summary,
          resource: hit.resource || {},
        });
      }
    }

    return hits;
  } catch (error) {
    logger.warn(`Search failed for query "${query}": ${error}`);
    return [];
  }
}

/**
 * Execute multiple sub-queries in parallel
 */
async function executeMultiQuery(
  subQueries: SubQueryWithContext[],
  graphClient: GraphClient,
  timeoutMs = 10000
): Promise<MultiQueryResult> {
  const startTime = Date.now();
  const timeoutPerQuery = Math.floor(timeoutMs / Math.max(1, subQueries.length));

  const promises = subQueries.map(async (subQuery) => {
    const queryStart = Date.now();
    try {
      const result = await Promise.race([
        executeSingleSearch(graphClient, subQuery.query, subQuery.entityTypes),
        new Promise<SearchHit[]>((_, reject) =>
          setTimeout(() => reject(new Error('Query timeout')), timeoutPerQuery)
        ),
      ]);

      return {
        subQuery,
        searchResult: result,
        error: null,
        durationMs: Date.now() - queryStart,
      };
    } catch (error) {
      return {
        subQuery,
        searchResult: null,
        error: error instanceof Error ? error.message : 'Unknown error',
        durationMs: Date.now() - queryStart,
      };
    }
  });

  const results = await Promise.allSettled(promises);

  const processedResults = results.map((r) =>
    r.status === 'fulfilled'
      ? r.value
      : {
          subQuery: subQueries[0],
          searchResult: null,
          error: 'Promise rejected',
          durationMs: 0,
        }
  );

  return {
    results: processedResults,
    totalDurationMs: Date.now() - startTime,
    successCount: processedResults.filter((r) => r.searchResult !== null).length,
  };
}

async function handleSearch(
  input: SearchInput,
  graphClient: GraphClient,
  _readOnly: boolean
): Promise<string> {
  const thinking: string[] = [];
  const searchStartTime = Date.now();

  thinking.push(`🔍 Microsoft 365 Search: "${input.query}"`);

  // =========================================================================
  // UQAS BILINGUAL ANALYSIS - Detect language and expand queries (DE/EN)
  // =========================================================================
  const uqasAnalysis = uqas.analyze(input.query);
  const uqasThinking = uqas.createThinkingSteps(uqasAnalysis);
  thinking.push(...uqasThinking);

  // Use cross-language search variants for better coverage
  const searchVariants = uqasAnalysis.searchQueries;
  const detectedLang = uqasAnalysis.language;

  // =========================================================================
  // NLP ANALYSIS - Use NLP to understand query intent
  // =========================================================================
  const decomposed = nlpEnhancer.decomposeQuery(input.query);
  const nlpIntent = decomposed.intent.type;
  const nlpService = decomposed.ms365Context?.service;
  const nlpEntities = decomposed.entities;

  thinking.push(
    `📊 NLP: Intent=${nlpIntent}, Service=${nlpService || 'general'}, Confidence=${Math.round(decomposed.confidence * 100)}%`
  );

  // =========================================================================
  // CALENDAR QUERY DETECTION - Direct Calendar API for calendar queries
  // =========================================================================
  const isCalendarQuery =
    nlpService === 'calendar' ||
    (input.entityTypes?.length === 1 && input.entityTypes[0] === 'event') ||
    /\b(kalender|termine?|meetings?|besprechung|events?|calendar|appointments?)\b/i.test(
      input.query
    );

  if (isCalendarQuery && !input.entityTypes?.includes('message')) {
    thinking.push('📅 Detected calendar query - using Calendar API directly');

    // Parse temporal info from NLP or query
    const days = decomposed.temporal?.relativeDays || 30;
    const requestedSize = input.size || 15;

    // Determine date range
    const now = new Date();
    let startDate: Date;
    let endDate: Date;

    // Check if looking for past or future events
    const isPastQuery =
      /\b(letzte[rn]?|vergangen|last|past|previous|recent)\b/i.test(input.query) ||
      (decomposed.temporal?.relativeDays && decomposed.temporal.relativeDays < 0);

    if (isPastQuery) {
      // Past events
      startDate = new Date(now);
      startDate.setDate(startDate.getDate() - Math.abs(days));
      startDate = setStartOfDay(startDate);
      endDate = setEndOfDay(now);
      thinking.push(`📅 Looking for past ${Math.abs(days)} days of events`);
    } else {
      // Future events (default)
      startDate = setStartOfDay(now);
      endDate = new Date(now);
      endDate.setDate(endDate.getDate() + days);
      endDate = setEndOfDay(endDate);
      thinking.push(`📅 Looking for next ${days} days of events`);
    }

    try {
      const eventsResult = await callGraph(graphClient, 'GET', '/me/calendarView', {
        startDateTime: startDate.toISOString(),
        endDateTime: endDate.toISOString(),
        $top: String(requestedSize),
        $orderby: isPastQuery ? 'start/dateTime desc' : 'start/dateTime asc',
        $select: 'id,subject,start,end,location,organizer,attendees,isOnlineMeeting,onlineMeeting',
      });

      const events = JSON.parse(eventsResult);
      const eventCount = events.value?.length || 0;

      thinking.push(`✅ Found ${eventCount} calendar events`);

      const toolSuggestions = [
        '💡 Use "calendar" tool with action "list" for more calendar options',
        '💡 Use "calendar" tool with action "get" to get event details by ID',
      ];

      if (eventCount > 0) {
        toolSuggestions.push('💡 Use "assistant" tool with action "my-day" for today\'s summary');
      }

      // Format events with local timezone conversion (no raw UTC)
      const rawEvents = events.value || [];
      const formattedEvents = rawEvents.map((event: Record<string, unknown>) =>
        formatCalendarEvent(event)
      );

      const output = {
        query: input.query,
        language: {
          detected: detectedLang,
          confidence: Math.round(uqasAnalysis.languageConfidence * 100),
        },
        nlpAnalysis: {
          intent: nlpIntent,
          service: nlpService,
          temporal: decomposed.temporal,
          confidence: decomposed.confidence,
        },
        totalHits: eventCount,
        entityTypes: ['event'],
        dateRange: {
          start: startDate.toISOString(),
          end: endDate.toISOString(),
          direction: isPastQuery ? 'past' : 'future',
        },
        results: {
          '#microsoft.graph.event': formattedEvents,
        },
        suggestions: toolSuggestions,
      };

      return addThinkingToResponse(JSON.stringify(output, null, 2), thinking);
    } catch (error) {
      thinking.push(
        `Calendar API error: ${error instanceof Error ? error.message : String(error)}`
      );
      thinking.push('Falling back to Search API...');
      // Fall through to regular search
    }
  }

  // =========================================================================
  // PERSON DETECTION - Improved with NLP and word filtering
  // =========================================================================

  // Check if NLP detected a person entity
  const nlpPersonEntity = nlpEntities.find((e) => e.type === 'person');
  let personName: string | null = nlpPersonEntity?.value || null;

  // Fallback to regex patterns if NLP didn't find a person
  if (!personName) {
    const personPatterns = [
      /\b(mit|von|to|from|with)\s+([A-ZÄÖÜ][a-zäöüß]+(?:\s+[A-ZÄÖÜ][a-zäöüß]+)?)\b/i,
      /\b(chat|nachricht|message|teams)\s+(mit|von|to|from|with)\s+([A-ZÄÖÜ][a-zäöüß]+)/i,
    ];

    for (const pattern of personPatterns) {
      const match = input.query.match(pattern);
      if (match) {
        const candidateName = match[3] || match[2];
        // Verify it's not a common word
        if (candidateName && !NON_PERSON_WORDS.has(candidateName.toLowerCase())) {
          personName = candidateName;
          break;
        }
      }
    }
  }

  // Validate personName - exclude common words
  if (personName && NON_PERSON_WORDS.has(personName.toLowerCase())) {
    personName = null;
  }

  const hasTeamsKeywords = /\b(teams|chat|nachricht|message)\b/i.test(input.query);
  const hasPersonQuery = personName !== null;

  // Intelligent entity type detection based on query and NLP
  let entityTypes: Array<(typeof searchEntityTypes)[number]> = input.entityTypes || [];

  // =========================================================================
  // HISTORY-BASED ENTITY TYPE RECOMMENDATION (USER-SPECIFIC)
  // =========================================================================
  const patternLearningEnabled = process.env.MS365_MCP_PATTERN_LEARNING_ENABLED !== 'false';
  let historyRecommendationUsed = false;

  if (entityTypes.length === 0 && patternLearningEnabled) {
    // Try to get entity types from user history
    try {
      const currentUserId = getUserId();
      if (currentUserId) {
        const queryStore = getQueryStore();
        const userIdHash = queryStore.hashUserId(currentUserId);
        const historyRecommendation = queryStore.getOptimalEntityTypes(input.query, userIdHash);

        if (historyRecommendation && historyRecommendation.confidence > 0.7) {
          // High confidence recommendation from history
          entityTypes = historyRecommendation.entityTypes as Array<
            (typeof searchEntityTypes)[number]
          >;
          historyRecommendationUsed = true;
          thinking.push(
            `💡 History-based entity types: ${entityTypes.join(', ')} (${Math.round(historyRecommendation.confidence * 100)}% confidence)`
          );
          thinking.push(`   ${historyRecommendation.reason}`);
        } else if (historyRecommendation && historyRecommendation.confidence > 0.5) {
          // Medium confidence - log but don't use directly, let NLP enhance
          thinking.push(
            `📊 History hint: ${historyRecommendation.entityTypes.join(', ')} (${Math.round(historyRecommendation.confidence * 100)}% confidence - using NLP instead)`
          );
        }
      }
    } catch (error) {
      // Silently ignore errors - history recommendation is optional
      logger.debug('History-based entity type lookup failed', { error });
    }
  }

  if (entityTypes.length === 0) {
    const learningSystem = getLearningSystem();
    if (learningSystem) {
      try {
        const recommended = learningSystem.getRecommendedEntityTypes(input.query, undefined);
        if (recommended && recommended.length > 0) {
          const validated = validateEntityTypeCombinations(recommended);
          if (validated.length > 0) {
            entityTypes = validated as Array<(typeof searchEntityTypes)[number]>;
            thinking.push(`📚 Learning System recommended entity types: ${entityTypes.join(', ')}`);
          }
        }
      } catch (error) {
        logger.debug('Learning system entity type recommendation failed', { error });
      }
    }
  }

  if (entityTypes.length === 0) {
    // Default: Use compatible entity types
    // File types are most versatile and commonly used, so prioritize them
    // Note: Message types (message, chatMessage) and file types cannot be combined
    // We'll use file types as default, but allow NLP hints to switch to message types
    const defaultComprehensiveTypes = ['driveItem', 'site', 'list', 'listItem'];

    // Get available entity types based on token
    try {
      const availableTypes = await getAvailableEntityTypes(graphClient, defaultComprehensiveTypes);
      // Validate the returned types to ensure compatibility
      const validatedTypes = validateEntityTypeCombinations(availableTypes);
      entityTypes = validatedTypes as Array<(typeof searchEntityTypes)[number]>;
      thinking.push(`🔍 Using comprehensive entity types: ${entityTypes.join(', ')}`);
    } catch (error) {
      // Fallback to safe compatible set (file types only)
      entityTypes = ['driveItem', 'site', 'listItem'] as Array<(typeof searchEntityTypes)[number]>;
      thinking.push('⚠️ Could not determine available entity types, using safe compatible set');
    }

    // Use NLP service hint to narrow down
    if (nlpService === 'mail') {
      entityTypes = ['message'];
      thinking.push('💡 NLP detected mail query - searching messages');
    } else if (nlpService === 'files') {
      entityTypes = ['driveItem', 'site', 'listItem'];
      thinking.push('💡 NLP detected files query - searching files and sites');
    } else if (nlpService === 'teams') {
      entityTypes = ['chatMessage', 'message'];
      thinking.push('💡 NLP detected Teams query - searching chats and messages');
    } else if (hasPersonQuery || hasTeamsKeywords) {
      // For person queries, use chatMessage and message (person cannot be combined with others)
      entityTypes = ['chatMessage', 'message'];
      thinking.push('💡 Detected person/Teams query - searching chats and messages');
    }
  }

  // Validate entity types - Microsoft Graph API has strict restrictions on combinations
  // Check for person queries first (special handling)
  const hasPerson = entityTypes.includes('person');
  const hasMultipleTypes = entityTypes.length > 1;

  if (hasPerson && hasMultipleTypes) {
    // Check if query looks like a person name (capitalized words, not all caps, not a single word that's all caps)
    const queryWords = input.query.trim().split(/\s+/);
    const looksLikePersonName =
      queryWords.length >= 2 &&
      queryWords.every((word) => /^[A-ZÄÖÜ][a-zäöüß]+$/.test(word)) &&
      !/^[A-Z]{2,}$/.test(input.query.trim()); // Not all caps (like "DZBANK")

    if (looksLikePersonName) {
      // Query looks like a person name - use only person
      thinking.push(
        '⚠️ Person entity type cannot be combined with others - query looks like a person name, using only person'
      );
      entityTypes = ['person'];
    } else {
      // Query doesn't look like a person name - remove person, keep others
      thinking.push(
        '⚠️ Person entity type cannot be combined with others - query does not look like a person name, removing person type'
      );
      entityTypes = entityTypes.filter((type) => type !== 'person');
    }
  }

  // Validate entity type combinations according to Microsoft Graph API rules
  const originalEntityTypes = [...entityTypes];
  const validatedTypes = validateEntityTypeCombinations(entityTypes);
  entityTypes = validatedTypes as Array<(typeof searchEntityTypes)[number]>;

  if (
    entityTypes.length !== originalEntityTypes.length ||
    entityTypes.some((t, i) => t !== originalEntityTypes[i])
  ) {
    thinking.push(
      `⚠️ Filtered incompatible entity types: ${originalEntityTypes.join(', ')} -> ${entityTypes.join(', ')}`
    );
  }

  // Rule: Ensure at least one valid entity type
  if (entityTypes.length === 0) {
    // Use a safe default combination (file types are most versatile)
    entityTypes = ['driveItem', 'site', 'listItem'] as Array<(typeof searchEntityTypes)[number]>;
    thinking.push('⚠️ No valid entity types after validation - using safe default set');
  }

  thinking.push(`Searching in: ${entityTypes.join(', ')}`);

  // =========================================================================
  // INTELLIGENT QUERY DECOMPOSITION (MULTI-QUERY STRATEGY)
  // =========================================================================
  const queryDecompositionEnabled =
    process.env.MS365_MCP_QUERY_DECOMPOSITION_ENABLED !== 'false' && patternLearningEnabled;

  const useMultiQuery = queryDecompositionEnabled && shouldDecomposeQuery(decomposed, input.query);

  if (useMultiQuery) {
    thinking.push(`🔀 Complex query detected - using multi-query strategy`);

    try {
      const currentUserId = getUserId();
      const userIdHash = currentUserId ? getQueryStore().hashUserId(currentUserId) : 'anonymous';
      const subQueries = selectBestSubQueries(decomposed, userIdHash, 4);

      thinking.push(
        `📋 Sub-queries: ${subQueries.map((q) => `"${q.query}" (${q.reason})`).join(', ')}`
      );

      const multiQueryTimeout = parseInt(process.env.MS365_MCP_MULTIQUERY_TIMEOUT || '10000', 10);
      const multiResult = await executeMultiQuery(subQueries, graphClient, multiQueryTimeout);

      thinking.push(
        `⏱️ Multi-query completed in ${multiResult.totalDurationMs}ms (${multiResult.successCount}/${subQueries.length} successful)`
      );

      const merged = mergeSearchResults(multiResult, input.query);

      thinking.push(
        `✅ Found ${merged.uniqueHits} unique results (${merged.multiMatchHits} multi-match)`
      );

      // Record pattern for learning
      if (patternLearningEnabled && currentUserId) {
        try {
          const queryStore = getQueryStore();
          const success = merged.uniqueHits > 0;
          const duration = Date.now() - searchStartTime;
          queryStore.recordQueryPattern(userIdHash, input.query, entityTypes, success, duration);
        } catch {
          // Ignore errors
        }
      }

      // Format output for multi-query result
      const multiQueryOutput = {
        query: input.query,
        strategy: 'multi-query',
        language: {
          detected: detectedLang,
          confidence: Math.round(uqasAnalysis.languageConfidence * 100),
          crossLangSearch: searchVariants.crossLangVariants.length > 0,
        },
        nlpAnalysis: {
          intent: nlpIntent,
          service: nlpService || 'general',
          entities: nlpEntities.map((e) => ({ value: e.value, type: e.type })),
          temporal: decomposed.temporal,
          confidence: decomposed.confidence,
        },
        subQueries: merged.queryBreakdown,
        totalHits: merged.totalHits,
        uniqueHits: merged.uniqueHits,
        multiMatchHits: merged.multiMatchHits,
        entityTypes: Object.keys(merged.results),
        results: merged.results,
        suggestions: [
          '💡 Multi-query strategy found cross-referenced results',
          merged.multiMatchHits > 0
            ? `💡 ${merged.multiMatchHits} results appeared in multiple sub-queries (higher relevance)`
            : null,
        ].filter(Boolean),
      };

      return addThinkingToResponse(JSON.stringify(multiQueryOutput, null, 2), thinking);
    } catch (error) {
      thinking.push(
        `⚠️ Multi-query failed, falling back to single query: ${error instanceof Error ? error.message : String(error)}`
      );
      // Fall through to single query
    }
  }

  // =========================================================================
  // AUTOMATIC QUERY OPTIMIZATION (PRE-EXECUTION)
  // =========================================================================
  const autoOptimizationEnabled = process.env.MS365_MCP_AUTO_QUERY_OPTIMIZATION_ENABLED !== 'false';
  let autoOptimizedQuery: OptimizedQuery | null = null;

  if (autoOptimizationEnabled) {
    const currentUserIdForOpt = getUserId();
    const userIdHashForOpt = currentUserIdForOpt
      ? getQueryStore().hashUserId(currentUserIdForOpt)
      : undefined;
    autoOptimizedQuery = optimizeQueryForSearch(input.query, {
      tool: 'search',
      entityTypes,
      userIdHash: userIdHashForOpt,
    });
    addOptimizationThinking(thinking, autoOptimizedQuery);
  }

  // Improve query for better search results
  // Simplify query for Teams/person searches to just the person name
  // Use auto-optimized query as base if available, otherwise original
  let searchQuery = autoOptimizedQuery?.optimizedQuery || input.query;

  if (personName && (hasTeamsKeywords || entityTypes.includes('chatMessage'))) {
    // For Teams/person queries, simplify to just the person name for better results
    searchQuery = personName;
    thinking.push(
      `💡 Simplified query to person name: "${searchQuery}" for better Teams search results`
    );
  } else if (searchVariants.crossLangVariants.length > 0) {
    // Use cross-language combined query for bilingual search (DE/EN)
    // Only add variants if query isn't already simplified to a person name
    const crossLangKeyword = searchVariants.crossLangVariants[0];
    if (crossLangKeyword && !searchQuery.toLowerCase().includes(crossLangKeyword.toLowerCase())) {
      // Add the first cross-language variant as OR clause for better coverage
      searchQuery = `${searchQuery} OR ${crossLangKeyword}`;
      thinking.push(
        `🌐 Added cross-language variant: "${crossLangKeyword}" (${detectedLang === 'de' ? 'EN' : 'DE'})`
      );
    }
  }

  // Format KQL query to ensure proper syntax (handles property filters with spaces, OR/AND operators)
  searchQuery = formatKQLQuery(searchQuery);
  if (searchQuery !== input.query) {
    thinking.push(`💡 Formatted KQL query: "${input.query}" -> "${searchQuery}"`);
  }

  // Build the search request
  const searchRequest = {
    requests: [
      {
        entityTypes: entityTypes,
        query: {
          queryString: searchQuery,
        },
        from: input.from || 0,
        size: input.size || 25,
        trimDuplicates: input.trimDuplicates !== false,
        ...(input.fields && { fields: input.fields }),
        ...(input.sortBy && {
          sortProperties: [{ name: input.sortBy, isDescending: true }],
        }),
      },
    ],
  };

  try {
    const result = await callGraph(graphClient, 'POST', '/search/query', undefined, searchRequest);
    const parsedResult = JSON.parse(result);

    // Extract and format results for better readability
    const formattedResults: Record<string, unknown[]> = {};
    let totalHits = 0;

    if (parsedResult.value && Array.isArray(parsedResult.value)) {
      for (const response of parsedResult.value) {
        if (response.hitsContainers && Array.isArray(response.hitsContainers)) {
          for (const container of response.hitsContainers) {
            totalHits += container.total || 0;
            if (container.hits && Array.isArray(container.hits)) {
              for (const hit of container.hits) {
                const entityType = hit.resource?.['@odata.type'] || 'unknown';
                if (!formattedResults[entityType]) {
                  formattedResults[entityType] = [];
                }
                formattedResults[entityType].push({
                  id: hit.resource?.id,
                  summary: hit.summary,
                  rank: hit.rank,
                  ...hit.resource,
                });
              }
            }
          }
        }
      }
    }

    // Record primary query result for learning (so optimizer can prefer this formulation next time)
    if (autoOptimizedQuery && searchQuery !== input.query) {
      const currentUserId = getUserId();
      const userIdHash = currentUserId ? getQueryStore().hashUserId(currentUserId) : undefined;
      if (userIdHash) {
        recordQueryOptimizationResult(
          input.query,
          searchQuery,
          totalHits > 0,
          userIdHash,
          'search'
        );
      }
    }

    // Auto-improve: if primary query returned no results, try optimizer variants and merge
    const variantFallbackEnabled =
      process.env.MS365_MCP_AUTO_QUERY_OPTIMIZATION_ENABLED !== 'false';
    if (
      totalHits === 0 &&
      variantFallbackEnabled &&
      autoOptimizedQuery?.variants &&
      autoOptimizedQuery.variants.length > 0
    ) {
      const seenIds = new Set<string>();
      for (const r of Object.values(formattedResults)) {
        for (const item of r as Array<{ id?: string }>) {
          if (item.id) seenIds.add(String(item.id));
        }
      }
      const variantsToTry = autoOptimizedQuery.variants.slice(0, 3);
      for (const variantQuery of variantsToTry) {
        const vq = formatKQLQuery(variantQuery);
        if (!vq || vq === searchQuery) continue;
        try {
          const variantRequest = {
            requests: [
              {
                entityTypes,
                query: { queryString: vq },
                from: input.from || 0,
                size: input.size || 25,
                trimDuplicates: input.trimDuplicates !== false,
                ...(input.fields && { fields: input.fields }),
                ...(input.sortBy && {
                  sortProperties: [{ name: input.sortBy, isDescending: true }],
                }),
              },
            ],
          };
          const variantResult = await callGraph(
            graphClient,
            'POST',
            '/search/query',
            undefined,
            variantRequest
          );
          const variantParsed = JSON.parse(variantResult);
          let variantHits = 0;
          if (variantParsed.value && Array.isArray(variantParsed.value)) {
            for (const response of variantParsed.value) {
              if (response.hitsContainers && Array.isArray(response.hitsContainers)) {
                for (const container of response.hitsContainers) {
                  if (container.hits && Array.isArray(container.hits)) {
                    for (const hit of container.hits) {
                      const rid = hit.resource?.id;
                      if (rid && !seenIds.has(String(rid))) {
                        seenIds.add(String(rid));
                        variantHits++;
                        const entityType = hit.resource?.['@odata.type'] || 'unknown';
                        if (!formattedResults[entityType]) {
                          formattedResults[entityType] = [];
                        }
                        formattedResults[entityType].push({
                          id: hit.resource?.id,
                          summary: hit.summary,
                          rank: hit.rank,
                          ...hit.resource,
                        });
                      }
                    }
                  }
                }
              }
            }
          }
          totalHits += variantHits;
          if (variantHits > 0) {
            thinking.push(
              `🧠 Auto-improved: variant query "${vq}" returned ${variantHits} additional result(s)`
            );
            const currentUserId = getUserId();
            const userIdHash = currentUserId
              ? getQueryStore().hashUserId(currentUserId)
              : undefined;
            if (userIdHash) {
              recordQueryOptimizationResult(input.query, variantQuery, true, userIdHash, 'search');
            }
            break; // One successful variant is enough; learning will prefer it next time
          }
        } catch (err) {
          logger.debug('Variant query failed', { variant: variantQuery, error: err });
        }
      }
    }

    thinking.push(
      `Found ${totalHits} results across ${Object.keys(formattedResults).length} entity types`
    );

    // If chatMessage results found, fetch full message content for better context
    const chatMessageResults =
      formattedResults['#microsoft.graph.chatMessage'] || formattedResults['chatMessage'] || [];
    const chatIds = new Set<string>();

    // Extract chat IDs from chatMessage results
    for (const msg of chatMessageResults) {
      const msgAny = msg as any;
      if (msgAny.id) {
        let chatId: string | null = null;

        // Try different ID formats
        // Format 1: "19:chatId_messageId@thread.v2"
        if (msgAny.id.includes(':')) {
          const parts = msgAny.id.split(':');
          if (parts.length > 1) {
            const afterColon = parts[1];
            if (afterColon.includes('_')) {
              chatId = afterColon.split('_')[0];
            } else if (afterColon.includes('@')) {
              chatId = afterColon.split('@')[0];
            }
          }
        }
        // Format 2: "chatId/messageId" or similar path format
        else if (msgAny.id.includes('/')) {
          const parts = msgAny.id.split('/');
          if (parts.length > 1) {
            chatId = parts[0];
          }
        }
        // Format 3: Check if chatId is in the resource itself
        else if (msgAny.chatId) {
          chatId = msgAny.chatId;
        }

        if (chatId) {
          chatIds.add(chatId);
        }
      }
    }

    // Fetch full messages for found chats (limit to avoid too many requests)
    const chatMessagesWithContent: Record<string, unknown[]> = {};
    if (chatIds.size > 0 && chatMessageResults.length > 0) {
      thinking.push(`Fetching full message content for ${Math.min(chatIds.size, 5)} chats...`);

      for (const chatId of Array.from(chatIds).slice(0, 5)) {
        try {
          const messagesResult = await callGraph(
            graphClient,
            'GET',
            `/me/chats/${chatId}/messages`,
            { $top: '10', $orderby: 'createdDateTime desc' }
          );
          const messagesData = JSON.parse(messagesResult);

          if (messagesData.value && Array.isArray(messagesData.value)) {
            chatMessagesWithContent[chatId] = messagesData.value.map((msg: any) => ({
              id: msg.id,
              from: msg.from?.user?.displayName || msg.from?.user?.userPrincipalName || 'Unknown',
              fromEmail: msg.from?.user?.userPrincipalName,
              content: msg.body?.content || '',
              contentType: msg.body?.contentType || 'text',
              createdDateTime: msg.createdDateTime,
              importance: msg.importance,
            }));
          }
        } catch (error) {
          logger.warn(`Failed to fetch messages for chat ${chatId}: ${error}`);
        }
      }

      if (Object.keys(chatMessagesWithContent).length > 0) {
        thinking.push(
          `Retrieved full message content for ${Object.keys(chatMessagesWithContent).length} chats`
        );
      }
    }

    // Provide NLP-informed guidance on which tools to use based on results
    const toolSuggestions: string[] = [];

    // NLP-based suggestions when no results found
    if (totalHits === 0) {
      // Use NLP service context for better suggestions
      if (nlpService === 'calendar') {
        toolSuggestions.push('💡 Use "calendar" tool with action "list" to list calendar events');
        toolSuggestions.push('💡 Use "assistant" tool with action "my-day" for today\'s schedule');
        toolSuggestions.push('💡 Use "assistant" tool with action "my-week" for week overview');
      } else if (nlpService === 'mail') {
        toolSuggestions.push('💡 Use "email" tool with action "list" to list emails');
        toolSuggestions.push('💡 Use "email" tool with action "search" for email search');
      } else if (nlpService === 'files') {
        toolSuggestions.push('💡 Use "files" tool with action "search" to search files');
        toolSuggestions.push('💡 Use "files" tool with action "list" to list files');
      } else if (nlpService === 'teams') {
        toolSuggestions.push('💡 Use "teams" tool with action "chats" to list Teams chats');
        toolSuggestions.push(
          '💡 Try compound tool "find-messages-with-person" for Teams chats with a specific person'
        );
      } else if (nlpService === 'tasks') {
        toolSuggestions.push('💡 Use "tasks" tool with action "lists" to list task lists');
      } else if (nlpService === 'contacts') {
        toolSuggestions.push('💡 Use "contacts" tool with action "list" to list contacts');
      }

      // Person-specific suggestions (only if we have a valid person name)
      if (personName) {
        toolSuggestions.push(
          `💡 Try compound tool "find-emails-with-person" to find emails with ${personName}`
        );
        toolSuggestions.push(
          `💡 Try compound tool "discover-person" for comprehensive info about ${personName}`
        );
      }

      // Teams-specific suggestions
      if (hasTeamsKeywords) {
        toolSuggestions.push(
          '💡 Try "teams" tool with action "chats" (includes last messages) to list all chats'
        );
      }

      // General fallback if no specific suggestions
      if (toolSuggestions.length === 0) {
        toolSuggestions.push(
          '💡 Try using specific tools: "calendar" for events, "email" for messages, "files" for documents'
        );
        toolSuggestions.push(
          '💡 Use "assistant" tool with action "discover" for comprehensive search'
        );
      }
    }

    // Suggest tools based on found entity types
    for (const entityType of Object.keys(formattedResults)) {
      if (entityType.includes('message') && !entityType.includes('chat')) {
        toolSuggestions.push('Use "email" tool for detailed email operations');
      }
      if (entityType.includes('event')) {
        toolSuggestions.push('Use "calendar" tool for calendar operations');
      }
      if (entityType.includes('driveItem')) {
        toolSuggestions.push('Use "files" tool for file operations');
      }
      if (entityType.includes('site') || entityType.includes('list')) {
        toolSuggestions.push('Use "sharepoint" tool for SharePoint operations');
      }
      if (entityType.includes('chatMessage')) {
        toolSuggestions.push('Use "teams" tool with action "chat-messages" for Teams chat details');
        if (personName) {
          toolSuggestions.push(
            `💡 Try compound tool "find-messages-with-person" for all chats with ${personName}`
          );
        }
      }
      if (entityType.includes('person')) {
        if (personName) {
          toolSuggestions.push(
            `💡 Try compound tool "discover-person" for comprehensive info about ${personName}`
          );
        }
      }
    }

    if (toolSuggestions.length > 0) {
      thinking.push('💡 Suggested next tools: ' + [...new Set(toolSuggestions)].join(', '));
    }

    // Build structured suggestedNextTools for LLM/client (tool, action, IDs, reason)
    type SuggestedNextTool = {
      tool: string;
      action?: string;
      messageId?: string;
      eventId?: string;
      chatId?: string;
      driveId?: string;
      itemId?: string;
      reason: string;
    };
    const suggestedNextTools: SuggestedNextTool[] = [];

    for (const entityType of Object.keys(formattedResults)) {
      const items = formattedResults[entityType] as Array<Record<string, unknown>>;
      if (!items?.length) continue;
      const first = items[0];
      const id = first?.id as string | undefined;
      if (!id) continue;
      if (entityType.includes('message') && !entityType.includes('chat')) {
        suggestedNextTools.push({
          tool: 'email',
          action: 'get',
          messageId: id,
          reason: 'Open this email for full content',
        });
      } else if (entityType.includes('event')) {
        suggestedNextTools.push({
          tool: 'calendar',
          action: 'get',
          eventId: id,
          reason: 'Get full event details',
        });
      } else if (entityType.includes('chatMessage') || entityType.includes('chat')) {
        const chatId = (first?.chatId ?? first?.parentReference?.id) as string | undefined;
        if (chatId) {
          suggestedNextTools.push({
            tool: 'teams',
            action: 'chat-messages',
            chatId,
            reason: 'View more messages in this chat',
          });
        }
      } else if (entityType.includes('driveItem') || entityType.includes('listItem')) {
        const driveId = first?.parentReference?.driveId as string | undefined;
        suggestedNextTools.push({
          tool: 'files',
          action: 'get',
          driveId: driveId ?? undefined,
          itemId: id,
          reason: 'Get file details or download',
        });
      }
    }

    // Add generic suggestions when no result-based ones
    if (suggestedNextTools.length === 0) {
      if (Object.keys(formattedResults).some((k) => k.includes('message'))) {
        suggestedNextTools.push({ tool: 'email', action: 'list', reason: 'List more emails' });
      }
      if (Object.keys(formattedResults).some((k) => k.includes('event'))) {
        suggestedNextTools.push({ tool: 'calendar', action: 'view', reason: 'View calendar' });
      }
      if (
        Object.keys(formattedResults).some((k) => k.includes('driveItem') || k.includes('listItem'))
      ) {
        suggestedNextTools.push({ tool: 'files', action: 'search', reason: 'Search more files' });
      }
      if (Object.keys(formattedResults).some((k) => k.includes('chat'))) {
        suggestedNextTools.push({ tool: 'teams', action: 'chats', reason: 'List Teams chats' });
      }
      if (suggestedNextTools.length === 0 && totalHits === 0) {
        suggestedNextTools.push(
          { tool: 'search', reason: 'Try a different query or entity types' },
          { tool: 'assistant', action: 'discover', reason: 'Comprehensive discovery' }
        );
      }
    }

    // =========================================================================
    // AUTOMATIC SUGGESTION EXECUTION - Execute suggested tools automatically
    // =========================================================================
    const autoExecutedResults: Record<string, unknown> = {};
    const uniqueSuggestions = [...new Set(toolSuggestions)];

    // Check for specific suggestions and execute them automatically
    const shouldExecuteCalendarList = uniqueSuggestions.some(
      (s) => s.includes('calendar') && s.includes('list')
    );
    const shouldExecuteCalendarGet = uniqueSuggestions.some(
      (s) => s.includes('calendar') && s.includes('get')
    );
    const shouldExecuteMyDay = uniqueSuggestions.some(
      (s) => s.includes('assistant') && s.includes('my-day')
    );

    if (shouldExecuteCalendarList || shouldExecuteMyDay || shouldExecuteCalendarGet) {
      thinking.push('🔄 Automatically executing suggested tools...');

      const autoExecutionPromises: Array<Promise<void>> = [];

      // Execute calendar list if suggested (needs to complete before calendar get)
      if (shouldExecuteCalendarList) {
        autoExecutionPromises.push(
          (async () => {
            try {
              thinking.push('📅 Executing: calendar tool with action "list"');
              const calendarListResult = await handleCalendar(
                { action: 'list', top: 25 },
                graphClient,
                _readOnly
              );
              // Parse the result to extract calendar events
              try {
                const parsed = JSON.parse(calendarListResult);
                // Handle different response formats
                if (parsed.data?.formatted?.value) {
                  autoExecutedResults.calendarList = parsed.data.formatted.value;
                } else if (parsed.data?.raw?.value) {
                  autoExecutedResults.calendarList = parsed.data.raw.value;
                } else if (parsed.data?.value) {
                  autoExecutedResults.calendarList = parsed.data.value;
                } else if (parsed.data) {
                  autoExecutedResults.calendarList = parsed.data;
                } else if (parsed.value) {
                  autoExecutedResults.calendarList = parsed.value;
                } else {
                  autoExecutedResults.calendarList = parsed;
                }
              } catch {
                autoExecutedResults.calendarList = calendarListResult;
              }
            } catch (error) {
              thinking.push(
                `⚠️ Failed to execute calendar list: ${error instanceof Error ? error.message : String(error)}`
              );
            }
          })()
        );
      }

      // Execute my-day if suggested (can run in parallel)
      if (shouldExecuteMyDay) {
        autoExecutionPromises.push(
          (async () => {
            try {
              thinking.push('📊 Executing: assistant tool with action "my-day"');
              const myDayResult = await handleAssistant(
                { action: 'my-day' },
                graphClient,
                _readOnly
              );
              // Parse the result to extract my-day data
              try {
                const parsed = JSON.parse(myDayResult);
                // Handle different response formats
                if (parsed.data?.formatted) {
                  autoExecutedResults.myDay = parsed.data.formatted;
                } else if (parsed.data?.raw) {
                  autoExecutedResults.myDay = parsed.data.raw;
                } else if (parsed.data) {
                  autoExecutedResults.myDay = parsed.data;
                } else {
                  autoExecutedResults.myDay = parsed;
                }
              } catch {
                autoExecutedResults.myDay = myDayResult;
              }
            } catch (error) {
              thinking.push(
                `⚠️ Failed to execute my-day: ${error instanceof Error ? error.message : String(error)}`
              );
            }
          })()
        );
      }

      // Wait for calendar list and my-day to complete first
      await Promise.allSettled(autoExecutionPromises);

      // Execute calendar get if suggested and we have events (after calendar list completes)
      if (shouldExecuteCalendarGet) {
        // Check if we have events from calendar list or search results
        const hasEvents =
          autoExecutedResults.calendarList ||
          formattedResults['#microsoft.graph.event'] ||
          formattedResults['event'];

        if (hasEvents) {
          try {
            // Get the first event ID from results
            let eventId: string | undefined;

            // Try to get event ID from calendar list results
            if (autoExecutedResults.calendarList) {
              const calendarData = autoExecutedResults.calendarList as {
                value?: Array<{ id?: string }>;
              };
              if (calendarData?.value?.[0]?.id) {
                eventId = calendarData.value[0].id;
              }
            }

            // Try to get event ID from search results
            if (!eventId) {
              const events =
                formattedResults['#microsoft.graph.event'] || formattedResults['event'];
              if (Array.isArray(events) && events.length > 0) {
                const firstEvent = events[0] as { id?: string };
                if (firstEvent?.id) {
                  eventId = firstEvent.id;
                }
              }
            }

            if (eventId) {
              thinking.push(`📅 Executing: calendar tool with action "get" for event ${eventId}`);
              const calendarGetResult = await handleCalendar(
                { action: 'get', eventId },
                graphClient,
                _readOnly
              );
              // Parse the result
              try {
                const parsed = JSON.parse(calendarGetResult);
                // Handle different response formats
                if (parsed.data?.formatted) {
                  autoExecutedResults.calendarGet = parsed.data.formatted;
                } else if (parsed.data?.raw) {
                  autoExecutedResults.calendarGet = parsed.data.raw;
                } else if (parsed.data) {
                  autoExecutedResults.calendarGet = parsed.data;
                } else {
                  autoExecutedResults.calendarGet = parsed;
                }
              } catch {
                autoExecutedResults.calendarGet = calendarGetResult;
              }
            } else {
              thinking.push('⚠️ Cannot execute calendar get: No event ID available');
            }
          } catch (error) {
            thinking.push(
              `⚠️ Failed to execute calendar get: ${error instanceof Error ? error.message : String(error)}`
            );
          }
        } else {
          thinking.push('⚠️ Cannot execute calendar get: No events found to get details for');
        }
      }

      if (Object.keys(autoExecutedResults).length > 0) {
        thinking.push(
          `✅ Auto-executed ${Object.keys(autoExecutedResults).length} suggested tool(s)`
        );
        // Merge auto-executed results into formatted results and update totalHits
        if (autoExecutedResults.calendarList) {
          if (!formattedResults['#microsoft.graph.event']) {
            formattedResults['#microsoft.graph.event'] = [];
          }
          const calendarData = autoExecutedResults.calendarList as {
            value?: unknown[];
          };
          if (calendarData?.value) {
            formattedResults['#microsoft.graph.event'].push(...calendarData.value);
            totalHits += calendarData.value.length;
          }
        }
        if (autoExecutedResults.calendarGet) {
          if (!formattedResults['eventDetails']) {
            formattedResults['eventDetails'] = [];
          }
          formattedResults['eventDetails'].push(autoExecutedResults.calendarGet);
          totalHits += 1;
        }
        if (autoExecutedResults.myDay) {
          formattedResults['myDaySummary'] = [autoExecutedResults.myDay];
          // Count items in myDay summary
          const myDayData = autoExecutedResults.myDay as Record<string, unknown>;
          const todayEvents = myDayData.todayEvents as { value?: unknown[] } | undefined;
          const recentEmails = myDayData.recentEmails as { value?: unknown[] } | undefined;
          if (todayEvents?.value && Array.isArray(todayEvents.value)) {
            totalHits += todayEvents.value.length;
          }
          if (recentEmails?.value && Array.isArray(recentEmails.value)) {
            totalHits += recentEmails.value.length;
          }
        }
      }
    }

    // Extract all driveItems for download link generation
    const driveItems: unknown[] = [];
    const driveItemTypes = [
      '#microsoft.graph.driveItem',
      'driveItem',
      '#microsoft.graph.listItem',
      'listItem',
    ];

    for (const entityType of Object.keys(formattedResults)) {
      if (driveItemTypes.some((type) => entityType.includes(type))) {
        driveItems.push(...formattedResults[entityType]);
      }
    }

    // Generate download links for files
    const documentLinks: Array<{
      fileName: string;
      webUrl: string;
      downloadUrl?: string;
      type: string;
    }> = [];
    const importantDocuments: Array<{
      name: string;
      type: string;
      webUrl: string;
      downloadUrl?: string;
      relevance: number;
    }> = [];

    if (downloadLinkGenerator && driveItems.length > 0) {
      thinking.push(`🔗 Generating download links for ${driveItems.length} files...`);
      try {
        const requestTokens = getRequestTokens();
        const accessToken = requestTokens?.accessToken;

        const enrichedResults = await downloadLinkGenerator.addDownloadLinksToResults(
          driveItems,
          accessToken
        );

        // Extract links and webUrls
        for (const item of enrichedResults) {
          if (typeof item === 'object' && item !== null) {
            const obj = item as Record<string, unknown>;
            const downloadLink = obj.downloadLink as
              | { fileName: string; downloadUrl: string; webUrl?: string }
              | undefined;

            if (downloadLink) {
              documentLinks.push({
                fileName: downloadLink.fileName,
                webUrl: downloadLink.webUrl || (obj.webUrl as string) || '',
                downloadUrl: downloadLink.downloadUrl,
                type: (obj['@odata.type'] as string) || 'driveItem',
              });
            } else if (obj.webUrl) {
              // Include items with webUrl even if no download link
              const name = (obj.name as string) || (obj.title as string) || 'Unknown';
              documentLinks.push({
                fileName: name,
                webUrl: obj.webUrl as string,
                type: (obj['@odata.type'] as string) || 'unknown',
              });
            }
          }
        }

        thinking.push(`✅ Generated ${documentLinks.length} document links`);
      } catch (error) {
        logger.warn(`Failed to generate download links: ${error}`);
        thinking.push(
          `⚠️ Could not generate download links: ${error instanceof Error ? error.message : String(error)}`
        );
      }
    }

    // Extract webUrls from all results for sources
    const allSources = new Set<string>();
    const sourcesWithUrls: Array<{
      type: string;
      name: string;
      webUrl?: string;
      downloadUrl?: string;
      relevance: number;
    }> = [];

    for (const [entityType, items] of Object.entries(formattedResults)) {
      for (const item of items) {
        if (typeof item === 'object' && item !== null) {
          const obj = item as Record<string, unknown>;
          const webUrl = obj.webUrl as string | undefined;
          const webLink = obj.webLink as string | undefined;
          const name =
            (obj.name as string) ||
            (obj.subject as string) ||
            (obj.title as string) ||
            (obj.displayName as string) ||
            'Unknown';

          if (webUrl || webLink) {
            const url = webUrl || webLink || '';
            allSources.add(url);

            // Find corresponding download link if available
            const downloadLink = documentLinks.find((link) => link.webUrl === url);

            sourcesWithUrls.push({
              type: entityType,
              name,
              webUrl: url,
              downloadUrl: downloadLink?.downloadUrl,
              relevance: (obj.rank as number) || 0.5,
            });
          } else {
            // Still add as source even without URL
            allSources.add(`${entityType}:${name}`);
            sourcesWithUrls.push({
              type: entityType,
              name,
              relevance: (obj.rank as number) || 0.5,
            });
          }
        }
      }
    }

    // Identify important documents (top 10 by relevance)
    const sortedSources = sourcesWithUrls
      .filter((s) => s.webUrl && (s.type.includes('driveItem') || s.type.includes('listItem')))
      .sort((a, b) => b.relevance - a.relevance)
      .slice(0, 10)
      .map((s) => ({ ...s, webUrl: s.webUrl! })); // Filter ensures webUrl is defined

    importantDocuments.push(...sortedSources);

    // =========================================================================
    // RECORD QUERY PATTERN FOR LEARNING (USER-SPECIFIC)
    // =========================================================================
    if (patternLearningEnabled) {
      try {
        const currentUserId = getUserId();
        if (currentUserId) {
          const queryStore = getQueryStore();
          const userIdHash = queryStore.hashUserId(currentUserId);
          const success = totalHits > 0;
          const duration = Date.now() - searchStartTime;

          // Record pattern for future learning
          queryStore.recordQueryPattern(userIdHash, input.query, entityTypes, success, duration);

          // Record optimization result for automatic query optimization learning
          if (autoOptimizedQuery && autoOptimizedQuery.optimizedQuery !== input.query) {
            recordQueryOptimizationResult(
              input.query,
              autoOptimizedQuery.optimizedQuery,
              success,
              userIdHash,
              'search'
            );
          }

          if (success && !historyRecommendationUsed) {
            thinking.push(`📚 Recorded successful pattern for future learning`);
          }
        }
      } catch (error) {
        // Silently ignore errors - pattern recording is optional
        logger.debug('Failed to record query pattern', { error });
      }
    }

    // Learning System: learn from this search (works without Discovery Tools when ensureLearningSystemInitialized was called)
    const learningSystem = getLearningSystem();
    if (learningSystem) {
      try {
        const items = Object.values(formattedResults).flat();
        await learningSystem.learnFromSearch(
          input.query,
          {
            items,
            sources: Array.from(allSources),
            query: input.query,
            entityTypes,
            totalResults: totalHits,
          },
          undefined,
          entityTypes.join(',')
        );
      } catch (error) {
        logger.debug('Learning system learnFromSearch failed', { error });
      }
    }

    const output: Record<string, unknown> = {
      query: input.query,
      language: {
        detected: detectedLang,
        confidence: Math.round(uqasAnalysis.languageConfidence * 100),
        crossLangSearch: searchVariants.crossLangVariants.length > 0,
      },
      nlpAnalysis: {
        intent: nlpIntent,
        service: nlpService || 'general',
        entities: nlpEntities.map((e) => ({ value: e.value, type: e.type })),
        temporal: decomposed.temporal,
        confidence: decomposed.confidence,
      },
      ...(autoOptimizedQuery &&
        autoOptimizedQuery.optimizations.length > 0 && {
          optimizationSummary: buildOptimizationSummary(autoOptimizedQuery),
        }),
      totalHits,
      entityTypes: Object.keys(formattedResults),
      results: formattedResults,
      suggestions: [...new Set(toolSuggestions)],
      suggestedNextTools,
      sources: {
        importantDocuments,
        allSources: Array.from(allSources),
        totalSources: allSources.size,
      },
      documentLinks: documentLinks.slice(0, 50), // Limit to top 50
    };

    // Add full chat messages if fetched
    if (Object.keys(chatMessagesWithContent).length > 0) {
      output.chatMessagesWithContent = chatMessagesWithContent;
      output.note =
        'Full message content has been fetched for the chats found in search results. Use "teams" tool with action "chat-messages" and the chatId to get more messages.';
    }

    return addThinkingToResponse(JSON.stringify(output, null, 2), thinking);
  } catch (error) {
    thinking.push(`Search error: ${error instanceof Error ? error.message : String(error)}`);
    throw error;
  }
}

// ============================================================================
// PRODUCT-BASED SEARCH HANDLER
// ============================================================================

/**
 * Detect products from search results based on entity types
 */
function detectProductsFromSearchResults(results: Record<string, unknown[]>): string[] {
  const detectedProducts = new Set<string>();

  // Map entity types to products
  for (const [entityType, items] of Object.entries(results)) {
    if (items && items.length > 0) {
      for (const mapping of PRODUCT_MAPPINGS) {
        if (mapping.entityTypes.includes(entityType)) {
          detectedProducts.add(mapping.product);
        }
      }
    }
  }

  // Also check for specific entity type patterns
  const entityTypeLower = Object.keys(results).join(',').toLowerCase();

  if (entityTypeLower.includes('message') && !entityTypeLower.includes('chat')) {
    detectedProducts.add('Outlook');
  }
  if (entityTypeLower.includes('event')) {
    detectedProducts.add('Calendar');
  }
  if (entityTypeLower.includes('driveitem')) {
    detectedProducts.add('OneDrive');
  }
  if (entityTypeLower.includes('site') || entityTypeLower.includes('list')) {
    detectedProducts.add('SharePoint');
  }
  if (entityTypeLower.includes('chatmessage')) {
    detectedProducts.add('Teams');
  }
  if (entityTypeLower.includes('person')) {
    detectedProducts.add('Users');
  }

  return Array.from(detectedProducts);
}

/**
 * Search a specific product using its API endpoint
 */
async function searchProduct(
  product: string,
  query: string,
  graphClient: GraphClient,
  topResults: number
): Promise<ProductSearchResult> {
  const mapping = PRODUCT_MAPPINGS.find((m) => m.product === product);
  if (!mapping) {
    return {
      product,
      resultCount: 0,
      topResults: [],
    };
  }

  const topResultsList: ProductSearchResult['topResults'] = [];

  try {
    if (product === 'Outlook') {
      // Search messages
      const result = await callGraph(graphClient, 'GET', '/me/messages', {
        $search: formatSearchQuery(query, 'displayName', 'email'),
        $top: String(topResults),
        $select: 'id,subject,from,receivedDateTime,bodyPreview,webLink,hasAttachments',
      });
      const messages = JSON.parse(result);
      const items = messages.value || [];

      for (const item of items.slice(0, topResults)) {
        topResultsList.push({
          title: item.subject || 'No subject',
          summary: item.bodyPreview || '',
          relevance: 100 - topResultsList.length * 5,
          webUrl: item.webLink,
          metadata: {
            from: item.from?.emailAddress?.address,
            receivedDateTime: item.receivedDateTime,
            hasAttachments: item.hasAttachments,
          },
        });
      }
    } else if (product === 'Calendar') {
      // Search calendar events
      const now = new Date();
      const startDate = new Date(
        now.getFullYear() - 1,
        now.getMonth(),
        now.getDate()
      ).toISOString();
      const endDate = new Date(now.getFullYear() + 1, now.getMonth(), now.getDate()).toISOString();

      const result = await callGraph(graphClient, 'GET', '/me/calendarView', {
        startDateTime: startDate,
        endDateTime: endDate,
        $top: String(topResults),
        $select: 'id,subject,start,end,location,organizer,attendees,webLink',
      });
      const events = JSON.parse(result);
      const items = events.value || [];

      // Filter events that match the query
      const queryLower = query.toLowerCase();
      const matchingEvents = items.filter((event: Record<string, unknown>) => {
        const subject = (event.subject || '').toString().toLowerCase();
        const location = (
          (event.location as { displayName?: string } | undefined)?.displayName || ''
        )
          .toString()
          .toLowerCase();
        return subject.includes(queryLower) || location.includes(queryLower);
      });

      for (const item of matchingEvents.slice(0, topResults)) {
        topResultsList.push({
          title: item.subject || 'No subject',
          summary: `Start: ${item.start?.dateTime}, Location: ${item.location?.displayName || 'N/A'}`,
          relevance: 100 - topResultsList.length * 5,
          webUrl: item.webLink,
          metadata: {
            start: item.start?.dateTime,
            end: item.end?.dateTime,
            location: item.location?.displayName,
          },
        });
      }
    } else if (product === 'OneDrive') {
      // Search OneDrive files
      const result = await callGraph(
        graphClient,
        'GET',
        `/me/drive/root/search(q='${encodeURIComponent(query)}')`,
        {
          $top: String(topResults),
          $select: 'id,name,webUrl,size,lastModifiedDateTime,createdBy',
        }
      );
      const files = JSON.parse(result);
      const items = files.value || [];

      for (const item of items.slice(0, topResults)) {
        topResultsList.push({
          title: item.name || 'Unnamed file',
          summary: `Size: ${item.size || 0} bytes, Modified: ${item.lastModifiedDateTime || 'N/A'}`,
          relevance: 100 - topResultsList.length * 5,
          webUrl: item.webUrl,
          metadata: {
            size: item.size,
            lastModifiedDateTime: item.lastModifiedDateTime,
            createdBy: item.createdBy?.user?.displayName,
          },
        });
      }
    } else if (product === 'SharePoint') {
      // Search SharePoint sites and lists using search API
      const searchRequest = {
        requests: [
          {
            entityTypes: ['site', 'list', 'listItem'],
            query: {
              queryString: query,
            },
            from: 0,
            size: topResults,
          },
        ],
      };

      try {
        const result = await callGraph(
          graphClient,
          'POST',
          '/search/query',
          undefined,
          searchRequest
        );
        const parsedResult = JSON.parse(result);
        const items: unknown[] = [];

        if (parsedResult.value && Array.isArray(parsedResult.value)) {
          for (const response of parsedResult.value) {
            if (response.hitsContainers && Array.isArray(response.hitsContainers)) {
              for (const container of response.hitsContainers) {
                if (container.hits && Array.isArray(container.hits)) {
                  for (const hit of container.hits) {
                    items.push({
                      ...hit.resource,
                      summary: hit.summary,
                      rank: hit.rank,
                    });
                  }
                }
              }
            }
          }
        }

        for (const item of items.slice(0, topResults)) {
          const itemAny = item as Record<string, unknown>;
          topResultsList.push({
            title: (itemAny.displayName || itemAny.name || 'Unnamed site') as string,
            summary: (itemAny.summary || `Site: ${itemAny.name || 'N/A'}`) as string,
            relevance: (itemAny.rank as number) || 100 - topResultsList.length * 5,
            webUrl: (itemAny.webUrl || itemAny.webLink) as string | undefined,
            metadata: {
              name: itemAny.name,
              displayName: itemAny.displayName,
              '@odata.type': itemAny['@odata.type'],
            },
          });
        }
      } catch (error) {
        logger.debug(`SharePoint search failed, trying direct sites API: ${error}`);
        // Fallback to direct sites API
        const params: Record<string, string> = {
          $top: String(topResults),
        };
        if (query) {
          params.search = query;
        }
        const result = await callGraph(graphClient, 'GET', '/sites', params);
        const sites = JSON.parse(result);
        const items = sites.value || [];

        for (const item of items.slice(0, topResults)) {
          topResultsList.push({
            title: item.displayName || item.name || 'Unnamed site',
            summary: `Site: ${item.name || 'N/A'}`,
            relevance: 100 - topResultsList.length * 5,
            webUrl: item.webUrl,
            metadata: {
              name: item.name,
              displayName: item.displayName,
            },
          });
        }
      }
    } else if (product === 'Teams') {
      // Search Teams chats
      const result = await callGraph(graphClient, 'GET', '/me/chats', {
        $top: String(topResults),
        $select: 'id,topic,chatType,lastUpdatedDateTime',
      });
      const chats = JSON.parse(result);
      const items = chats.value || [];

      // Filter chats that match the query
      const queryLower = query.toLowerCase();
      const matchingChats = items.filter((chat: Record<string, unknown>) => {
        const topic = (chat.topic || '').toString().toLowerCase();
        return topic.includes(queryLower);
      });

      for (const item of matchingChats.slice(0, topResults)) {
        topResultsList.push({
          title: item.topic || 'Untitled chat',
          summary: `Type: ${item.chatType}, Updated: ${item.lastUpdatedDateTime || 'N/A'}`,
          relevance: 100 - topResultsList.length * 5,
          metadata: {
            chatType: item.chatType,
            lastUpdatedDateTime: item.lastUpdatedDateTime,
          },
        });
      }
    } else if (product === 'OneNote') {
      // Search OneNote pages
      const result = await callGraph(graphClient, 'GET', '/me/onenote/pages', {
        $top: String(topResults),
        $select: 'id,title,createdDateTime,lastModifiedDateTime,contentUrl',
      });
      const pages = JSON.parse(result);
      const items = pages.value || [];

      // Filter pages that match the query
      const queryLower = query.toLowerCase();
      const matchingPages = items.filter((page: Record<string, unknown>) => {
        const title = (page.title || '').toString().toLowerCase();
        return title.includes(queryLower);
      });

      for (const item of matchingPages.slice(0, topResults)) {
        topResultsList.push({
          title: item.title || 'Untitled page',
          summary: `Created: ${item.createdDateTime || 'N/A'}, Modified: ${item.lastModifiedDateTime || 'N/A'}`,
          relevance: 100 - topResultsList.length * 5,
          webUrl: item.contentUrl,
          metadata: {
            createdDateTime: item.createdDateTime,
            lastModifiedDateTime: item.lastModifiedDateTime,
          },
        });
      }
    } else if (product === 'Users') {
      // Search users - Microsoft Graph API requires property:value format for $search
      const params: Record<string, string> = {
        $search: formatSearchQuery(query, 'displayName'),
        $top: String(topResults),
        $select: 'id,displayName,mail,userPrincipalName,jobTitle,department',
      };
      const result = await callGraph(graphClient, 'GET', '/users', params, undefined, {
        ConsistencyLevel: 'eventual',
      });
      const users = JSON.parse(result);
      const items = users.value || [];

      for (const item of items.slice(0, topResults)) {
        topResultsList.push({
          title: item.displayName || 'Unknown user',
          summary: `Email: ${item.mail || item.userPrincipalName || 'N/A'}, Title: ${item.jobTitle || 'N/A'}`,
          relevance: 100 - topResultsList.length * 5,
          metadata: {
            mail: item.mail,
            userPrincipalName: item.userPrincipalName,
            jobTitle: item.jobTitle,
            department: item.department,
          },
        });
      }
    } else if (product === 'Groups') {
      // Search groups - use filter since $search may not be supported
      const result = await callGraph(graphClient, 'GET', '/groups', {
        $filter: `contains(displayName,'${query}') or contains(mail,'${query}') or contains(description,'${query}')`,
        $top: String(topResults),
        $select: 'id,displayName,mail,description',
      });
      const groups = JSON.parse(result);
      const items = groups.value || [];

      for (const item of items.slice(0, topResults)) {
        topResultsList.push({
          title: item.displayName || 'Unnamed group',
          summary: item.description || '',
          relevance: 100 - topResultsList.length * 5,
          metadata: {
            mail: item.mail,
            description: item.description,
          },
        });
      }
    } else if (product === 'Planner') {
      // Search Planner plans
      const result = await callGraph(graphClient, 'GET', '/me/planner/plans', {
        $top: String(topResults),
        $select: 'id,title,createdDateTime',
      });
      const plans = JSON.parse(result);
      const items = plans.value || [];

      // Filter plans that match the query
      const queryLower = query.toLowerCase();
      const matchingPlans = items.filter((plan: Record<string, unknown>) => {
        const title = (plan.title || '').toString().toLowerCase();
        return title.includes(queryLower);
      });

      for (const item of matchingPlans.slice(0, topResults)) {
        topResultsList.push({
          title: item.title || 'Untitled plan',
          summary: `Created: ${item.createdDateTime || 'N/A'}`,
          relevance: 100 - topResultsList.length * 5,
          metadata: {
            createdDateTime: item.createdDateTime,
          },
        });
      }
    } else if (product === 'ToDo') {
      // Search ToDo lists
      const result = await callGraph(graphClient, 'GET', '/me/todo/lists', {
        $top: String(topResults),
        $select: 'id,displayName,wellknownListName',
      });
      const lists = JSON.parse(result);
      const items = lists.value || [];

      // Filter lists that match the query
      const queryLower = query.toLowerCase();
      const matchingLists = items.filter((list: Record<string, unknown>) => {
        const displayName = (list.displayName || '').toString().toLowerCase();
        return displayName.includes(queryLower);
      });

      for (const item of matchingLists.slice(0, topResults)) {
        topResultsList.push({
          title: item.displayName || 'Unnamed list',
          summary: `List type: ${item.wellknownListName || 'custom'}`,
          relevance: 100 - topResultsList.length * 5,
          metadata: {
            wellknownListName: item.wellknownListName,
          },
        });
      }
    }
  } catch (error) {
    logger.error(`Error searching product ${product}: ${error}`);
    // Return empty result instead of throwing
  }

  return {
    product,
    resultCount: topResultsList.length,
    topResults: topResultsList,
  };
}

/**
 * Handle product-based search
 */
async function handleProductSearch(
  input: ProductSearchInput,
  graphClient: GraphClient,
  _readOnly: boolean
): Promise<string> {
  const thinking: string[] = [];
  const startTime = Date.now();

  thinking.push(`🔍 Product-based Search: "${input.query}"`);

  const maxResults = input.maxResults || 50;
  const topPerProduct = input.topPerProduct || 5;

  // Step 1: Microsoft 365 Search - Initial search with Top 50 results
  thinking.push(`📊 Step 1: Executing Microsoft 365 Search (max ${maxResults} results)`);

  const entityTypes: Array<
    'message' | 'event' | 'driveItem' | 'site' | 'list' | 'listItem' | 'chatMessage' | 'person'
  > = ['message', 'event', 'driveItem', 'site', 'list', 'listItem', 'chatMessage', 'person'];

  // Add time context for events
  const now = new Date();
  const defaultStartDate = new Date(
    now.getFullYear() - 2,
    now.getMonth(),
    now.getDate()
  ).toISOString();
  const defaultEndDate = new Date(
    now.getFullYear() + 1,
    now.getMonth(),
    now.getDate()
  ).toISOString();

  const searchRequest: Record<string, unknown> = {
    requests: [
      {
        entityTypes,
        query: {
          queryString: input.query,
        },
        from: 0,
        size: Math.min(maxResults, 500),
        trimDuplicates: true,
        // Add time context if events are included
        ...(entityTypes.includes('event') && {
          timeContext: {
            startDateTime: defaultStartDate,
            endDateTime: defaultEndDate,
          },
        }),
      },
    ],
  };

  let initialResults: Record<string, unknown[]> = {};
  let totalHits = 0;

  try {
    const result = await callGraph(graphClient, 'POST', '/search/query', undefined, searchRequest);
    const parsedResult = JSON.parse(result);

    // Extract and format results
    if (parsedResult.value && Array.isArray(parsedResult.value)) {
      for (const response of parsedResult.value) {
        if (response.hitsContainers && Array.isArray(response.hitsContainers)) {
          for (const container of response.hitsContainers) {
            totalHits += container.total || 0;
            if (container.hits && Array.isArray(container.hits)) {
              for (const hit of container.hits) {
                const entityType = hit.resource?.['@odata.type'] || 'unknown';
                if (!initialResults[entityType]) {
                  initialResults[entityType] = [];
                }
                initialResults[entityType].push({
                  id: hit.resource?.id,
                  summary: hit.summary,
                  rank: hit.rank,
                  ...hit.resource,
                });
              }
            }
          }
        }
      }
    }

    thinking.push(
      `✅ Found ${totalHits} initial results across ${Object.keys(initialResults).length} entity types`
    );
  } catch (error) {
    thinking.push(
      `⚠️ Initial search failed: ${error instanceof Error ? error.message : String(error)}`
    );
    thinking.push('Continuing with product-specific searches...');
  }

  // Step 2: Product Detection
  thinking.push(`📊 Step 2: Detecting products from search results`);
  const productsDetected = detectProductsFromSearchResults(initialResults);
  thinking.push(
    `✅ Detected products: ${productsDetected.length > 0 ? productsDetected.join(', ') : 'None'}`
  );

  // Step 3: Product-specific Queries
  thinking.push(`📊 Step 3: Executing product-specific searches`);
  const productResults: ProductSearchResult[] = [];

  if (productsDetected.length === 0) {
    thinking.push('⚠️ No products detected from initial search - trying all products');
    // If no products detected, try searching all products
    for (const mapping of PRODUCT_MAPPINGS) {
      if (mapping.requiresDirectApi) {
        try {
          const result = await searchProduct(
            mapping.product,
            input.query,
            graphClient,
            topPerProduct
          );
          if (result.resultCount > 0) {
            productResults.push(result);
            thinking.push(`✅ ${mapping.product}: Found ${result.resultCount} results`);
          }
        } catch (error) {
          thinking.push(
            `⚠️ ${mapping.product}: Search failed - ${error instanceof Error ? error.message : String(error)}`
          );
        }
      }
    }
  } else {
    // Search only detected products
    for (const product of productsDetected) {
      try {
        const result = await searchProduct(product, input.query, graphClient, topPerProduct);
        productResults.push(result);
        thinking.push(`✅ ${product}: Found ${result.resultCount} results`);
      } catch (error) {
        thinking.push(
          `⚠️ ${product}: Search failed - ${error instanceof Error ? error.message : String(error)}`
        );
      }
    }
  }

  // Step 4: Format Response
  thinking.push(`📊 Step 4: Formatting response`);
  const response: ProductSearchResponse = {
    query: input.query,
    initialSearchResults: {
      totalHits,
      productsDetected,
    },
    productResults,
    thinking,
  };

  const executionTime = Date.now() - startTime;
  thinking.push(`⏱️ Total execution time: ${executionTime}ms`);

  return addThinkingToResponse(JSON.stringify(response, null, 2), thinking);
}

// ============================================================================
// 11. ASSISTANT SUPER-TOOL (Smart/Compound Operations)
// ============================================================================
const assistantActions = z.enum([
  'ask', // Natural language question about M365 data
  'search', // Search across all M365 data
  'my-day', // Get today's summary
  'my-week', // Get week summary
  'person-info', // Get all info about a person
  'project-overview', // Get project overview
  'follow-ups', // Get pending follow-up items
  'meeting-prep', // Prepare for upcoming meeting
  // NEW: Discovery Actions with NLP
  'discover', // Automatic discovery - NLP analyzes query and chooses best strategy
  'discover-person', // Comprehensive person discovery (emails, meetings, chats, files)
  'discover-project', // Comprehensive project discovery (files, sites, tasks, meetings)
  'discover-topic', // Comprehensive topic search across all M365 products
  'discover-company', // Customer 360 - comprehensive company/organization view
]);

/**
 * Discovery Response Interface for structured output
 */
interface DiscoveryResponse {
  target: string;
  targetType: 'person' | 'project' | 'topic' | 'company';
  nlpAnalysis: {
    detectedIntent: string;
    detectedEntities: Array<{ value: string; type: string; confidence: number }>;
    confidence: number;
    suggestedFollowUps: string[];
  };
  summary: {
    totalItems: number;
    sources: string[];
    timeRange: string;
  };
  results: {
    emails?: unknown[];
    files?: unknown[];
    meetings?: unknown[];
    chats?: unknown[];
    tasks?: unknown[];
    contacts?: unknown[];
    sites?: unknown[];
    listItems?: unknown[];
    documentsWithContent?: Array<{ file: unknown; contentPreview: string | null }>;
  };
  insights?: {
    relationshipScore?: number;
    recommendations?: string[];
    recentActivity?: string;
    lastInteraction?: string;
    dataSummary?: string;
  };
  keyFindings?: {
    companyInfo?: string[];
    keyContacts?: Array<{ name?: string; email?: string; role?: string }>;
    recentTopics?: string[];
    importantFiles?: Array<{ name?: string; webUrl?: string; lastModified?: string }>;
    sites?: Array<{ name?: string; webUrl?: string }>;
  };
}

const assistantSchema = z.object({
  action: assistantActions.describe(
    'The assistant operation to perform. Use "discover" for automatic NLP-based discovery, or specific discover-* actions for targeted searches.'
  ),
  // Query
  query: z.string().optional().describe('Natural language query or search term'),
  // Target for discovery actions
  target: z
    .string()
    .optional()
    .describe('Target for discovery actions (person name, project name, topic, or company name)'),
  // Person context
  person: z.string().optional().describe('Person name or email'),
  // Project/topic context
  topic: z.string().optional().describe('Topic or project name'),
  // Time context
  days: z.number().optional().default(7).describe('Number of days to look back (max: 365)'),
  // Limits
  limit: z.number().optional().default(25).describe('Maximum results per category (max: 100)'),
  // Include download links for files
  includeDownloadLinks: z
    .boolean()
    .optional()
    .default(false)
    .describe('Include download links for discovered files'),
});

/** Input type for assistant tool; use z.input so partial objects (e.g. auto-execution) are accepted. */
type AssistantInput = z.input<typeof assistantSchema>;

async function handleAssistant(
  input: AssistantInput,
  graphClient: GraphClient,
  _readOnly: boolean
): Promise<string> {
  const thinking: string[] = [];
  const results: Record<string, unknown> = {};
  // Assistant operations are read-only (queries only)
  const limit = input.limit || 25;
  const days = input.days || 7;

  switch (input.action) {
    case 'ask': {
      if (!input.query) throw new Error('query is required for ask action');

      // Automatic query optimization for question processing
      const startTime = Date.now();
      const currentUserId = getUserId();
      const userIdHash = currentUserId ? getQueryStore().hashUserId(currentUserId) : undefined;
      const optimized = optimizeQueryForSearch(input.query, {
        tool: 'assistant-ask',
        userIdHash,
      });
      thinking.push(`💭 Processing question: "${input.query}"`);
      addOptimizationThinking(thinking, optimized);
      if (optimized.nlpAnalysis.intent) {
        thinking.push(
          `📊 Detected intent: ${optimized.nlpAnalysis.intent}, service: ${optimized.nlpAnalysis.service || 'general'}`
        );
      }
      thinking.push('Searching across emails, calendar, files, and chats...');

      // Parallel API calls with optimized query
      const [emailResult, filesResult] = await Promise.allSettled([
        callGraph(graphClient, 'GET', '/me/messages', {
          $search: formatSearchQuery(optimized.optimizedQuery),
          $top: String(Math.min(limit, 10)),
        }),
        callGraph(
          graphClient,
          'GET',
          `/me/drive/root/search(q='${encodeURIComponent(optimized.optimizedQuery)}')`
        ),
      ]);

      if (emailResult.status === 'fulfilled') {
        results.emails = JSON.parse(emailResult.value);
      } else {
        results.emailError = emailResult.reason?.message || 'Unknown error';
      }

      if (filesResult.status === 'fulfilled') {
        results.files = JSON.parse(filesResult.value);
      } else {
        results.filesError = filesResult.reason?.message || 'Unknown error';
      }

      const executionTime = Date.now() - startTime;

      // Record optimization result for learning
      const askSuccess =
        (results.emails as Record<string, unknown>)?.value !== undefined ||
        (results.files as Record<string, unknown>)?.value !== undefined;
      if (userIdHash) {
        recordQueryOptimizationResult(
          input.query,
          optimized.optimizedQuery,
          askSuccess,
          userIdHash,
          'assistant-ask'
        );
      }

      // Format response with metadata
      const responseWithMetadata = formatStandardResponse(results, {
        executionTime,
        sources: ['emails', 'files'].filter((s) => results[s]),
        cacheHit: false,
        nlpAnalysis: optimized.nlpAnalysis,
        errors: [
          ...(results.emailError
            ? [{ message: `Email search failed: ${results.emailError}`, retryable: true }]
            : []),
          ...(results.filesError
            ? [{ message: `File search failed: ${results.filesError}`, retryable: true }]
            : []),
        ],
        suggestions: [
          '💡 Use "search" tool for unified Microsoft 365 search',
          '💡 Use "assistant" tool with action "discover" for comprehensive discovery',
          '💡 Refine your query for better results',
        ],
      });

      return formatAndReturnToolResponse(responseWithMetadata, thinking);
    }

    case 'search': {
      if (!input.query) throw new Error('query is required for search action');

      // Automatic query optimization for comprehensive search
      const startTime = Date.now();
      const currentUserId = getUserId();
      const userIdHash = currentUserId ? getQueryStore().hashUserId(currentUserId) : undefined;
      const optimized = optimizeQueryForSearch(input.query, {
        tool: 'assistant-search',
        userIdHash,
      });
      thinking.push(`🔍 Searching everything for: "${input.query}"`);
      addOptimizationThinking(thinking, optimized);
      if (optimized.nlpAnalysis.intent) {
        thinking.push(
          `📊 Detected intent: ${optimized.nlpAnalysis.intent}, service: ${optimized.nlpAnalysis.service || 'general'}`
        );
      }

      // Parallel API calls with optimized query
      const [emailResult, filesResult] = await Promise.allSettled([
        callGraph(graphClient, 'GET', '/me/messages', {
          $search: formatSearchQuery(optimized.optimizedQuery),
          $top: String(limit),
        }),
        callGraph(
          graphClient,
          'GET',
          `/me/drive/root/search(q='${encodeURIComponent(optimized.optimizedQuery)}')`
        ),
      ]);

      if (emailResult.status === 'fulfilled') {
        results.emails = JSON.parse(emailResult.value);
      } else {
        results.emailError = emailResult.reason?.message || 'Unknown error';
      }

      if (filesResult.status === 'fulfilled') {
        results.files = JSON.parse(filesResult.value);
      } else {
        results.filesError = filesResult.reason?.message || 'Unknown error';
      }

      const executionTime = Date.now() - startTime;

      // Record optimization result for learning
      const searchSuccess =
        (results.emails as Record<string, unknown>)?.value !== undefined ||
        (results.files as Record<string, unknown>)?.value !== undefined;
      if (userIdHash) {
        recordQueryOptimizationResult(
          input.query,
          optimized.optimizedQuery,
          searchSuccess,
          userIdHash,
          'assistant-search'
        );
      }

      // Format response with metadata
      const responseWithMetadata = formatStandardResponse(results, {
        executionTime,
        sources: ['emails', 'files'].filter((s) => results[s]),
        cacheHit: false,
        nlpAnalysis: optimized.nlpAnalysis,
        errors: [
          ...(results.emailError
            ? [{ message: `Email search failed: ${results.emailError}`, retryable: true }]
            : []),
          ...(results.filesError
            ? [{ message: `File search failed: ${results.filesError}`, retryable: true }]
            : []),
        ],
        suggestions: [
          '💡 Use "search" tool for unified Microsoft 365 search',
          '💡 Use "assistant" tool with action "discover" for comprehensive discovery',
        ],
      });

      return formatAndReturnToolResponse(responseWithMetadata, thinking);
    }

    case 'my-day': {
      const startTime = Date.now();
      thinking.push("Getting today's summary");
      const today = setStartOfDay(new Date());
      const todayEnd = setEndOfDay(new Date());

      const batchRequests: GraphBatchRequest[] = [
        {
          id: 'calendar',
          method: 'GET',
          url: buildGraphBatchUrl('/me/calendarView', {
            startDateTime: today.toISOString(),
            endDateTime: todayEnd.toISOString(),
          }),
        },
        {
          id: 'messages',
          method: 'GET',
          url: buildGraphBatchUrl('/me/messages', {
            $top: '10',
            $orderby: 'receivedDateTime desc',
          }),
        },
      ];

      const errors: ErrorInfo[] = [];
      try {
        const batchResponses = await graphClient.performBatch(batchRequests);
        for (const res of batchResponses) {
          if (res.status >= 400) {
            const desc = res.id === 'calendar' ? "Today's calendar events" : 'Recent emails';
            errors.push({
              message: `${desc} failed: ${res.status} ${typeof res.body === 'object' && res.body && 'error' in res.body ? JSON.stringify((res.body as { error: unknown }).error) : ''}`,
              retryable: res.status === 429 || res.status === 503,
              details: res.body,
            });
          } else if (res.id === 'calendar' && res.body) {
            results.todayEvents = res.body as Record<string, unknown>;
          } else if (res.id === 'messages' && res.body) {
            results.recentEmails = res.body as Record<string, unknown>;
          }
        }
      } catch (err) {
        errors.push({
          message: err instanceof Error ? err.message : String(err),
          retryable: isRetryableError(err),
          details: err,
        });
      }

      const executionTime = Date.now() - startTime;

      // Format response with metadata
      const responseWithMetadata = formatStandardResponse(results, {
        executionTime,
        sources: ['calendar', 'email'].filter((s) => {
          if (s === 'calendar') return results.todayEvents;
          if (s === 'email') return results.recentEmails;
          return false;
        }),
        cacheHit: false,
        ...(errors.length > 0 && { errors }),
        suggestions: [
          '💡 Use "calendar" tool with action "view" for detailed calendar view',
          '💡 Use "email" tool with action "list" to see more emails',
          '💡 Use "assistant" tool with action "my-week" for week overview',
        ],
      });

      return formatAndReturnToolResponse(responseWithMetadata, thinking);
    }

    case 'my-week': {
      const startTime = Date.now();
      thinking.push('Getting week summary');
      const today = setStartOfDay(new Date());
      const weekEnd = new Date(today);
      weekEnd.setDate(weekEnd.getDate() + 7);
      const weekEndTime = setEndOfDay(weekEnd);

      const batchRequestsWeek: GraphBatchRequest[] = [
        {
          id: 'calendar',
          method: 'GET',
          url: buildGraphBatchUrl('/me/calendarView', {
            startDateTime: today.toISOString(),
            endDateTime: weekEndTime.toISOString(),
          }),
        },
        {
          id: 'todo',
          method: 'GET',
          url: '/me/todo/lists',
        },
      ];

      const errors: ErrorInfo[] = [];
      try {
        const batchResponses = await graphClient.performBatch(batchRequestsWeek);
        for (const res of batchResponses) {
          if (res.status >= 400) {
            const desc = res.id === 'calendar' ? 'Week calendar events' : 'To-Do lists';
            errors.push({
              message: `${desc} failed: ${res.status} ${typeof res.body === 'object' && res.body && 'error' in res.body ? JSON.stringify((res.body as { error: unknown }).error) : ''}`,
              retryable: res.status === 429 || res.status === 503,
              details: res.body,
            });
          } else if (res.id === 'calendar' && res.body) {
            results.weekEvents = res.body as Record<string, unknown>;
          } else if (res.id === 'todo' && res.body) {
            results.tasks = res.body as Record<string, unknown>;
          }
        }
      } catch (err) {
        errors.push({
          message: err instanceof Error ? err.message : String(err),
          retryable: isRetryableError(err),
          details: err,
        });
      }

      const executionTime = Date.now() - startTime;

      // Format response with metadata
      const responseWithMetadata = formatStandardResponse(results, {
        executionTime,
        sources: ['calendar', 'tasks'].filter((s) => {
          if (s === 'calendar') return results.weekEvents;
          if (s === 'tasks') return results.tasks;
          return false;
        }),
        cacheHit: false,
        ...(errors.length > 0 && { errors }),
        suggestions: [
          '💡 Use "calendar" tool with action "view" for detailed calendar view',
          '💡 Use "tasks" tool with action "todo-lists" to see task lists',
          '💡 Use "assistant" tool with action "my-day" for today\'s summary',
        ],
      });

      return formatAndReturnToolResponse(responseWithMetadata, thinking);
    }

    case 'person-info': {
      if (!input.person) throw new Error('person is required');
      thinking.push(`Getting all info about: ${input.person}`);

      const emailsResult = await callGraph(graphClient, 'GET', '/me/messages', {
        $filter: `from/emailAddress/address eq '${input.person}' or contains(from/emailAddress/name, '${input.person}')`,
        $top: String(limit),
      });
      results.emails = JSON.parse(emailsResult);

      return addThinkingToResponse(JSON.stringify(results, null, 2), thinking);
    }

    case 'project-overview': {
      if (!input.topic) throw new Error('topic is required');
      thinking.push(`Getting project overview for: ${input.topic}`);

      const emailsResult = await callGraph(graphClient, 'GET', '/me/messages', {
        $search: `"${input.topic}"`,
        $top: String(limit),
      });
      results.emails = JSON.parse(emailsResult);

      const filesResult = await callGraph(
        graphClient,
        'GET',
        `/me/drive/root/search(q='${input.topic}')`
      );
      results.files = JSON.parse(filesResult);

      return addThinkingToResponse(JSON.stringify(results, null, 2), thinking);
    }

    case 'follow-ups': {
      thinking.push('Getting pending follow-up items');

      const flaggedEmails = await callGraph(graphClient, 'GET', '/me/messages', {
        $filter: "flag/flagStatus eq 'flagged'",
        $top: String(limit),
      });
      results.flaggedEmails = JSON.parse(flaggedEmails);

      const tasks = await callGraph(graphClient, 'GET', '/me/todo/lists');
      results.tasks = JSON.parse(tasks);

      return addThinkingToResponse(JSON.stringify(results, null, 2), thinking);
    }

    case 'meeting-prep': {
      thinking.push('Preparing for upcoming meetings');
      const today = setStartOfDay(new Date());
      const todayEnd = setEndOfDay(new Date());

      const eventsResult = await callGraph(graphClient, 'GET', '/me/calendarView', {
        startDateTime: today.toISOString(),
        endDateTime: todayEnd.toISOString(),
      });
      results.upcomingMeetings = JSON.parse(eventsResult);

      return addThinkingToResponse(JSON.stringify(results, null, 2), thinking);
    }

    // =========================================================================
    // DISCOVERY ACTIONS - NLP-powered comprehensive search
    // =========================================================================

    case 'discover': {
      // Automatic discovery - NLP analyzes and routes to best strategy
      const queryText = input.query || input.target || input.topic || input.person;
      if (!queryText) {
        throw new Error(
          'Query, target, topic, or person is required for discover action. ' +
            'Example: { "action": "discover", "query": "Alle Infos zu Max Müller" }'
        );
      }

      thinking.push(`🔍 Analyzing query with NLP: "${queryText}"`);
      const decomposed = nlpEnhancer.decomposeQuery(queryText);

      thinking.push(`📊 NLP Analysis:`);
      thinking.push(
        `   Intent: ${decomposed.intent.type} (confidence: ${Math.round(decomposed.intent.confidence * 100)}%)`
      );
      thinking.push(
        `   Entities: ${decomposed.entities.map((e) => `${e.value} (${e.type})`).join(', ') || 'none detected'}`
      );
      if (decomposed.temporal) {
        thinking.push(
          `   Temporal: ${decomposed.temporal.expression} (${decomposed.temporal.type})`
        );
      }
      if (decomposed.ms365Context) {
        thinking.push(`   MS365 Context: ${decomposed.ms365Context.service}`);
      }

      // Route to best discovery action based on NLP analysis
      const primaryEntity = decomposed.entities[0];

      // Check for company/bank patterns (e.g., "DZ Bank", "Bank", "AG", "GmbH", etc.)
      const companyPatterns = /\b(bank|ag|gmbh|inc|corp|company|ltd|llc|gmbh|kg)\b/i;
      const isLikelyCompany =
        primaryEntity?.type === 'organization' ||
        companyPatterns.test(queryText) ||
        queryText.split(/\s+/).length <= 3; // Short queries are often company names

      if (primaryEntity?.type === 'person' || decomposed.intent.type === 'who') {
        thinking.push(`🎯 Routing to: discover-person (detected person entity)`);
        return handleDiscoverPerson(
          { ...input, target: primaryEntity?.value || queryText },
          graphClient,
          decomposed,
          thinking
        );
      } else if (primaryEntity?.type === 'organization' || isLikelyCompany) {
        thinking.push(
          `🎯 Routing to: discover-company (detected organization entity or company pattern)`
        );
        return handleDiscoverCompany(
          { ...input, target: primaryEntity?.value || queryText },
          graphClient,
          decomposed,
          thinking
        );
      } else if (primaryEntity?.type === 'project') {
        thinking.push(`🎯 Routing to: discover-project (detected project entity)`);
        return handleDiscoverProject(
          { ...input, target: primaryEntity?.value || queryText },
          graphClient,
          decomposed,
          thinking
        );
      } else {
        thinking.push(`🎯 Routing to: discover-topic (general topic search)`);
        return handleDiscoverTopic(
          { ...input, target: queryText },
          graphClient,
          decomposed,
          thinking
        );
      }
    }

    case 'discover-person': {
      const personName = input.target || input.person || input.query;
      if (!personName) {
        throw new Error(
          'Target, person, or query is required for discover-person action. ' +
            'Example: { "action": "discover-person", "target": "Max Müller" }'
        );
      }

      thinking.push(`👤 Discovering person: "${personName}"`);
      const decomposed = nlpEnhancer.decomposeQuery(personName);
      return handleDiscoverPerson(
        { ...input, target: personName },
        graphClient,
        decomposed,
        thinking
      );
    }

    case 'discover-project': {
      const projectName = input.target || input.topic || input.query;
      if (!projectName) {
        throw new Error(
          'Target, topic, or query is required for discover-project action. ' +
            'Example: { "action": "discover-project", "target": "Project Alpha" }'
        );
      }

      thinking.push(`📁 Discovering project: "${projectName}"`);
      const decomposed = nlpEnhancer.decomposeQuery(projectName);
      return handleDiscoverProject(
        { ...input, target: projectName },
        graphClient,
        decomposed,
        thinking
      );
    }

    case 'discover-topic': {
      const topicName = input.target || input.topic || input.query;
      if (!topicName) {
        throw new Error(
          'Target, topic, or query is required for discover-topic action. ' +
            'Example: { "action": "discover-topic", "target": "Budget 2024" }'
        );
      }

      thinking.push(`🔎 Discovering topic: "${topicName}"`);
      const decomposed = nlpEnhancer.decomposeQuery(topicName);
      return handleDiscoverTopic(
        { ...input, target: topicName },
        graphClient,
        decomposed,
        thinking
      );
    }

    case 'discover-company': {
      const companyName = input.target || input.query;
      if (!companyName) {
        throw new Error(
          'Target or query is required for discover-company action. ' +
            'Example: { "action": "discover-company", "target": "Acme Corp" }'
        );
      }

      thinking.push(`🏢 Discovering company: "${companyName}"`);
      const decomposed = nlpEnhancer.decomposeQuery(companyName);
      return handleDiscoverCompany(
        { ...input, target: companyName },
        graphClient,
        decomposed,
        thinking
      );
    }

    default:
      throw new Error(`Unknown assistant action: ${input.action}`);
  }
}

// ============================================================================
// DISCOVERY HELPER FUNCTIONS
// ============================================================================

/**
 * Discover comprehensive information about a person
 */
async function handleDiscoverPerson(
  input: AssistantInput & { target?: string },
  graphClient: GraphClient,
  decomposed: DecomposedQuery,
  thinking: string[]
): Promise<string> {
  const personName = input.target || '';
  const limit = Math.min(input.limit || 25, 100);
  const days = Math.min(input.days || 90, 365);

  let startDate = new Date();
  startDate.setDate(startDate.getDate() - days);
  startDate = setStartOfDay(startDate);
  let endDate = new Date();
  endDate.setDate(endDate.getDate() + 30); // Include future meetings
  endDate = setEndOfDay(endDate);

  thinking.push(`📊 Searching across multiple sources for: ${personName}`);
  thinking.push(`   Time range: last ${days} days`);

  // Parallel API calls for comprehensive person discovery
  const [emailsResult, eventsResult, filesResult, contactsResult, chatsResult] =
    await Promise.allSettled([
      // Emails from/to this person
      callGraph(graphClient, 'GET', '/me/messages', {
        $search: `"from:${personName}" OR "to:${personName}"`,
        $top: String(limit),
        $orderby: 'receivedDateTime desc',
      }),
      // Calendar events with this person
      callGraph(graphClient, 'GET', '/me/calendarView', {
        startDateTime: startDate.toISOString(),
        endDateTime: endDate.toISOString(),
        $top: String(limit),
      }),
      // Files shared/created by this person
      callGraph(graphClient, 'GET', `/me/drive/root/search(q='${encodeURIComponent(personName)}')`),
      // Contact information
      callGraph(graphClient, 'GET', '/me/contacts', {
        $search: `"${personName}"`,
        $top: '10',
      }),
      // Teams chats (if available)
      callGraph(graphClient, 'GET', '/me/chats', {
        $top: '25',
      }),
    ]);

  // Process results
  const emails =
    emailsResult.status === 'fulfilled' ? JSON.parse(emailsResult.value) : { value: [] };
  const events =
    eventsResult.status === 'fulfilled' ? JSON.parse(eventsResult.value) : { value: [] };
  const files = filesResult.status === 'fulfilled' ? JSON.parse(filesResult.value) : { value: [] };
  const contacts =
    contactsResult.status === 'fulfilled' ? JSON.parse(contactsResult.value) : { value: [] };
  const chats = chatsResult.status === 'fulfilled' ? JSON.parse(chatsResult.value) : { value: [] };

  // Filter events that include this person - person must be an actual attendee or organizer
  const personLower = personName.toLowerCase();
  // Split name into parts for better matching (e.g., "Max Mustermann" -> ["max", "mustermann"])
  const nameParts = personLower.split(/\s+/).filter((part) => part.length > 0);

  const relevantEvents = (events.value || []).filter((event: Record<string, unknown>) => {
    const attendees = event.attendees as
      | Array<{ emailAddress?: { name?: string; address?: string } }>
      | undefined;
    const organizer = event.organizer as
      | { emailAddress?: { name?: string; address?: string } }
      | undefined;

    // Check organizer - must be exact name match OR all name parts present
    if (organizer?.emailAddress?.name) {
      const organizerNameLower = organizer.emailAddress.name.toLowerCase();
      // Exact name match
      if (organizerNameLower === personLower) {
        return true;
      }
      // Check if all name parts are present in organizer name
      if (nameParts.length > 0) {
        const allPartsMatch = nameParts.every((part) => organizerNameLower.includes(part));
        if (allPartsMatch) {
          return true;
        }
      }
    }

    // Check attendees - person must be an actual attendee
    if (attendees && Array.isArray(attendees)) {
      const hasAttendee = attendees.some((attendee) => {
        const attendeeName = attendee.emailAddress?.name?.toLowerCase();
        const attendeeEmail = attendee.emailAddress?.address?.toLowerCase();

        // Exact email match (if we had email, but we only have name)
        // Exact name match
        if (attendeeName === personLower) {
          return true;
        }

        // Check if all name parts are present in attendee name
        if (attendeeName && nameParts.length > 0) {
          const allPartsMatch = nameParts.every((part) => attendeeName.includes(part));
          if (allPartsMatch) {
            return true;
          }
        }

        return false;
      });

      if (hasAttendee) {
        return true;
      }
    }

    return false;
  });

  // Use DataAggregator for consistent deduplication and sorting
  const aggregated = dataAggregator.aggregate(
    [
      { source: 'emails', items: emails.value || [] },
      { source: 'calendar', items: relevantEvents },
      { source: 'files', items: files.value || [] },
      { source: 'contacts', items: contacts.value || [] },
      { source: 'chats', items: chats.value || [] },
    ],
    {
      sortBy: 'timestamp',
      sortOrder: 'desc',
      maxItems: limit * 2, // Get more items for better aggregation
      deduplicate: true,
    }
  );

  // Categorize aggregated items by source
  const categorizedResults: Record<string, unknown[]> = {
    emails: [],
    meetings: [],
    files: [],
    contacts: [],
    chats: [],
  };

  for (const item of aggregated.items) {
    const data = item.data as Record<string, unknown>;
    if (item.source === 'emails') {
      categorizedResults.emails.push(data);
    } else if (item.source === 'calendar') {
      categorizedResults.meetings.push(data);
    } else if (item.source === 'files') {
      categorizedResults.files.push(data);
    } else if (item.source === 'contacts') {
      categorizedResults.contacts.push(data);
    } else if (item.source === 'chats') {
      categorizedResults.chats.push(data);
    }
  }

  // Build discovery response
  const response: DiscoveryResponse = {
    target: personName,
    targetType: 'person',
    nlpAnalysis: {
      detectedIntent: decomposed.intent.type,
      detectedEntities: decomposed.entities.map((e) => ({
        value: e.value,
        type: e.type,
        confidence: e.confidence,
      })),
      confidence: decomposed.confidence,
      suggestedFollowUps: [
        `Show me recent emails from ${personName}`,
        `What meetings do I have with ${personName}?`,
        `Find files shared by ${personName}`,
      ],
    },
    summary: {
      totalItems: aggregated.uniqueItems,
      sources: aggregated.sources || [],
      timeRange: `Last ${days} days`,
    },
    results: {
      emails: categorizedResults.emails.slice(0, limit),
      meetings: categorizedResults.meetings.slice(0, limit),
      files: categorizedResults.files.slice(0, limit),
      contacts: categorizedResults.contacts,
      chats: categorizedResults.chats.slice(0, limit),
    },
    insights: {
      lastInteraction:
        (categorizedResults.emails[0] as { receivedDateTime?: string } | undefined)
          ?.receivedDateTime ||
        (categorizedResults.meetings[0] as { start?: { dateTime?: string } } | undefined)?.start
          ?.dateTime ||
        'Unknown',
      recentActivity: `${categorizedResults.emails.length} emails, ${categorizedResults.meetings.length} meetings`,
      recommendations: generatePersonRecommendations(
        categorizedResults.emails,
        categorizedResults.meetings,
        personName
      ),
    },
  };

  thinking.push(`✅ Discovery complete:`);
  thinking.push(`   📧 Emails: ${emails.value?.length || 0}`);
  thinking.push(`   📅 Meetings: ${relevantEvents.length}`);
  thinking.push(`   📁 Files: ${files.value?.length || 0}`);
  thinking.push(`   👤 Contacts: ${contacts.value?.length || 0}`);

  return addThinkingToResponse(JSON.stringify(response, null, 2), thinking);
}

/**
 * Discover comprehensive information about a project
 */
async function handleDiscoverProject(
  input: AssistantInput & { target?: string },
  graphClient: GraphClient,
  decomposed: DecomposedQuery,
  thinking: string[]
): Promise<string> {
  const projectName = input.target || '';
  const limit = Math.min(input.limit || 25, 100);
  const days = Math.min(input.days || 90, 365);

  const startDate = new Date();
  startDate.setDate(startDate.getDate() - days);

  thinking.push(`📊 Searching project resources for: ${projectName}`);

  // Parallel API calls for comprehensive project discovery
  const [emailsResult, filesResult, sitesResult, eventsResult, tasksResult] =
    await Promise.allSettled([
      // Emails mentioning this project
      callGraph(graphClient, 'GET', '/me/messages', {
        $search: `"${projectName}"`,
        $top: String(limit),
        $orderby: 'receivedDateTime desc',
      }),
      // Files related to this project
      callGraph(
        graphClient,
        'GET',
        `/me/drive/root/search(q='${encodeURIComponent(projectName)}')`
      ),
      // SharePoint sites
      callGraph(graphClient, 'GET', '/sites', {
        search: projectName,
        $top: '10',
      }),
      // Related meetings
      callGraph(graphClient, 'GET', '/me/events', {
        $search: `"${projectName}"`,
        $top: String(limit),
      }),
      // To-Do tasks
      callGraph(graphClient, 'GET', '/me/todo/lists'),
    ]);

  // Process results
  const emails =
    emailsResult.status === 'fulfilled' ? JSON.parse(emailsResult.value) : { value: [] };
  const files = filesResult.status === 'fulfilled' ? JSON.parse(filesResult.value) : { value: [] };
  const sites = sitesResult.status === 'fulfilled' ? JSON.parse(sitesResult.value) : { value: [] };
  const events =
    eventsResult.status === 'fulfilled' ? JSON.parse(eventsResult.value) : { value: [] };
  const tasks = tasksResult.status === 'fulfilled' ? JSON.parse(tasksResult.value) : { value: [] };

  // Use DataAggregator for consistent deduplication and sorting
  const aggregated = dataAggregator.aggregate(
    [
      { source: 'emails', items: emails.value || [] },
      { source: 'files', items: files.value || [] },
      { source: 'sites', items: sites.value || [] },
      { source: 'calendar', items: events.value || [] },
      { source: 'tasks', items: tasks.value || [] },
    ],
    {
      sortBy: 'timestamp',
      sortOrder: 'desc',
      maxItems: limit * 2,
      deduplicate: true,
    }
  );

  // Categorize aggregated items by source
  const categorizedResults: Record<string, unknown[]> = {
    emails: [],
    files: [],
    sites: [],
    meetings: [],
    tasks: [],
  };

  for (const item of aggregated.items) {
    const data = item.data as Record<string, unknown>;
    if (item.source === 'emails') {
      categorizedResults.emails.push(data);
    } else if (item.source === 'files') {
      categorizedResults.files.push(data);
    } else if (item.source === 'sites') {
      categorizedResults.sites.push(data);
    } else if (item.source === 'calendar') {
      categorizedResults.meetings.push(data);
    } else if (item.source === 'tasks') {
      categorizedResults.tasks.push(data);
    }
  }

  // Build discovery response
  const response: DiscoveryResponse = {
    target: projectName,
    targetType: 'project',
    nlpAnalysis: {
      detectedIntent: decomposed.intent.type,
      detectedEntities: decomposed.entities.map((e) => ({
        value: e.value,
        type: e.type,
        confidence: e.confidence,
      })),
      confidence: decomposed.confidence,
      suggestedFollowUps: [
        `Show me all files for ${projectName}`,
        `What are the upcoming meetings for ${projectName}?`,
        `List open tasks for ${projectName}`,
      ],
    },
    summary: {
      totalItems: aggregated.uniqueItems,
      sources: aggregated.sources || [],
      timeRange: `Last ${days} days`,
    },
    results: {
      emails: categorizedResults.emails.slice(0, limit),
      files: categorizedResults.files.slice(0, limit),
      sites: categorizedResults.sites,
      meetings: categorizedResults.meetings.slice(0, limit),
      tasks: categorizedResults.tasks,
    },
    insights: {
      recentActivity: `${categorizedResults.emails.length} emails, ${categorizedResults.files.length} files, ${categorizedResults.meetings.length} meetings`,
      recommendations: [
        categorizedResults.files.length > 0
          ? `📁 ${categorizedResults.files.length} project files found`
          : null,
        categorizedResults.sites.length > 0
          ? `🌐 ${categorizedResults.sites.length} SharePoint sites related`
          : null,
        categorizedResults.meetings.length > 0
          ? `📅 ${categorizedResults.meetings.length} meetings scheduled`
          : null,
      ].filter(Boolean) as string[],
    },
  };

  thinking.push(`✅ Project discovery complete:`);
  thinking.push(`   📧 Emails: ${emails.value?.length || 0}`);
  thinking.push(`   📁 Files: ${files.value?.length || 0}`);
  thinking.push(`   🌐 Sites: ${sites.value?.length || 0}`);
  thinking.push(`   📅 Meetings: ${events.value?.length || 0}`);

  return addThinkingToResponse(JSON.stringify(response, null, 2), thinking);
}

/**
 * Extract key findings from topic discovery results
 */
function extractTopicKeyFindings(
  topicName: string,
  categorizedResults: Record<string, unknown[]>
): {
  importantFiles?: Array<{ name?: string; webUrl?: string; lastModified?: string }>;
  recentTopics?: string[];
  keyContacts?: Array<{ name?: string; email?: string }>;
  sites?: Array<{ name?: string; webUrl?: string }>;
} {
  const keyFindings: {
    importantFiles?: Array<{ name?: string; webUrl?: string; lastModified?: string }>;
    recentTopics?: string[];
    keyContacts?: Array<{ name?: string; email?: string }>;
    sites?: Array<{ name?: string; webUrl?: string }>;
  } = {};

  // Extract important files
  const importantFiles: Array<{ name?: string; webUrl?: string; lastModified?: string }> = [];
  for (const file of categorizedResults.files.slice(0, 10)) {
    const fileObj = file as Record<string, unknown>;
    const name = fileObj.name as string | undefined;
    const webUrl = fileObj.webUrl as string | undefined;
    const lastModified = fileObj.lastModifiedDateTime as string | undefined;

    if (name) {
      importantFiles.push({
        name,
        webUrl,
        lastModified,
      });
    }
  }
  if (importantFiles.length > 0) {
    keyFindings.importantFiles = importantFiles;
  }

  // Extract recent topics from email subjects
  const topicsSet = new Set<string>();
  for (const email of categorizedResults.emails.slice(0, 15)) {
    const emailObj = email as Record<string, unknown>;
    const subject = emailObj.subject as string | undefined;
    if (subject && subject.length > 0) {
      // Remove common prefixes
      const cleanSubject = subject.replace(/^(RE:|AW:|FWD:|FW:)\s*/i, '').trim();
      if (cleanSubject.length > 5 && cleanSubject.length < 100) {
        topicsSet.add(cleanSubject);
      }
    }
  }
  if (topicsSet.size > 0) {
    keyFindings.recentTopics = Array.from(topicsSet).slice(0, 10);
  }

  // Extract key contacts from emails
  const contactsMap = new Map<string, { name?: string; email?: string }>();
  for (const email of categorizedResults.emails.slice(0, 20)) {
    const emailObj = email as Record<string, unknown>;
    const from = emailObj.from as Record<string, unknown> | undefined;

    if (from) {
      const emailAddress = from.emailAddress as Record<string, unknown> | undefined;
      const email = emailAddress?.address as string | undefined;
      const name = emailAddress?.name as string | undefined;
      if (email && !contactsMap.has(email)) {
        contactsMap.set(email, {
          name: name || email.split('@')[0],
          email,
        });
      }
    }
  }
  if (contactsMap.size > 0) {
    keyFindings.keyContacts = Array.from(contactsMap.values()).slice(0, 10);
  }

  // Extract sites
  const siteFindings: Array<{ name?: string; webUrl?: string }> = [];
  for (const site of categorizedResults.sites) {
    const siteObj = site as Record<string, unknown>;
    const name = siteObj.displayName as string | undefined;
    const webUrl = siteObj.webUrl as string | undefined;

    if (name || webUrl) {
      siteFindings.push({
        name: name || webUrl,
        webUrl,
      });
    }
  }
  if (siteFindings.length > 0) {
    keyFindings.sites = siteFindings;
  }

  return Object.keys(keyFindings).length > 0 ? keyFindings : {};
}

/**
 * Discover comprehensive information about a topic
 */
async function handleDiscoverTopic(
  input: AssistantInput & { target?: string },
  graphClient: GraphClient,
  decomposed: DecomposedQuery,
  thinking: string[]
): Promise<string> {
  const topicName = input.target || '';
  const limit = Math.min(input.limit || 25, 100);
  const days = Math.min(input.days || 90, 365);

  thinking.push(`📊 Comprehensive topic search for: ${topicName}`);

  // Use Microsoft Search API for unified search
  const searchRequest = {
    requests: [
      {
        entityTypes: ['message', 'event', 'driveItem', 'site', 'listItem', 'chatMessage'],
        query: {
          queryString: topicName,
        },
        from: 0,
        size: limit,
      },
    ],
  };

  // Parallel API calls
  const [searchResult, filesResult, eventsResult] = await Promise.allSettled([
    // Microsoft Search API
    callGraph(graphClient, 'POST', '/search/query', undefined, searchRequest),
    // Direct file search
    callGraph(graphClient, 'GET', `/me/drive/root/search(q='${encodeURIComponent(topicName)}')`),
    // Calendar events
    callGraph(graphClient, 'GET', '/me/events', {
      $search: `"${topicName}"`,
      $top: String(limit),
    }),
  ]);

  // Process search results
  let searchItems: unknown[] = [];
  let totalHits = 0;

  if (searchResult.status === 'fulfilled') {
    const parsed = JSON.parse(searchResult.value);
    if (parsed.value && Array.isArray(parsed.value)) {
      for (const response of parsed.value) {
        if (response.hitsContainers && Array.isArray(response.hitsContainers)) {
          for (const container of response.hitsContainers) {
            totalHits += container.total || 0;
            if (container.hits) {
              searchItems.push(...container.hits);
            }
          }
        }
      }
    }
  }

  const files = filesResult.status === 'fulfilled' ? JSON.parse(filesResult.value) : { value: [] };
  const events =
    eventsResult.status === 'fulfilled' ? JSON.parse(eventsResult.value) : { value: [] };

  // Extract resources from search hits
  const searchResources: unknown[] = [];
  for (const hit of searchItems as Array<{ resource?: Record<string, unknown> }>) {
    if (hit.resource) {
      searchResources.push(hit.resource);
    }
  }

  // Store raw counts BEFORE aggregation for accurate statistics
  const searchEmails = searchResources.filter((r: unknown) => {
    const res = r as Record<string, unknown>;
    return (res['@odata.type'] as string)?.includes('message');
  });
  const searchFiles = searchResources.filter((r: unknown) => {
    const res = r as Record<string, unknown>;
    return (res['@odata.type'] as string)?.includes('driveItem');
  });
  const searchMeetings = searchResources.filter((r: unknown) => {
    const res = r as Record<string, unknown>;
    return (res['@odata.type'] as string)?.includes('event');
  });
  const searchChats = searchResources.filter((r: unknown) => {
    const res = r as Record<string, unknown>;
    return (res['@odata.type'] as string)?.includes('chatMessage');
  });
  const searchSites = searchResources.filter((r: unknown) => {
    const res = r as Record<string, unknown>;
    return (res['@odata.type'] as string)?.includes('site');
  });

  const rawCounts = {
    emails: searchEmails.length + (events.value?.length || 0), // Include calendar events that might be emails
    files: searchFiles.length + (files.value?.length || 0),
    meetings: searchMeetings.length + (events.value?.length || 0),
    chats: searchChats.length,
    sites: searchSites.length,
  };

  // Use DataAggregator for consistent deduplication and sorting
  const aggregated = dataAggregator.aggregate(
    [
      {
        source: 'search-emails',
        items: searchResources.filter((r: unknown) => {
          const res = r as Record<string, unknown>;
          return (res['@odata.type'] as string)?.includes('message');
        }),
      },
      {
        source: 'search-files',
        items: searchResources.filter((r: unknown) => {
          const res = r as Record<string, unknown>;
          return (res['@odata.type'] as string)?.includes('driveItem');
        }),
      },
      {
        source: 'search-meetings',
        items: searchResources.filter((r: unknown) => {
          const res = r as Record<string, unknown>;
          return (res['@odata.type'] as string)?.includes('event');
        }),
      },
      {
        source: 'search-chats',
        items: searchResources.filter((r: unknown) => {
          const res = r as Record<string, unknown>;
          return (res['@odata.type'] as string)?.includes('chatMessage');
        }),
      },
      {
        source: 'search-sites',
        items: searchResources.filter((r: unknown) => {
          const res = r as Record<string, unknown>;
          return (res['@odata.type'] as string)?.includes('site');
        }),
      },
      { source: 'files', items: files.value || [] },
      { source: 'calendar', items: events.value || [] },
    ],
    {
      sortBy: 'timestamp',
      sortOrder: 'desc',
      maxItems: limit * 3, // More items for better aggregation
      deduplicate: true,
    }
  );

  // Categorize aggregated results by source
  const categorizedResults: Record<string, unknown[]> = {
    emails: [],
    files: [],
    meetings: [],
    chats: [],
    sites: [],
  };

  for (const item of aggregated.items) {
    const data = item.data as Record<string, unknown>;
    if (item.source.includes('email') || item.source.includes('message')) {
      categorizedResults.emails.push(data);
    } else if (item.source.includes('file') || item.source.includes('driveItem')) {
      categorizedResults.files.push(data);
    } else if (
      item.source.includes('meeting') ||
      item.source.includes('event') ||
      item.source === 'calendar'
    ) {
      categorizedResults.meetings.push(data);
    } else if (item.source.includes('chat')) {
      categorizedResults.chats.push(data);
    } else if (item.source.includes('site')) {
      categorizedResults.sites.push(data);
    }
  }

  // Build discovery response
  const response: DiscoveryResponse = {
    target: topicName,
    targetType: 'topic',
    nlpAnalysis: {
      detectedIntent: decomposed.intent.type,
      detectedEntities: decomposed.entities.map((e) => ({
        value: e.value,
        type: e.type,
        confidence: e.confidence,
      })),
      confidence: decomposed.confidence,
      suggestedFollowUps: [
        `Show me recent emails about ${topicName}`,
        `Find all files related to ${topicName}`,
        `What meetings discussed ${topicName}?`,
      ],
    },
    summary: {
      totalItems:
        rawCounts.emails + rawCounts.files + rawCounts.meetings + rawCounts.chats + rawCounts.sites,
      sources: aggregated.sources || [],
      timeRange: `Last ${days} days`,
    },
    results: {
      emails: categorizedResults.emails.slice(0, limit),
      files: categorizedResults.files.slice(0, limit),
      meetings: categorizedResults.meetings.slice(0, limit),
      chats: categorizedResults.chats.slice(0, limit),
      sites: categorizedResults.sites.slice(0, limit),
    },
    insights: {
      recentActivity: `Found ${rawCounts.emails} emails, ${rawCounts.files} files, ${rawCounts.meetings} meetings, ${rawCounts.chats} chats, ${rawCounts.sites} sites related to ${topicName}`,
      recommendations: [
        `📧 ${rawCounts.emails} emails found`,
        `📁 ${rawCounts.files} files found`,
        `📅 ${rawCounts.meetings} calendar items found`,
        `💬 ${rawCounts.chats} chat messages found`,
        `🌐 ${rawCounts.sites} sites found`,
      ],
      dataSummary: `IMPORTANT: Found ${rawCounts.emails} emails, ${rawCounts.files} files, ${rawCounts.meetings} meetings, ${rawCounts.chats} chats, and ${rawCounts.sites} sites related to "${topicName}". Use the data in 'results' and 'keyFindings' sections to answer questions. Do NOT say you found nothing - use the provided data.`,
    },
    keyFindings: extractTopicKeyFindings(topicName, categorizedResults),
  };

  thinking.push(`✅ Topic discovery complete:`);
  thinking.push(`   📊 Total hits: ${totalHits}, Unique items: ${aggregated.uniqueItems}`);
  thinking.push(
    `   📧 Emails: ${rawCounts.emails} (${categorizedResults.emails.length} in results)`
  );
  thinking.push(`   📁 Files: ${rawCounts.files} (${categorizedResults.files.length} in results)`);
  thinking.push(
    `   📅 Meetings: ${rawCounts.meetings} (${categorizedResults.meetings.length} in results)`
  );
  thinking.push(`   💬 Chats: ${rawCounts.chats} (${categorizedResults.chats.length} in results)`);
  thinking.push(`   🌐 Sites: ${rawCounts.sites} (${categorizedResults.sites.length} in results)`);

  return addThinkingToResponse(JSON.stringify(response, null, 2), thinking);
}

/**
 * Discover comprehensive information about a company (Customer 360)
 */
async function handleDiscoverCompany(
  input: AssistantInput & { target?: string },
  graphClient: GraphClient,
  decomposed: DecomposedQuery,
  thinking: string[]
): Promise<string> {
  const companyName = input.target || '';
  const limit = Math.min(input.limit || 25, 100);
  const days = Math.min(input.days || 90, 365);

  const startDate = new Date();
  startDate.setDate(startDate.getDate() - days);

  thinking.push(`📊 Customer 360 discovery for: ${companyName}`);

  // Microsoft Search API has strict rules about entity type combinations
  // Split into compatible groups to avoid "Invalid entity type combination" errors
  // Group 1: File types (can combine with each other)
  const fileTypes = validateEntityTypeCombinations(['driveItem', 'site', 'listItem']);
  // Group 2: Message types (can combine with each other)
  const messageTypes = validateEntityTypeCombinations(['message']);
  // Group 3: Standalone types (must be searched separately)
  // Note: 'event' and 'person' cannot be combined with other types

  // Build separate search requests for compatible groups
  const searchRequests = [];

  if (fileTypes.length > 0) {
    searchRequests.push({
      entityTypes: fileTypes,
      query: { queryString: companyName },
      from: 0,
      size: limit * 2,
    });
  }

  if (messageTypes.length > 0) {
    searchRequests.push({
      entityTypes: messageTypes,
      query: { queryString: companyName },
      from: 0,
      size: limit * 2,
    });
  }

  // Generate search variants for better coverage (e.g., "DZBANK", "DZ Bank", "DZ-Bank")
  const searchVariants = [
    companyName,
    companyName.replace(/([A-Z])([A-Z])/g, '$1 $2'), // "DZBANK" -> "DZ BANK"
    companyName.replace(/([A-Z])([A-Z])/g, '$1-$2'), // "DZBANK" -> "DZ-BANK"
  ].filter((v, i, arr) => arr.indexOf(v) === i); // Remove duplicates

  thinking.push(`🔍 Using search variants: ${searchVariants.join(', ')}`);
  thinking.push(
    `📋 Split entity types into compatible groups: fileTypes=[${fileTypes.join(', ')}], messageTypes=[${messageTypes.join(', ')}]`
  );

  // Try multiple search approaches for better coverage
  // 1. Microsoft Search API (unified search) - split by compatible groups
  // 2. Direct API searches with variants (throttled)
  // 3. Event searches (fetch and filter client-side, no $search support)

  // Execute Microsoft Search API first (if we have requests)
  let searchApiResults: unknown[] = [];
  if (searchRequests.length > 0) {
    try {
      const searchApiResponse = await callGraph(graphClient, 'POST', '/search/query', undefined, {
        requests: searchRequests,
      });
      const parsed = JSON.parse(searchApiResponse);
      if (parsed.value && Array.isArray(parsed.value)) {
        for (const response of parsed.value) {
          if (response.hitsContainers && Array.isArray(response.hitsContainers)) {
            for (const container of response.hitsContainers) {
              if (container.hits) {
                for (const hit of container.hits) {
                  if (hit.resource) {
                    searchApiResults.push(hit.resource);
                  }
                }
              }
            }
          }
        }
      }
    } catch (err) {
      logger.warn(`Microsoft Search API failed for "${companyName}": ${err}`);
    }
  }

  // Throttle direct searches to avoid rate limits
  // Process variants sequentially with small delays instead of all at once
  const emailResults: unknown[] = [];
  const eventResults: unknown[] = [];

  for (const variant of searchVariants) {
    // Email searches - try multiple approaches with throttling
    try {
      const emailResponse1 = await callGraph(graphClient, 'GET', '/me/messages', {
        $search: formatSearchQuery(variant, 'displayName', 'email'),
        $top: String(limit),
      });
      const emails1 = JSON.parse(emailResponse1);
      if (emails1.value) emailResults.push(...emails1.value);
    } catch (err) {
      logger.debug(`Email search variant "${variant}" failed: ${err}`);
    }

    // Small delay to avoid rate limits
    await new Promise((resolve) => setTimeout(resolve, 200));

    try {
      const emailResponse2 = await callGraph(graphClient, 'GET', '/me/messages', {
        $search: `"from:${variant}"`,
        $top: String(Math.floor(limit / 2)),
      });
      const emails2 = JSON.parse(emailResponse2);
      if (emails2.value) emailResults.push(...emails2.value);
    } catch (err) {
      logger.debug(`Email from search variant "${variant}" failed: ${err}`);
    }

    await new Promise((resolve) => setTimeout(resolve, 200));

    try {
      const emailResponse3 = await callGraph(graphClient, 'GET', '/me/messages', {
        $search: `"subject:${variant}"`,
        $top: String(Math.floor(limit / 2)),
      });
      const emails3 = JSON.parse(emailResponse3);
      if (emails3.value) emailResults.push(...emails3.value);
    } catch (err) {
      logger.debug(`Email subject search variant "${variant}" failed: ${err}`);
    }

    // Event searches - fetch events once and filter for all variants (more efficient)
    // Only fetch once for the first variant, then reuse for others
    if (variant === searchVariants[0]) {
      try {
        const eventsResponse = await callGraph(graphClient, 'GET', '/me/events', {
          $top: String(limit * 3), // Get more to filter for all variants
          $orderby: 'start/dateTime desc',
          $select: 'subject,body,organizer,attendees,start,end,location',
        });
        const events = JSON.parse(eventsResponse);
        if (events.value && Array.isArray(events.value)) {
          // Filter events that contain any variant
          const allVariantsLower = searchVariants.map((v) => v.toLowerCase());
          const filtered = events.value.filter((event: Record<string, unknown>) => {
            const subject = ((event.subject as string) || '').toLowerCase();
            const bodyObj = event.body as Record<string, unknown> | undefined;
            const body = ((bodyObj?.content as string) || '').toLowerCase();
            const locationObj = event.location as Record<string, unknown> | undefined;
            const location = ((locationObj?.displayName as string) || '').toLowerCase();
            return allVariantsLower.some(
              (variantLower) =>
                subject.includes(variantLower) ||
                body.includes(variantLower) ||
                location.includes(variantLower)
            );
          });
          eventResults.push(...filtered.slice(0, limit));
        }
      } catch (err) {
        logger.debug(`Event search failed: ${err}`);
      }
    }

    await new Promise((resolve) => setTimeout(resolve, 200));
  }

  // Single searches for files, contacts, sites (don't need variants)
  let filesResults: unknown[] = [];
  let contactsResults: unknown[] = [];
  let sitesResults: unknown[] = [];

  try {
    const filesResponse = await callGraph(
      graphClient,
      'GET',
      `/me/drive/root/search(q='${encodeURIComponent(companyName)}')`
    );
    const files = JSON.parse(filesResponse);
    if (files.value) filesResults = files.value;
  } catch (err) {
    logger.warn(`File search failed for "${companyName}": ${err}`);
  }

  try {
    const contactsResponse = await callGraph(graphClient, 'GET', '/me/contacts', {
      $search: formatSearchQuery(companyName, 'companyName', 'contact'),
      $top: '50',
    });
    const contacts = JSON.parse(contactsResponse);
    if (contacts.value) contactsResults = contacts.value;
  } catch (err) {
    logger.warn(`Contact search failed for "${companyName}": ${err}`);
  }

  try {
    const sitesResponse = await callGraph(graphClient, 'GET', '/sites', {
      search: companyName,
      $top: '10',
    });
    const sites = JSON.parse(sitesResponse);
    if (sites.value) sitesResults = sites.value;
  } catch (err) {
    logger.warn(`Site search failed for "${companyName}": ${err}`);
  }

  // Process Microsoft Search API results
  const searchEmails: unknown[] = [];
  const searchEvents: unknown[] = [];
  const searchFiles: unknown[] = [];
  const searchContacts: unknown[] = [];
  const searchSites: unknown[] = [];
  const searchListItems: unknown[] = [];

  // Categorize search API results
  for (const resource of searchApiResults) {
    const resourceObj = resource as Record<string, unknown>;
    const odataType = resourceObj['@odata.type'] as string | undefined;
    if (odataType?.includes('message')) {
      searchEmails.push(resourceObj);
    } else if (odataType?.includes('event')) {
      searchEvents.push(resourceObj);
    } else if (odataType?.includes('driveItem')) {
      searchFiles.push(resourceObj);
    } else if (odataType?.includes('person')) {
      searchContacts.push(resourceObj);
    } else if (odataType?.includes('site')) {
      searchSites.push(resourceObj);
    } else if (odataType?.includes('listItem')) {
      searchListItems.push(resourceObj);
    }
  }

  // Search SharePoint list items for found sites
  thinking.push(`🔍 Searching SharePoint list items for ${searchSites.length} sites`);
  const siteListItemsPromises = searchSites.slice(0, 5).map(async (site) => {
    // Limit to first 5 sites to avoid too many API calls
    const siteObj = site as Record<string, unknown>;
    const siteId = siteObj.id as string | undefined;
    if (!siteId) return [];

    try {
      // Get lists for this site
      const listsResponse = await callGraph(graphClient, 'GET', `/sites/${siteId}/lists`, {
        $top: '10',
      }).catch(() => JSON.stringify({ value: [] }));

      const lists = JSON.parse(listsResponse);
      if (!lists.value || !Array.isArray(lists.value)) return [];

      // Search list items for each list
      const listItemPromises = lists.value
        .slice(0, 5)
        .map(async (list: Record<string, unknown>) => {
          const listId = list.id as string | undefined;
          if (!listId) return [];

          try {
            const itemsResponse = await callGraph(
              graphClient,
              'GET',
              `/sites/${siteId}/lists/${listId}/items`,
              {
                $search: formatSearchQuery(companyName, 'displayName', 'general'),
                $top: '20',
                $expand: 'fields',
              }
            ).catch(() => JSON.stringify({ value: [] }));

            const items = JSON.parse(itemsResponse);
            return items.value || [];
          } catch (err) {
            logger.debug(`Failed to get list items for list ${listId}: ${err}`);
            return [];
          }
        });

      const allItems = await Promise.all(listItemPromises);
      return allItems.flat();
    } catch (err) {
      logger.debug(`Failed to get lists for site ${siteId}: ${err}`);
      return [];
    }
  });

  const allSiteListItems = await Promise.all(siteListItemsPromises);
  const siteListItems = allSiteListItems.flat();
  searchListItems.push(...siteListItems);
  thinking.push(`📄 Found ${siteListItems.length} SharePoint list items`);

  // Extract document content for found files (limited to avoid too many API calls)
  // Also detect and handle Loop files specially
  thinking.push(`📄 Extracting content from ${Math.min(searchFiles.length, 10)} documents`);

  // First, detect Loop files in the search results
  let loopFilesFoundCount = 0;
  for (const file of searchFiles) {
    const fileObj = file as Record<string, unknown>;
    const loopDetection = detectLoopFile(fileObj);
    if (loopDetection.isLoopFile) {
      loopFilesFoundCount++;
      fileObj.isLoopFile = true;
      fileObj.loopDetection = {
        method: loopDetection.detectionMethod,
        confidence: loopDetection.confidence,
      };
    }
  }
  if (loopFilesFoundCount > 0) {
    thinking.push(`📋 Found ${loopFilesFoundCount} Loop file(s) in search results`);
  }

  const documentContentPromises = searchFiles.slice(0, 10).map(async (file) => {
    // Limit to first 10 files
    const fileObj = file as Record<string, unknown>;
    const fileId = fileObj.id as string | undefined;
    const driveId = (fileObj.parentReference as Record<string, unknown> | undefined)?.driveId as
      | string
      | undefined;
    const fileName = (fileObj.name as string) || 'unknown';
    const fileType = fileName.split('.').pop()?.toLowerCase() || '';
    const fileIsLoop = fileObj.isLoopFile === true;

    // Try to extract content from text-based documents AND Loop files
    const textFileTypes = ['txt', 'md', 'csv', 'json', 'xml', 'html', 'htm', 'loop', 'fluid'];
    if (!fileId || (!textFileTypes.includes(fileType) && !fileIsLoop)) {
      return { file, content: null, isLoopFile: fileIsLoop };
    }

    try {
      const endpoint = driveId
        ? `/drives/${driveId}/items/${fileId}/content`
        : `/me/drive/items/${fileId}/content`;
      const content = await callGraph(graphClient, 'GET', endpoint).catch(() => null);

      if (content && typeof content === 'string') {
        // If it's a Loop file, try to parse and extract readable content
        if (fileIsLoop) {
          const parsedLoop = parseLoopContent(content);
          if (parsedLoop.success && parsedLoop.textContent) {
            const maxContentLength = 5000;
            const truncatedContent =
              parsedLoop.textContent.length > maxContentLength
                ? parsedLoop.textContent.substring(0, maxContentLength) + '...'
                : parsedLoop.textContent;
            return {
              file,
              content: truncatedContent,
              isLoopFile: true,
              loopMetadata: parsedLoop.metadata,
            };
          }
        }

        // Limit content length to avoid huge responses
        const maxContentLength = 5000;
        const truncatedContent =
          content.length > maxContentLength
            ? content.substring(0, maxContentLength) + '...'
            : content;
        return { file, content: truncatedContent, isLoopFile: fileIsLoop };
      }
    } catch (err) {
      logger.debug(`Failed to extract content from file ${fileName}: ${err}`);
    }

    return { file, content: null, isLoopFile: fileIsLoop };
  });

  const documentContents = await Promise.all(documentContentPromises);
  const filesWithContent = documentContents.filter((dc) => dc.content !== null);
  const loopFilesWithContent = documentContents.filter(
    (dc) => dc.isLoopFile && dc.content !== null
  );
  thinking.push(`✅ Extracted content from ${filesWithContent.length} documents`);
  if (loopFilesWithContent.length > 0) {
    thinking.push(`📋 Extracted content from ${loopFilesWithContent.length} Loop file(s)`);
  }

  // Combine all results (search API + direct searches)
  const allEmailsFromDirect = emailResults;
  const allEventsFromDirect = eventResults;

  // Combine search API results with direct API results (deduplicate by ID)
  const emailMap = new Map<string, unknown>();
  const eventMap = new Map<string, unknown>();

  // Add search API results
  for (const email of searchEmails) {
    const emailObj = email as Record<string, unknown>;
    const id = emailObj.id as string;
    if (id) emailMap.set(id, email);
  }
  for (const event of searchEvents) {
    const eventObj = event as Record<string, unknown>;
    const id = eventObj.id as string;
    if (id) eventMap.set(id, event);
  }

  // Add direct API results (will overwrite duplicates)
  for (const email of allEmailsFromDirect) {
    const emailObj = email as Record<string, unknown>;
    const id = emailObj.id as string;
    if (id) emailMap.set(id, email);
  }
  for (const event of allEventsFromDirect) {
    const eventObj = event as Record<string, unknown>;
    const id = eventObj.id as string;
    if (id) eventMap.set(id, event);
  }

  const allEmails = Array.from(emailMap.values());
  const allEvents = Array.from(eventMap.values());
  const allFiles = [...searchFiles, ...filesResults];
  const allContacts = [...searchContacts, ...contactsResults];
  const allSites = [...searchSites, ...sitesResults];
  const allListItems = searchListItems; // SharePoint list items

  // Store raw counts BEFORE aggregation for accurate statistics
  const rawCounts = {
    emails: allEmails.length,
    events: allEvents.length,
    files: allFiles.length,
    contacts: allContacts.length,
    sites: allSites.length,
    listItems: allListItems.length,
  };

  // Use DataAggregator for consistent deduplication and sorting
  // Increase maxItems to ensure we capture items from all sources
  // Use a per-source balanced approach: aggregate each source separately, then combine top items
  const sourceAggregations: Record<string, AggregationResult> = {};
  const sourceLimit = Math.ceil((limit * 2) / 6); // Distribute limit across 6 sources

  // Aggregate each source separately to ensure representation
  const sources = [
    { name: 'emails', items: allEmails },
    { name: 'calendar', items: allEvents },
    { name: 'files', items: allFiles },
    { name: 'contacts', items: allContacts },
    { name: 'sites', items: allSites },
    { name: 'listItems', items: allListItems },
  ];

  for (const { name, items } of sources) {
    if (items.length > 0) {
      sourceAggregations[name] = dataAggregator.aggregate([{ source: name, items }], {
        sortBy: 'timestamp',
        sortOrder: 'desc',
        maxItems: sourceLimit,
        deduplicate: true,
      });
    }
  }

  // Combine top items from each source
  const combinedItems: AggregatedItem[] = [];
  for (const { name } of sources) {
    const agg = sourceAggregations[name];
    if (agg) {
      combinedItems.push(...agg.items);
    }
  }

  // Sort combined items by timestamp
  combinedItems.sort((a, b) => {
    const aTime = a.timestamp?.getTime() || 0;
    const bTime = b.timestamp?.getTime() || 0;
    return bTime - aTime; // Descending
  });

  // Limit to maxItems total
  const limitedItems = combinedItems.slice(0, limit * 2);

  // Create aggregated result structure
  const aggregated: AggregationResult = {
    items: limitedItems,
    totalItems: combinedItems.length,
    uniqueItems: limitedItems.length,
    sources: sources.filter((s) => sourceAggregations[s.name]).map((s) => s.name),
  };

  // Log search statistics for debugging
  thinking.push(`🔍 Search statistics:`);
  thinking.push(`   Microsoft Search API: ${searchApiResults.length} items found`);
  thinking.push(
    `   Direct API calls: ${rawCounts.emails} emails, ${rawCounts.events} events, ${rawCounts.files} files, ${rawCounts.contacts} contacts, ${rawCounts.sites} sites`
  );

  // Categorize aggregated items by source
  const categorizedResults: Record<string, unknown[]> = {
    emails: [],
    meetings: [],
    files: [],
    contacts: [],
    sites: [],
    listItems: [],
  };

  // Create a map of files with content
  const fileContentMap = new Map<string, string>();
  for (const dc of documentContents) {
    const fileObj = dc.file as Record<string, unknown>;
    const fileId = fileObj.id as string | undefined;
    if (fileId && dc.content) {
      fileContentMap.set(fileId, dc.content);
    }
  }

  for (const item of aggregated.items) {
    const data = item.data as Record<string, unknown>;
    if (item.source === 'emails') {
      categorizedResults.emails.push(data);
    } else if (item.source === 'calendar') {
      categorizedResults.meetings.push(data);
    } else if (item.source === 'files') {
      // Add content if available
      const fileId = data.id as string | undefined;
      if (fileId && fileContentMap.has(fileId)) {
        (data as Record<string, unknown>).extractedContent = fileContentMap.get(fileId);
      }
      categorizedResults.files.push(data);
    } else if (item.source === 'contacts') {
      categorizedResults.contacts.push(data);
    } else if (item.source === 'sites') {
      categorizedResults.sites.push(data);
    } else if (item.source === 'listItems') {
      categorizedResults.listItems.push(data);
    }
  }

  // Calculate relationship score based on interaction frequency
  // Use raw counts for accurate scoring, but categorized counts for results
  const emailCount = rawCounts.emails;
  const meetingCount = rawCounts.events;
  const contactCount = rawCounts.contacts;
  const fileCount = rawCounts.files;
  const siteCount = rawCounts.sites;
  const listItemCount = rawCounts.listItems;

  // Counts for results (limited by aggregation)
  const resultEmailCount = categorizedResults.emails.length;
  const resultMeetingCount = categorizedResults.meetings.length;
  const resultContactCount = categorizedResults.contacts.length;
  const resultFileCount = categorizedResults.files.length;
  const resultSiteCount = categorizedResults.sites.length;
  const resultListItemCount = categorizedResults.listItems.length;

  const relationshipScore = Math.min(
    100,
    Math.round(
      ((emailCount * 2 + meetingCount * 5 + contactCount * 3 + fileCount * 1 + listItemCount * 2) /
        days) *
        10
    )
  );

  // Extract key findings for LLM to use directly
  const keyFindings: {
    companyInfo?: string[];
    keyContacts?: Array<{ name?: string; email?: string; role?: string }>;
    recentTopics?: string[];
    importantFiles?: Array<{ name?: string; webUrl?: string; lastModified?: string }>;
    sites?: Array<{ name?: string; webUrl?: string }>;
  } = {};

  // Extract company info from email signatures and content
  const companyInfoSet = new Set<string>();
  for (const email of categorizedResults.emails.slice(0, 10)) {
    const emailObj = email as Record<string, unknown>;
    const bodyObj = emailObj.body as Record<string, unknown> | undefined;
    const body = bodyObj?.content as string | undefined;
    const subject = emailObj.subject as string | undefined;

    if (body) {
      // Look for company name patterns in email body
      const companyPattern = new RegExp(`${companyName}[^\\s]*`, 'gi');
      const matches = body.match(companyPattern);
      if (matches) {
        matches.forEach((m) => companyInfoSet.add(m));
      }

      // Extract address patterns
      const addressPattern = /(\d{5})\s+([A-ZÄÖÜ][a-zäöüß\s]+)/g;
      const addressMatches = body.match(addressPattern);
      if (addressMatches) {
        addressMatches.forEach((m) => companyInfoSet.add(m.trim()));
      }
    }
  }
  if (companyInfoSet.size > 0) {
    keyFindings.companyInfo = Array.from(companyInfoSet).slice(0, 10);
  }

  // Extract key contacts from emails and contacts
  const contactsMap = new Map<string, { name?: string; email?: string; role?: string }>();

  // From contacts
  for (const contact of categorizedResults.contacts) {
    const contactObj = contact as Record<string, unknown>;
    const emailAddresses = contactObj.emailAddresses as Array<Record<string, unknown>> | undefined;
    const email = emailAddresses?.[0]?.address as string | undefined;
    const name = contactObj.displayName as string | undefined;
    const company = contactObj.companyName as string | undefined;

    if (email && !contactsMap.has(email)) {
      contactsMap.set(email, {
        name: name || email.split('@')[0],
        email,
        role: company,
      });
    }
  }

  // From emails (senders and recipients)
  for (const email of categorizedResults.emails.slice(0, 20)) {
    const emailObj = email as Record<string, unknown>;
    const from = emailObj.from as Record<string, unknown> | undefined;
    const toRecipients = (emailObj.toRecipients as Array<Record<string, unknown>>) || [];

    if (from) {
      const emailAddress = from.emailAddress as Record<string, unknown> | undefined;
      const email = emailAddress?.address as string | undefined;
      const name = emailAddress?.name as string | undefined;
      if (email && email.includes(companyName.toLowerCase()) && !contactsMap.has(email)) {
        contactsMap.set(email, {
          name: name || email.split('@')[0],
          email,
        });
      }
    }

    for (const recipient of toRecipients) {
      const emailAddress = recipient.emailAddress as Record<string, unknown> | undefined;
      const email = emailAddress?.address as string | undefined;
      const name = emailAddress?.name as string | undefined;
      if (email && email.includes(companyName.toLowerCase()) && !contactsMap.has(email)) {
        contactsMap.set(email, {
          name: name || email.split('@')[0],
          email,
        });
      }
    }
  }

  if (contactsMap.size > 0) {
    keyFindings.keyContacts = Array.from(contactsMap.values()).slice(0, 10);
  }

  // Extract recent topics from email subjects
  const topicsSet = new Set<string>();
  for (const email of categorizedResults.emails.slice(0, 20)) {
    const emailObj = email as Record<string, unknown>;
    const subject = emailObj.subject as string | undefined;
    if (subject && subject.length > 0) {
      // Remove common prefixes like "RE:", "AW:", "FWD:"
      const cleanSubject = subject.replace(/^(RE:|AW:|FWD:|FW:)\s*/i, '').trim();
      if (cleanSubject.length > 5 && cleanSubject.length < 100) {
        topicsSet.add(cleanSubject);
      }
    }
  }
  if (topicsSet.size > 0) {
    keyFindings.recentTopics = Array.from(topicsSet).slice(0, 10);
  }

  // Extract important files
  const importantFiles: Array<{ name?: string; webUrl?: string; lastModified?: string }> = [];
  for (const file of categorizedResults.files.slice(0, 10)) {
    const fileObj = file as Record<string, unknown>;
    const name = fileObj.name as string | undefined;
    const webUrl = fileObj.webUrl as string | undefined;
    const lastModified = fileObj.lastModifiedDateTime as string | undefined;

    if (name) {
      importantFiles.push({
        name,
        webUrl,
        lastModified,
      });
    }
  }
  if (importantFiles.length > 0) {
    keyFindings.importantFiles = importantFiles;
  }

  // Extract sites
  const siteFindings: Array<{ name?: string; webUrl?: string }> = [];
  for (const site of categorizedResults.sites) {
    const siteObj = site as Record<string, unknown>;
    const name = siteObj.displayName as string | undefined;
    const webUrl = siteObj.webUrl as string | undefined;

    if (name || webUrl) {
      siteFindings.push({
        name: name || webUrl,
        webUrl,
      });
    }
  }
  if (siteFindings.length > 0) {
    keyFindings.sites = siteFindings;
  }

  // Build discovery response
  const response: DiscoveryResponse = {
    target: companyName,
    targetType: 'company',
    nlpAnalysis: {
      detectedIntent: decomposed.intent.type,
      detectedEntities: decomposed.entities.map((e) => ({
        value: e.value,
        type: e.type,
        confidence: e.confidence,
      })),
      confidence: decomposed.confidence,
      suggestedFollowUps: [
        `Show me all contacts at ${companyName}`,
        `What meetings do I have with ${companyName}?`,
        `Find recent emails from ${companyName}`,
        `Show me contracts or documents for ${companyName}`,
      ],
    },
    summary: {
      totalItems: emailCount + meetingCount + fileCount + contactCount + siteCount + listItemCount,
      sources: aggregated.sources || [],
      timeRange: `Last ${days} days`,
    },
    results: {
      emails: categorizedResults.emails.slice(0, limit),
      meetings: categorizedResults.meetings.slice(0, limit),
      files: categorizedResults.files.slice(0, limit),
      contacts: categorizedResults.contacts,
      sites: categorizedResults.sites,
      listItems: categorizedResults.listItems.slice(0, limit),
      documentsWithContent: filesWithContent.map((dc) => ({
        file: dc.file,
        contentPreview: dc.content,
      })),
    },
    insights: {
      relationshipScore,
      lastInteraction:
        (categorizedResults.emails[0] as { receivedDateTime?: string } | undefined)
          ?.receivedDateTime ||
        (categorizedResults.meetings[0] as { start?: { dateTime?: string } } | undefined)?.start
          ?.dateTime ||
        'Unknown',
      recentActivity: `${emailCount} emails, ${meetingCount} meetings, ${fileCount} files, ${contactCount} contacts, ${siteCount} sites`,
      recommendations: generateCompanyRecommendations(
        relationshipScore,
        emailCount,
        meetingCount,
        companyName
      ),
      dataSummary: `IMPORTANT: Found ${emailCount} emails, ${fileCount} files, ${siteCount} sites, ${contactCount} contacts, and ${meetingCount} meetings related to ${companyName}. Use the data in 'results' and 'keyFindings' sections to answer questions about ${companyName}. Do NOT use general knowledge - use only the data provided here.`,
    },
    keyFindings: Object.keys(keyFindings).length > 0 ? keyFindings : undefined,
  };

  thinking.push(`✅ Customer 360 discovery complete:`);
  thinking.push(`   📧 Emails: ${emailCount} (${resultEmailCount} in results)`);
  thinking.push(`   📅 Meetings: ${meetingCount} (${resultMeetingCount} in results)`);
  thinking.push(`   📁 Files: ${fileCount} (${resultFileCount} in results)`);
  thinking.push(`   👥 Contacts: ${contactCount} (${resultContactCount} in results)`);
  thinking.push(`   🌐 Sites: ${siteCount} (${resultSiteCount} in results)`);
  thinking.push(
    `   📄 SharePoint List Items: ${listItemCount} (${resultListItemCount} in results)`
  );
  thinking.push(`   📄 Documents with extracted content: ${filesWithContent.length}`);
  thinking.push(`   💯 Relationship Score: ${relationshipScore}/100`);

  return addThinkingToResponse(JSON.stringify(response, null, 2), thinking);
}

/**
 * Generate recommendations for person discovery
 */
function generatePersonRecommendations(
  emails: unknown[] | undefined,
  meetings: unknown[],
  personName: string
): string[] {
  const recommendations: string[] = [];

  if (!emails || emails.length === 0) {
    recommendations.push(`💡 No recent emails found with ${personName}`);
  } else if (emails.length > 20) {
    recommendations.push(`📧 High email volume with ${personName} - consider scheduling a call`);
  }

  if (meetings.length === 0) {
    recommendations.push(`📅 No meetings scheduled with ${personName}`);
  } else if (meetings.length > 5) {
    recommendations.push(`📅 Frequent meetings with ${personName}`);
  }

  return recommendations;
}

/**
 * Generate recommendations for company discovery
 */
function generateCompanyRecommendations(
  relationshipScore: number,
  emailCount: number,
  meetingCount: number,
  companyName: string
): string[] {
  const recommendations: string[] = [];

  if (relationshipScore >= 80) {
    recommendations.push(
      `🌟 Strong relationship with ${companyName} (Score: ${relationshipScore})`
    );
  } else if (relationshipScore >= 50) {
    recommendations.push(`👍 Good relationship with ${companyName} (Score: ${relationshipScore})`);
  } else if (relationshipScore >= 20) {
    recommendations.push(
      `📈 Growing relationship with ${companyName} (Score: ${relationshipScore})`
    );
  } else {
    recommendations.push(
      `💡 Consider strengthening relationship with ${companyName} (Score: ${relationshipScore})`
    );
  }

  if (emailCount > 0 && meetingCount === 0) {
    recommendations.push(`📅 Consider scheduling a meeting with ${companyName}`);
  }

  if (emailCount === 0 && meetingCount === 0) {
    recommendations.push(`📧 No recent interactions - consider reaching out to ${companyName}`);
  }

  return recommendations;
}

// ============================================================================
// REGISTRATION FUNCTION
// ============================================================================
// Global DownloadLinkGenerator instance (initialized when secrets available)
let downloadLinkGenerator: DownloadLinkGenerator | null = null;

export function registerSuperTools(
  server: McpServer,
  graphClient: GraphClient,
  readOnly: boolean = false,
  secrets?: AppSecrets
): void {
  logger.info(`Registering Super-Tools (consolidated interface, readOnly=${readOnly})`);

  // Initialize DownloadLinkGenerator if secrets provided
  if (secrets) {
    downloadLinkGenerator = new DownloadLinkGenerator(graphClient, secrets);
    // Also set in UQAS integration
    uqas.setDownloadLinkGenerator(downloadLinkGenerator);
  }

  // 0. SEARCH (Microsoft 365 Unified Search - RECOMMENDED FIRST TOOL)
  server.tool(
    'search',
    'Recommended first tool for any Microsoft 365 question. Use search before email, calendar, files, or teams to find content across all M365 services (emails, calendar events, files, SharePoint, Teams messages, people). Returns unified results and suggestedNextTools with exact parameters for the next call. When to use: start here for "find X", "emails about Y", "meetings with Z", or cross-product discovery. Entity types: message, event, driveItem, site, list, listItem, chatMessage, person.',
    searchSchema.shape,
    async (input: SearchInput) => {
      try {
        const result = await handleSearch(input, graphClient, readOnly);
        return { content: [{ type: 'text' as const, text: result }] };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // 1. Email
  server.tool(
    'email',
    `Unified email operations for Outlook/Exchange. Read operations: list messages (with pagination, filtering, search), get message details, list folders and subfolders, get attachments, search emails. ${readOnly ? '' : 'Write operations: send new emails, reply to messages, delete emails, move emails to folders.'} Use this tool when you need to work with email messages, folders, or attachments. Supports OData filtering, search queries, and pagination.`,
    emailSchema.shape,
    async (input: EmailInput) => {
      try {
        const result = await handleEmail(input, graphClient, readOnly);
        return { content: [{ type: 'text' as const, text: result }] };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // 2. Calendar
  server.tool(
    'calendar',
    `Unified calendar operations for Outlook Calendar. Read operations: list events (with filtering and pagination), get event details, view calendar events in date range, list available calendars. ${readOnly ? '' : 'Write operations: create new events, update existing events, delete events.'} Supports timezone handling, date range queries, attendee management, and online meeting creation. Use this tool when working with calendar events, scheduling, or meeting information.`,
    calendarSchema.shape,
    async (input: CalendarInput) => {
      try {
        const result = await handleCalendar(input, graphClient, readOnly);
        return { content: [{ type: 'text' as const, text: result }] };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // 3. Teams
  server.tool(
    'teams',
    'Unified Teams operations: list teams, channels, chats, and chat messages. When to use: after search finds chatMessage results, or when user asks for Teams chats/channels by name.',
    teamsSchema.shape,
    async (input: TeamsInput) => {
      try {
        const result = await handleTeams(input, graphClient, readOnly);
        return { content: [{ type: 'text' as const, text: result }] };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // 4. Files
  server.tool(
    'files',
    'Unified OneDrive/file operations: list drives, list files, get file, download, search. When to use: after search finds driveItem/listItem, or when user asks for files in a folder or by name.',
    filesSchema.shape,
    async (input: FilesInput) => {
      try {
        const result = await handleFiles(input, graphClient, readOnly);
        return { content: [{ type: 'text' as const, text: result }] };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // 5. Tasks
  server.tool(
    'tasks',
    'Unified task operations: To-Do lists/tasks, Planner plans/tasks. When to use: user asks for tasks, to-do, or planner; or after assistant my-week shows task lists.',
    tasksSchema.shape,
    async (input: TasksInput) => {
      try {
        const result = await handleTasks(input, graphClient, readOnly);
        return { content: [{ type: 'text' as const, text: result }] };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // 6. Contacts
  server.tool(
    'contacts',
    'Unified contact and user operations: list contacts, users, current user, search. When to use: resolve recipient names/emails before send-mail, or list org users.',
    contactsSchema.shape,
    async (input: ContactsInput) => {
      try {
        const result = await handleContacts(input, graphClient, readOnly);
        return { content: [{ type: 'text' as const, text: result }] };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // 7. Meetings
  server.tool(
    'meetings',
    'Unified online meeting operations: list meetings, get meeting, transcripts, recordings. When to use: user asks for meeting transcript or recording; requires org-mode and appropriate permissions.',
    meetingsSchema.shape,
    async (input: MeetingsInput) => {
      try {
        const result = await handleMeetings(input, graphClient, readOnly);
        return { content: [{ type: 'text' as const, text: result }] };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // 8. SharePoint
  server.tool(
    'sharepoint',
    'Unified SharePoint operations: search sites, get site, site drives, site lists. When to use: after search finds site/listItem, or when user asks for SharePoint sites or lists.',
    sharepointSchema.shape,
    async (input: SharePointInput) => {
      try {
        const result = await handleSharePoint(input, graphClient, readOnly);
        return { content: [{ type: 'text' as const, text: result }] };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // 9. Notes (OneNote)
  server.tool(
    'notes',
    'Unified OneNote operations: notebooks, sections, pages, page content. When to use: user asks for OneNote notes or notebooks by name.',
    notesSchema.shape,
    async (input: NotesInput) => {
      try {
        const result = await handleNotes(input, graphClient, readOnly);
        return { content: [{ type: 'text' as const, text: result }] };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // 10. Assistant
  server.tool(
    'assistant',
    'Use for my-day, my-week, discover, follow-ups, person-info, project-overview (high-level summaries and discovery). For specific operations (list emails, get one event, send mail, list files) use the domain tools: email, calendar, files, teams, tasks. When to use: daily/weekly digest, "everything about X", or multi-source overview; otherwise prefer search first, then email/calendar/files/teams.',
    assistantSchema.shape,
    async (input: AssistantInput) => {
      try {
        const result = await handleAssistant(input, graphClient, readOnly);
        return { content: [{ type: 'text' as const, text: result }] };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // 11. Product Search
  server.tool(
    'product-search',
    'Product-based search: First uses Microsoft 365 Search to find results, then detects affected products and provides product-specific summaries with top 3-5 results per product',
    productSearchSchema.shape,
    async (input: ProductSearchInput) => {
      try {
        const result = await handleProductSearch(input, graphClient, readOnly);
        return { content: [{ type: 'text' as const, text: result }] };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  logger.info('Registered 12 Super-Tools (search is the recommended first tool)');
}
