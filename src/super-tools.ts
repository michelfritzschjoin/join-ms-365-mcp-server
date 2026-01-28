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
import GraphClient from './graph-client.js';
import logger from './logger.js';
import { addThinkingToResponse, isThinkingEnabled } from './thinking-process.js';
import {
  formatCalendarResponse,
  calendarResponseToText,
  formatMailResponse,
  mailResponseToText,
  isCalendarResponse,
  isMailResponse,
} from './response-formatter.js';
import NLPEnhancer, { type DecomposedQuery, type ExtractedEntity } from './nlp-enhancer.js';
import DataAggregator from './data-aggregator.js';
import {
  GraphApiError,
  RateLimitError,
  ServiceUnavailableError,
  isRetryableError,
  getRetryAfter,
} from './errors.js';

// Initialize NLP Enhancer for intelligent query processing
const nlpEnhancer = new NLPEnhancer();

// Initialize Data Aggregator for consistent data processing
const dataAggregator = new DataAggregator();

/**
 * Format search query for Microsoft Graph API endpoints that require property:value format
 * @param searchValue - The search query string
 * @param defaultProperty - The default property to use if not already specified (e.g., 'displayName')
 * @returns Formatted search query in property:value format, wrapped in double quotes
 */
function formatSearchQuery(searchValue: string, defaultProperty = 'displayName'): string {
  if (!searchValue) return '';

  // Check if search already contains a property prefix (e.g., "displayName:John")
  const propertyValuePattern = /^[a-zA-Z]+:/i;
  const trimmedValue = searchValue.trim();
  const formattedSearch = propertyValuePattern.test(trimmedValue)
    ? trimmedValue
    : `${defaultProperty}:${trimmedValue}`;

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
  } = options;

  const metadata: StandardResponseMetadata = {
    timestamp: new Date().toISOString(),
    executionTime,
    sources,
    cacheHit,
    ...(pagination && { pagination }),
    ...(requestId && { requestId }),
  };

  const response: StandardResponse<T> = {
    success,
    ...(data !== undefined && { data }),
    metadata,
    ...(errors && errors.length > 0 && { errors }),
    ...(suggestions && suggestions.length > 0 && { suggestions }),
    ...(nlpAnalysis && { nlpAnalysis }),
    ...(thinking && thinking.length > 0 && { thinking }),
  };

  return response;
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
  if (normalized !== query) {
    optimizedQuery = normalized;
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

// Common schemas
const paginationSchema = {
  top: z.number().optional().describe('Maximum number of items to return (default: 25)'),
  skip: z.number().optional().describe('Number of items to skip for pagination'),
};

const filterSchema = {
  filter: z.string().optional().describe('OData filter expression'),
  search: z.string().optional().describe('Search query string'),
  orderby: z.string().optional().describe('OData orderby expression'),
};

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
  'attachments',
  'search',
  // Write operations (blocked in read-only mode)
  'send',
  'reply',
  'delete',
  'move',
]);

const emailSchema = z.object({
  action: emailActions.describe(
    'The email operation: list, get, folders, attachments, search (read) | send, reply, delete, move (write)'
  ),
  // Identifiers
  messageId: z
    .string()
    .optional()
    .describe('Message ID (required for get, attachments, reply, delete, move)'),
  folderId: z.string().optional().describe('Folder ID to list messages from or move to'),
  attachmentId: z.string().optional().describe('Attachment ID (for getting specific attachment)'),
  // For send/reply
  to: z.string().optional().describe('Recipient email address(es), comma-separated (for send)'),
  subject: z.string().optional().describe('Email subject (for send)'),
  body: z.string().optional().describe('Email body content (for send/reply)'),
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

      // Format mail response with quick summary
      if (isMailResponse(parsedResult)) {
        const formatted = formatMailResponse(parsedResult);
        const formattedText = mailResponseToText(formatted);

        // Add metadata to response
        const responseWithMetadata = formatStandardResponse(
          { formatted: formattedText, raw: parsedResult },
          {
            executionTime,
            sources: ['email'],
            cacheHit: false,
            pagination,
            suggestions: [
              '💡 Use "email" tool with action "get" to view email details',
              '💡 Use "email" tool with action "search" for advanced search',
            ],
          }
        );

        return addThinkingToResponse(JSON.stringify(responseWithMetadata, null, 2), thinking);
      }

      // Fallback: return with metadata
      const responseWithMetadata = formatStandardResponse(parsedResult, {
        executionTime,
        sources: ['email'],
        cacheHit: false,
        pagination,
      });

      return addThinkingToResponse(JSON.stringify(responseWithMetadata, null, 2), thinking);
    }

    case 'get': {
      if (!input.messageId) throw new Error('messageId is required for action "get"');
      thinking.push(`Getting email with ID: ${input.messageId}`);
      const result = await callGraph(graphClient, 'GET', `/me/messages/${input.messageId}`);
      return addThinkingToResponse(result, thinking);
    }

    case 'folders': {
      thinking.push('Listing mail folders');
      const params: Record<string, string> = { $top: String(input.top || 50) };
      const result = await callGraph(graphClient, 'GET', '/me/mailFolders', params);
      return addThinkingToResponse(result, thinking);
    }

    case 'attachments': {
      if (!input.messageId) throw new Error('messageId is required for action "attachments"');
      thinking.push(`Getting attachments for message: ${input.messageId}`);
      const endpoint = input.attachmentId
        ? `/me/messages/${input.messageId}/attachments/${input.attachmentId}`
        : `/me/messages/${input.messageId}/attachments`;
      const result = await callGraph(graphClient, 'GET', endpoint);
      return addThinkingToResponse(result, thinking);
    }

    case 'search': {
      if (!input.search) throw new Error('search query is required for action "search"');

      // NLP optimization for search query
      const startTime = Date.now();
      const optimized = optimizeQueryWithNLP(input.search);
      thinking.push(`🔍 Searching emails for: "${input.search}"`);
      if (optimized.optimizedQuery !== input.search) {
        thinking.push(`💡 NLP optimized query: "${optimized.optimizedQuery}"`);
      }
      if (optimized.nlpAnalysis.intent) {
        thinking.push(`📊 Detected intent: ${optimized.nlpAnalysis.intent}`);
      }

      const params: Record<string, string> = {
        $search: formatSearchQuery(optimized.optimizedQuery),
        $top: String(input.top || 25),
      };

      // Apply temporal filters if NLP detected them
      if (optimized.filters?.dateFilter) {
        const dateFilter = optimized.filters.dateFilter as string;
        params.$filter = `receivedDateTime ge ${dateFilter}`;
        thinking.push(`📅 Applying date filter: after ${dateFilter}`);
      }

      const result = await callGraph(graphClient, 'GET', '/me/messages', params);
      const parsedResult = JSON.parse(result);
      const executionTime = Date.now() - startTime;

      // Format mail response with quick summary
      if (isMailResponse(parsedResult)) {
        const formatted = formatMailResponse(parsedResult);
        const formattedText = mailResponseToText(formatted);

        // Add NLP insights to response
        const responseWithMetadata = formatStandardResponse(
          { formatted: formattedText, raw: parsedResult },
          {
            executionTime,
            sources: ['email'],
            cacheHit: false,
            nlpAnalysis: optimized.nlpAnalysis,
            suggestions: [
              '💡 Use "email" tool with action "get" to view full email details',
              '💡 Use "email" tool with action "list" to browse more emails',
            ],
          }
        );

        return addThinkingToResponse(JSON.stringify(responseWithMetadata, null, 2), thinking);
      }

      return addThinkingToResponse(result, thinking);
    }

    // Write operations (blocked in read-only mode - check happens at function start)
    case 'send': {
      if (!input.to) throw new Error('to (recipient) is required for action "send"');
      if (!input.subject) throw new Error('subject is required for action "send"');
      if (!input.body) throw new Error('body is required for action "send"');
      thinking.push(`Sending email to: ${input.to}`);
      const recipients = input.to.split(',').map((email) => ({
        emailAddress: { address: email.trim() },
      }));
      const message = {
        subject: input.subject,
        body: { contentType: 'Text', content: input.body },
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
      if (!input.body) throw new Error('body is required for action "reply"');
      thinking.push(`Replying to email: ${input.messageId}`);
      const result = await callGraph(
        graphClient,
        'POST',
        `/me/messages/${input.messageId}/reply`,
        undefined,
        { comment: input.body }
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
      return addThinkingToResponse(result, thinking);
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
  timezone: z.string().optional().describe('Timezone for date/time values (e.g., "Europe/Berlin")'),
  // For create/update event
  subject: z.string().optional().describe('Event subject/title (for create-event, update-event)'),
  body: z.string().optional().describe('Event body/description (for create-event, update-event)'),
  location: z.string().optional().describe('Event location (for create-event, update-event)'),
  attendees: z.string().optional().describe('Attendee emails, comma-separated (for create-event)'),
  isOnline: z.boolean().optional().describe('Create as online meeting (for create-event)'),
  // Filters
  ...filterSchema,
  ...paginationSchema,
});

type CalendarInput = z.infer<typeof calendarSchema>;

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
      if (isCalendarResponse(parsedResult)) {
        const formatted = formatCalendarResponse(parsedResult);
        const formattedText = calendarResponseToText(formatted);

        // Add metadata to response
        const responseWithMetadata = formatStandardResponse(
          { formatted: formattedText, raw: parsedResult },
          {
            executionTime,
            sources: ['calendar'],
            cacheHit: false,
            pagination,
            suggestions: [
              '💡 Use "calendar" tool with action "get" to view event details',
              '💡 Use "calendar" tool with action "view" for date range queries',
            ],
          }
        );

        return addThinkingToResponse(JSON.stringify(responseWithMetadata, null, 2), thinking);
      }

      // Fallback: return with metadata
      const responseWithMetadata = formatStandardResponse(parsedResult, {
        executionTime,
        sources: ['calendar'],
        cacheHit: false,
        pagination,
      });

      return addThinkingToResponse(JSON.stringify(responseWithMetadata, null, 2), thinking);
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
      return addThinkingToResponse(result, thinking);
    }

    case 'view': {
      if (!input.startDateTime || !input.endDateTime) {
        throw new Error('startDateTime and endDateTime are required for action "view"');
      }
      thinking.push(`Getting calendar view from ${input.startDateTime} to ${input.endDateTime}`);
      const params: Record<string, string> = {
        startDateTime: input.startDateTime,
        endDateTime: input.endDateTime,
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

      // Format calendar response with quick summary
      if (isCalendarResponse(parsedResult)) {
        const formatted = formatCalendarResponse(
          parsedResult,
          input.startDateTime,
          input.endDateTime
        );
        const formattedText = calendarResponseToText(formatted);
        return addThinkingToResponse(formattedText, thinking);
      }

      return addThinkingToResponse(result, thinking);
    }

    case 'calendars': {
      thinking.push('Listing all calendars');
      const result = await callGraph(graphClient, 'GET', '/me/calendars');
      return addThinkingToResponse(result, thinking);
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

      // Format calendar response with quick summary
      if (isCalendarResponse(parsedResult)) {
        const formatted = formatCalendarResponse(parsedResult);
        const formattedText = calendarResponseToText(formatted);
        return addThinkingToResponse(formattedText, thinking);
      }

      return addThinkingToResponse(result, thinking);
    }

    // Write operations (blocked in read-only mode - check happens at function start)
    case 'create-event': {
      if (!input.subject) throw new Error('subject is required for create-event');
      if (!input.startDateTime) throw new Error('startDateTime is required for create-event');
      if (!input.endDateTime) throw new Error('endDateTime is required for create-event');
      thinking.push(`Creating event: ${input.subject}`);
      const event: Record<string, unknown> = {
        subject: input.subject,
        start: { dateTime: input.startDateTime, timeZone: input.timezone || 'UTC' },
        end: { dateTime: input.endDateTime, timeZone: input.timezone || 'UTC' },
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
      return addThinkingToResponse(result, thinking);
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
      return addThinkingToResponse(result, thinking);
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
      return addThinkingToResponse(result, thinking);
    }

    case 'get-team': {
      if (!input.teamId) throw new Error('teamId is required');
      thinking.push(`Getting team: ${input.teamId}`);
      const result = await callGraph(graphClient, 'GET', `/teams/${input.teamId}`);
      return addThinkingToResponse(result, thinking);
    }

    case 'channels': {
      if (!input.teamId) throw new Error('teamId is required for channels');
      thinking.push(`Listing channels for team: ${input.teamId}`);
      const result = await callGraph(graphClient, 'GET', `/teams/${input.teamId}/channels`);
      return addThinkingToResponse(result, thinking);
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

      return addThinkingToResponse(result, thinking);
    }

    case 'chats': {
      thinking.push('Listing chats');
      const params: Record<string, string> = { $top: String(input.top || 25) };
      const chatsResult = await callGraph(graphClient, 'GET', '/me/chats', params);
      const chatsData = JSON.parse(chatsResult);

      // Default: include messages for chats action (unless explicitly disabled)
      const includeMessages = input.includeMessages !== false;

      if (includeMessages && chatsData.value && Array.isArray(chatsData.value)) {
        thinking.push(`Fetching last messages for ${chatsData.value.length} chats...`);

        // Fetch last messages for each chat (limit to avoid too many requests)
        const chatsWithMessages = await Promise.allSettled(
          chatsData.value.slice(0, 10).map(async (chat: { id: string }) => {
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

        thinking.push(`Retrieved last messages for ${formattedChats.length} chats`);
        return addThinkingToResponse(JSON.stringify(output, null, 2), thinking);
      }

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

      return addThinkingToResponse(result, thinking);
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
      return addThinkingToResponse(result, thinking);
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
      return addThinkingToResponse(result, thinking);
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
      return addThinkingToResponse(result, thinking);
    }

    case 'download': {
      if (!input.itemId) throw new Error('itemId is required for download');
      const driveId = input.driveId || 'me';
      thinking.push(`Downloading file: ${input.itemId}`);
      const endpoint =
        driveId === 'me'
          ? `/me/drive/items/${input.itemId}/content`
          : `/drives/${driveId}/items/${input.itemId}/content`;
      const result = await callGraph(graphClient, 'GET', endpoint);
      return addThinkingToResponse(result, thinking);
    }

    case 'root': {
      const driveId = input.driveId || 'me';
      thinking.push('Getting drive root');
      const endpoint = driveId === 'me' ? '/me/drive/root' : `/drives/${driveId}/root`;
      const result = await callGraph(graphClient, 'GET', endpoint);
      return addThinkingToResponse(result, thinking);
    }

    case 'search': {
      if (!input.search) throw new Error('search query is required');

      // NLP optimization for file search
      const startTime = Date.now();
      const optimized = optimizeQueryWithNLP(input.search);
      thinking.push(`🔍 Searching files for: "${input.search}"`);
      if (optimized.optimizedQuery !== input.search) {
        thinking.push(`💡 NLP optimized query: "${optimized.optimizedQuery}"`);
      }
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

      return addThinkingToResponse(JSON.stringify(responseWithMetadata, null, 2), thinking);
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
  title: z.string().optional().describe('Task title (for create-todo, update-todo)'),
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
      return addThinkingToResponse(result, thinking);
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
      return addThinkingToResponse(result, thinking);
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
      return addThinkingToResponse(result, thinking);
    }

    case 'planner-tasks': {
      thinking.push('Listing Planner tasks assigned to me');
      const result = await callGraph(graphClient, 'GET', '/me/planner/tasks');
      return addThinkingToResponse(result, thinking);
    }

    case 'planner-plans': {
      if (!input.planId) throw new Error('planId is required');
      thinking.push(`Getting Planner plan: ${input.planId}`);
      const result = await callGraph(graphClient, 'GET', `/planner/plans/${input.planId}`);
      return addThinkingToResponse(result, thinking);
    }

    case 'plan-tasks': {
      if (!input.planId) throw new Error('planId is required');
      thinking.push(`Listing tasks in plan: ${input.planId}`);
      const result = await callGraph(graphClient, 'GET', `/planner/plans/${input.planId}/tasks`);
      return addThinkingToResponse(result, thinking);
    }

    // Write operations (blocked in read-only mode - check happens at function start)
    case 'create-todo': {
      if (!input.taskListId) throw new Error('taskListId is required for create-todo');
      if (!input.title) throw new Error('title is required for create-todo');
      thinking.push(`Creating To-Do task: ${input.title}`);
      const task: Record<string, unknown> = { title: input.title };
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
      return addThinkingToResponse(result, thinking);
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
      return addThinkingToResponse(result, thinking);
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

const contactsSchema = z.object({
  action: contactsActions.describe('The contacts operation to perform'),
  // Identifiers
  contactId: z.string().optional().describe('Contact ID'),
  userId: z.string().optional().describe('User ID'),
  // Filters
  ...filterSchema,
  ...paginationSchema,
});

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
      return addThinkingToResponse(result, thinking);
    }

    case 'get': {
      if (!input.contactId) throw new Error('contactId is required');
      thinking.push(`Getting contact: ${input.contactId}`);
      const result = await callGraph(graphClient, 'GET', `/me/contacts/${input.contactId}`);
      return addThinkingToResponse(result, thinking);
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
      return addThinkingToResponse(result, thinking);
    }

    case 'current-user': {
      thinking.push('Getting current user info');
      const result = await callGraph(graphClient, 'GET', '/me');
      return addThinkingToResponse(result, thinking);
    }

    case 'search': {
      if (!input.search) throw new Error('search query is required');

      // NLP optimization for contact/user search
      const startTime = Date.now();
      const optimized = optimizeQueryWithNLP(input.search);
      thinking.push(`🔍 Searching contacts/users for: "${input.search}"`);
      if (optimized.optimizedQuery !== input.search) {
        thinking.push(`💡 NLP optimized query: "${optimized.optimizedQuery}"`);
      }
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

      return addThinkingToResponse(JSON.stringify(responseWithMetadata, null, 2), thinking);
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
      return addThinkingToResponse(result, thinking);
    }

    case 'get': {
      if (!input.meetingId) throw new Error('meetingId is required');
      thinking.push(`Getting meeting: ${input.meetingId}`);
      const result = await callGraph(graphClient, 'GET', `/me/onlineMeetings/${input.meetingId}`);
      return addThinkingToResponse(result, thinking);
    }

    case 'recordings': {
      if (!input.meetingId) throw new Error('meetingId is required');
      thinking.push(`Getting recordings for meeting: ${input.meetingId}`);
      const endpoint = input.recordingId
        ? `/me/onlineMeetings/${input.meetingId}/recordings/${input.recordingId}`
        : `/me/onlineMeetings/${input.meetingId}/recordings`;
      const result = await callGraph(graphClient, 'GET', endpoint);
      return addThinkingToResponse(result, thinking);
    }

    case 'transcripts': {
      if (!input.meetingId) throw new Error('meetingId is required');
      thinking.push(`Getting transcripts for meeting: ${input.meetingId}`);
      const endpoint = input.transcriptId
        ? `/me/onlineMeetings/${input.meetingId}/transcripts/${input.transcriptId}`
        : `/me/onlineMeetings/${input.meetingId}/transcripts`;
      const result = await callGraph(graphClient, 'GET', endpoint);
      return addThinkingToResponse(result, thinking);
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
      return addThinkingToResponse(result, thinking);
    }

    default:
      throw new Error(`Unknown meetings action: ${input.action}`);
  }
}

// ============================================================================
// 8. SHAREPOINT SUPER-TOOL
// ============================================================================
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
  // Identifiers
  siteId: z.string().optional().describe('Site ID'),
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
      return addThinkingToResponse(result, thinking);
    }

    case 'get-site': {
      if (!input.siteId) throw new Error('siteId is required');
      thinking.push(`Getting site: ${input.siteId}`);
      const result = await callGraph(graphClient, 'GET', `/sites/${input.siteId}`);
      return addThinkingToResponse(result, thinking);
    }

    case 'site-drives': {
      if (!input.siteId) throw new Error('siteId is required');
      thinking.push(`Listing drives for site: ${input.siteId}`);
      const result = await callGraph(graphClient, 'GET', `/sites/${input.siteId}/drives`);
      return addThinkingToResponse(result, thinking);
    }

    case 'site-lists': {
      if (!input.siteId) throw new Error('siteId is required');
      thinking.push(`Listing lists for site: ${input.siteId}`);
      const result = await callGraph(graphClient, 'GET', `/sites/${input.siteId}/lists`);
      return addThinkingToResponse(result, thinking);
    }

    case 'list-items': {
      if (!input.siteId || !input.listId) {
        throw new Error('siteId and listId are required');
      }
      thinking.push(`Listing items in list: ${input.listId}`);
      const params: Record<string, string> = { $top: String(input.top || 50) };
      if (input.filter) params.$filter = input.filter;
      const result = await callGraph(
        graphClient,
        'GET',
        `/sites/${input.siteId}/lists/${input.listId}/items`,
        params
      );
      return addThinkingToResponse(result, thinking);
    }

    case 'site-items': {
      if (!input.siteId) throw new Error('siteId is required');
      thinking.push(`Listing items in site: ${input.siteId}`);
      const params: Record<string, string> = { $top: String(input.top || 50) };
      const result = await callGraph(graphClient, 'GET', `/sites/${input.siteId}/items`, params);
      return addThinkingToResponse(result, thinking);
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
  'pages', // List pages in section
  'page-content', // Get page content
]);

const notesSchema = z.object({
  action: notesActions.describe('The OneNote operation to perform'),
  // Identifiers
  notebookId: z.string().optional().describe('Notebook ID'),
  sectionId: z.string().optional().describe('Section ID'),
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
      return addThinkingToResponse(result, thinking);
    }

    case 'sections': {
      if (!input.notebookId) throw new Error('notebookId is required');
      thinking.push(`Listing sections in notebook: ${input.notebookId}`);
      const result = await callGraph(
        graphClient,
        'GET',
        `/me/onenote/notebooks/${input.notebookId}/sections`
      );
      return addThinkingToResponse(result, thinking);
    }

    case 'pages': {
      if (!input.sectionId) throw new Error('sectionId is required');
      thinking.push(`Listing pages in section: ${input.sectionId}`);
      const params: Record<string, string> = { $top: String(input.top || 50) };
      const result = await callGraph(
        graphClient,
        'GET',
        `/me/onenote/sections/${input.sectionId}/pages`,
        params
      );
      return addThinkingToResponse(result, thinking);
    }

    case 'page-content': {
      if (!input.pageId) throw new Error('pageId is required');
      thinking.push(`Getting page content: ${input.pageId}`);
      const result = await callGraph(
        graphClient,
        'GET',
        `/me/onenote/pages/${input.pageId}/content`
      );
      return addThinkingToResponse(result, thinking);
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
] as const;

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

async function handleSearch(
  input: SearchInput,
  graphClient: GraphClient,
  _readOnly: boolean
): Promise<string> {
  const thinking: string[] = [];

  thinking.push(`🔍 Microsoft 365 Search: "${input.query}"`);

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
      endDate = now;
      thinking.push(`📅 Looking for past ${Math.abs(days)} days of events`);
    } else {
      // Future events (default)
      startDate = now;
      endDate = new Date(now);
      endDate.setDate(endDate.getDate() + days);
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

      const output = {
        query: input.query,
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
          '#microsoft.graph.event': events.value || [],
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
  let entityTypes = input.entityTypes;

  if (!entityTypes) {
    // Default: include all common types including chatMessage
    entityTypes = ['message', 'event', 'driveItem', 'site', 'chatMessage'];

    // Use NLP service hint
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
      // Prioritize chatMessage and person for person/Teams queries
      entityTypes = ['chatMessage', 'message', 'person', 'event'];
      thinking.push('💡 Detected person/Teams query - prioritizing chatMessage and person');
    }
  }

  thinking.push(`Searching in: ${entityTypes.join(', ')}`);

  // Improve query for better search results
  // Simplify query for Teams/person searches to just the person name
  let searchQuery = input.query;

  if (personName && (hasTeamsKeywords || entityTypes.includes('chatMessage'))) {
    // For Teams/person queries, simplify to just the person name for better results
    searchQuery = personName;
    thinking.push(
      `💡 Simplified query to person name: "${searchQuery}" for better Teams search results`
    );
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

    const output: Record<string, unknown> = {
      query: input.query,
      nlpAnalysis: {
        intent: nlpIntent,
        service: nlpService || 'general',
        entities: nlpEntities.map((e) => ({ value: e.value, type: e.type })),
        temporal: decomposed.temporal,
        confidence: decomposed.confidence,
      },
      totalHits,
      entityTypes: Object.keys(formattedResults),
      results: formattedResults,
      suggestions: [...new Set(toolSuggestions)],
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
  };
  insights?: {
    relationshipScore?: number;
    recommendations?: string[];
    recentActivity?: string;
    lastInteraction?: string;
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
  days: z.number().optional().describe('Number of days to look back (default: 7, max: 365)'),
  // Limits
  limit: z.number().optional().describe('Maximum results per category (default: 25, max: 100)'),
  // Include download links for files
  includeDownloadLinks: z
    .boolean()
    .optional()
    .describe('Include download links for discovered files (default: false)'),
});

type AssistantInput = z.infer<typeof assistantSchema>;

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

      // NLP optimization for question processing
      const startTime = Date.now();
      const optimized = optimizeQueryWithNLP(input.query);
      thinking.push(`💭 Processing question: "${input.query}"`);
      if (optimized.optimizedQuery !== input.query) {
        thinking.push(`💡 NLP optimized query: "${optimized.optimizedQuery}"`);
      }
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

      return addThinkingToResponse(JSON.stringify(responseWithMetadata, null, 2), thinking);
    }

    case 'search': {
      if (!input.query) throw new Error('query is required for search action');

      // NLP optimization for comprehensive search
      const startTime = Date.now();
      const optimized = optimizeQueryWithNLP(input.query);
      thinking.push(`🔍 Searching everything for: "${input.query}"`);
      if (optimized.optimizedQuery !== input.query) {
        thinking.push(`💡 NLP optimized query: "${optimized.optimizedQuery}"`);
      }
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

      return addThinkingToResponse(JSON.stringify(responseWithMetadata, null, 2), thinking);
    }

    case 'my-day': {
      thinking.push("Getting today's summary");
      const today = new Date();
      const tomorrow = new Date(today);
      tomorrow.setDate(tomorrow.getDate() + 1);

      const eventsResult = await callGraph(graphClient, 'GET', '/me/calendarView', {
        startDateTime: today.toISOString(),
        endDateTime: tomorrow.toISOString(),
      });
      results.todayEvents = JSON.parse(eventsResult);

      const emailsResult = await callGraph(graphClient, 'GET', '/me/messages', {
        $top: '10',
        $orderby: 'receivedDateTime desc',
      });
      results.recentEmails = JSON.parse(emailsResult);

      return addThinkingToResponse(JSON.stringify(results, null, 2), thinking);
    }

    case 'my-week': {
      thinking.push('Getting week summary');
      const today = new Date();
      const weekEnd = new Date(today);
      weekEnd.setDate(weekEnd.getDate() + 7);

      const eventsResult = await callGraph(graphClient, 'GET', '/me/calendarView', {
        startDateTime: today.toISOString(),
        endDateTime: weekEnd.toISOString(),
      });
      results.weekEvents = JSON.parse(eventsResult);

      const tasksResult = await callGraph(graphClient, 'GET', '/me/todo/lists');
      results.tasks = JSON.parse(tasksResult);

      return addThinkingToResponse(JSON.stringify(results, null, 2), thinking);
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
      const today = new Date();
      const tomorrow = new Date(today);
      tomorrow.setDate(tomorrow.getDate() + 1);

      const eventsResult = await callGraph(graphClient, 'GET', '/me/calendarView', {
        startDateTime: today.toISOString(),
        endDateTime: tomorrow.toISOString(),
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

      if (primaryEntity?.type === 'person' || decomposed.intent.type === 'who') {
        thinking.push(`🎯 Routing to: discover-person (detected person entity)`);
        return handleDiscoverPerson(
          { ...input, target: primaryEntity?.value || queryText },
          graphClient,
          decomposed,
          thinking
        );
      } else if (primaryEntity?.type === 'organization') {
        thinking.push(`🎯 Routing to: discover-company (detected organization entity)`);
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

  const startDate = new Date();
  startDate.setDate(startDate.getDate() - days);
  const endDate = new Date();
  endDate.setDate(endDate.getDate() + 30); // Include future meetings

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

  // Filter events that include this person
  const relevantEvents = (events.value || []).filter((event: Record<string, unknown>) => {
    const attendees = event.attendees as Array<{ emailAddress?: { name?: string } }> | undefined;
    const organizer = event.organizer as { emailAddress?: { name?: string } } | undefined;
    const personLower = personName.toLowerCase();

    return (
      attendees?.some((a) => a.emailAddress?.name?.toLowerCase().includes(personLower)) ||
      organizer?.emailAddress?.name?.toLowerCase().includes(personLower)
    );
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
        categorizedResults.emails[0]?.['receivedDateTime'] ||
        categorizedResults.meetings[0]?.['start']?.['dateTime'] ||
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
      totalItems: aggregated.uniqueItems,
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
      recentActivity: `Found ${aggregated.uniqueItems} unique items across Microsoft 365 (${totalHits} total hits)`,
      recommendations: [
        `📧 ${categorizedResults.emails.length} emails found`,
        `📁 ${categorizedResults.files.length} files found`,
        `📅 ${categorizedResults.meetings.length} calendar items found`,
        `💬 ${categorizedResults.chats.length} chat messages found`,
        `🌐 ${categorizedResults.sites.length} sites found`,
      ],
    },
  };

  thinking.push(`✅ Topic discovery complete:`);
  thinking.push(`   📊 Total hits: ${totalHits}, Unique items: ${aggregated.uniqueItems}`);
  thinking.push(`   📧 Emails: ${categorizedResults.emails.length}`);
  thinking.push(`   📁 Files: ${categorizedResults.files.length}`);
  thinking.push(`   📅 Meetings: ${categorizedResults.meetings.length}`);
  thinking.push(`   💬 Chats: ${categorizedResults.chats.length}`);
  thinking.push(`   🌐 Sites: ${categorizedResults.sites.length}`);

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

  // Parallel API calls for comprehensive company discovery
  const [emailsResult, eventsResult, filesResult, contactsResult, sitesResult] =
    await Promise.allSettled([
      // Emails with this company
      callGraph(graphClient, 'GET', '/me/messages', {
        $search: `"${companyName}"`,
        $top: String(limit),
        $orderby: 'receivedDateTime desc',
      }),
      // Meetings with company contacts
      callGraph(graphClient, 'GET', '/me/events', {
        $search: `"${companyName}"`,
        $top: String(limit),
      }),
      // Files related to this company
      callGraph(
        graphClient,
        'GET',
        `/me/drive/root/search(q='${encodeURIComponent(companyName)}')`
      ),
      // Contacts from this company
      callGraph(graphClient, 'GET', '/me/contacts', {
        $search: `"${companyName}"`,
        $top: '50',
      }),
      // SharePoint sites
      callGraph(graphClient, 'GET', '/sites', {
        search: companyName,
        $top: '10',
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
  const sites = sitesResult.status === 'fulfilled' ? JSON.parse(sitesResult.value) : { value: [] };

  // Use DataAggregator for consistent deduplication and sorting
  const aggregated = dataAggregator.aggregate(
    [
      { source: 'emails', items: emails.value || [] },
      { source: 'calendar', items: events.value || [] },
      { source: 'files', items: files.value || [] },
      { source: 'contacts', items: contacts.value || [] },
      { source: 'sites', items: sites.value || [] },
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
    meetings: [],
    files: [],
    contacts: [],
    sites: [],
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
    } else if (item.source === 'sites') {
      categorizedResults.sites.push(data);
    }
  }

  // Calculate relationship score based on interaction frequency
  const emailCount = categorizedResults.emails.length;
  const meetingCount = categorizedResults.meetings.length;
  const contactCount = categorizedResults.contacts.length;
  const fileCount = categorizedResults.files.length;

  const relationshipScore = Math.min(
    100,
    Math.round(((emailCount * 2 + meetingCount * 5 + contactCount * 3 + fileCount * 1) / days) * 10)
  );

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
      totalItems: aggregated.uniqueItems,
      sources: aggregated.sources || [],
      timeRange: `Last ${days} days`,
    },
    results: {
      emails: categorizedResults.emails.slice(0, limit),
      meetings: categorizedResults.meetings.slice(0, limit),
      files: categorizedResults.files.slice(0, limit),
      contacts: categorizedResults.contacts,
      sites: categorizedResults.sites,
    },
    insights: {
      relationshipScore,
      lastInteraction:
        categorizedResults.emails[0]?.['receivedDateTime'] ||
        categorizedResults.meetings[0]?.['start']?.['dateTime'] ||
        'Unknown',
      recentActivity: `${emailCount} emails, ${meetingCount} meetings, ${contactCount} contacts`,
      recommendations: generateCompanyRecommendations(
        relationshipScore,
        emailCount,
        meetingCount,
        companyName
      ),
    },
  };

  thinking.push(`✅ Customer 360 discovery complete:`);
  thinking.push(`   📧 Emails: ${emailCount}`);
  thinking.push(`   📅 Meetings: ${meetingCount}`);
  thinking.push(`   📁 Files: ${fileCount}`);
  thinking.push(`   👥 Contacts: ${contactCount}`);
  thinking.push(`   🌐 Sites: ${categorizedResults.sites.length}`);
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
export function registerSuperTools(
  server: McpServer,
  graphClient: GraphClient,
  readOnly: boolean = false
): void {
  logger.info(`Registering Super-Tools (consolidated interface, readOnly=${readOnly})`);

  // 0. SEARCH (Microsoft 365 Unified Search - RECOMMENDED FIRST TOOL)
  server.tool(
    'search',
    'Microsoft 365 Unified Search - USE THIS FIRST to find content across emails, calendar, files, SharePoint, Teams. Returns results and suggests which specific tools to use next.',
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
    `Unified email operations: list, get, folders, attachments, search${readOnly ? '' : ' | send, reply, delete, move (write)'}`,
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
    'Unified calendar operations: list events, get event, calendar view, list calendars',
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
    'Unified Teams operations: teams, channels, chats, messages',
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
    'Unified file operations: drives, list files, get file, download, search',
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
    'Unified task operations: To-Do lists/tasks, Planner plans/tasks',
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
    'Unified contact operations: contacts, users, current user, search',
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
    'Unified meeting operations: online meetings, recordings, transcripts',
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
    'Unified SharePoint operations: sites, drives, lists, items',
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
    'Unified OneNote operations: notebooks, sections, pages, content',
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
    'Smart assistant: natural language queries, search everything, daily/weekly summaries, person info, project overview, follow-ups, meeting prep',
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

  logger.info('Registered 11 Super-Tools (search is the recommended first tool)');
}
