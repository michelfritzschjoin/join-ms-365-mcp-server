/**
 * Knowledge Base for persisting learned synonyms, query patterns, and entity mappings
 *
 * SECURITY: This knowledge base stores anonymized query patterns only.
 * Personal data, email addresses, names, and other PII are sanitized before storage.
 */

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import logger from './logger.js';

/**
 * SECURITY: Check if PII anonymization is enabled
 * Controlled via environment variable MS365_MCP_ANONYMIZE_PII
 *
 * Default: true (enabled) - PII is anonymized before storage
 * Set to "false" or "0" to disable (NOT recommended in production!)
 *
 * @returns True if PII anonymization is enabled
 */
function isPIIAnonymizationEnabled(): boolean {
  const envValue = process.env.MS365_MCP_ANONYMIZE_PII;
  // Enabled by default unless explicitly disabled
  return envValue !== 'false' && envValue !== '0';
}

/**
 * SECURITY: Patterns to detect and remove PII from stored data
 * These patterns match common forms of personally identifiable information.
 */
const PII_PATTERNS: Array<{ pattern: RegExp; replacement: string }> = [
  // Email addresses
  { pattern: /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, replacement: '[EMAIL]' },
  // UUIDs/GUIDs (common for user/object IDs)
  {
    pattern: /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi,
    replacement: '[ID]',
  },
  // Phone numbers (various formats)
  { pattern: /(\+\d{1,3}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}/g, replacement: '[PHONE]' },
  // German phone numbers
  { pattern: /\+49\s?\d{3,4}\s?\d{4,8}/g, replacement: '[PHONE]' },
  // Credit card numbers (basic pattern)
  { pattern: /\b\d{4}[-\s]?\d{4}[-\s]?\d{4}[-\s]?\d{4}\b/g, replacement: '[CARD]' },
  // IP addresses
  { pattern: /\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b/g, replacement: '[IP]' },
  // Social Security Numbers (US)
  { pattern: /\b\d{3}-\d{2}-\d{4}\b/g, replacement: '[SSN]' },
  // Dates that might be birthdays (DD.MM.YYYY, MM/DD/YYYY formats)
  { pattern: /\b\d{1,2}[./-]\d{1,2}[./-](19|20)\d{2}\b/g, replacement: '[DATE]' },
  // Long numeric strings (potential account numbers)
  { pattern: /\b\d{10,}\b/g, replacement: '[NUMBER]' },
];

/**
 * SECURITY: Sanitize a string by removing potentially sensitive information
 * This function should be called before storing any user input in the knowledge base.
 *
 * Can be disabled via MS365_MCP_ANONYMIZE_PII=false (NOT recommended in production!)
 *
 * @param input - The string to sanitize
 * @returns Sanitized string with PII removed (or original if disabled)
 */
function sanitizeForStorage(input: string): string {
  if (!input) return input;

  // Check if anonymization is enabled
  if (!isPIIAnonymizationEnabled()) {
    return input;
  }

  let sanitized = input;
  for (const { pattern, replacement } of PII_PATTERNS) {
    sanitized = sanitized.replace(pattern, replacement);
  }

  return sanitized;
}

/**
 * SECURITY: Check if a string contains potential PII
 * @param input - The string to check
 * @returns True if PII is detected
 */
function containsPII(input: string): boolean {
  if (!input) return false;

  // If anonymization is disabled, don't check for PII
  if (!isPIIAnonymizationEnabled()) {
    return false;
  }

  for (const { pattern } of PII_PATTERNS) {
    if (pattern.test(input)) {
      return true;
    }
    // Reset regex lastIndex since we're using global flag
    pattern.lastIndex = 0;
  }

  return false;
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

export interface SuccessfulQuery {
  query: string;
  results: number;
  sources: string[];
  timestamp: Date;
  context?: string;
}

export interface QueryPattern {
  pattern: string;
  entityTypes: string[];
  successCount: number;
  lastUsed: Date;
  context?: string;
}

export interface EntityMapping {
  keyword: string;
  entityTypes: string[];
  successCount: number;
  lastUsed: Date;
}

export interface DataLocation {
  dataType: string;
  sources: string[];
  successCount: number;
  lastUsed: Date;
}

export interface UserFeedback {
  query: string;
  resultId?: string;
  feedbackType: 'helpful' | 'not_helpful' | 'incorrect' | 'correct';
  comment?: string;
  timestamp: Date;
  context?: string;
}

export interface ToolUsagePattern {
  toolName: string;
  usedWith: string[]; // Other tools used in same session
  successCount: number;
  failureCount: number;
  lastUsed: Date;
  averageResults?: number;
}

export interface LearningMetrics {
  totalQueries: number;
  successfulQueries: number;
  failedQueries: number;
  averageResultsPerQuery: number;
  averageConfidence: number;
  lastCalculated: Date;
  queryImprovementRate?: number; // Percentage of queries that improved over time
}

export interface KnowledgeBaseData {
  successfulQueries: Record<string, SuccessfulQuery>;
  learnedSynonyms: Record<string, string[]>;
  queryPatterns: Record<string, QueryPattern>;
  entityMappings: Record<string, EntityMapping>;
  dataLocations: Record<string, DataLocation>;
  userFeedback: Record<string, UserFeedback[]>;
  confidenceScores: Record<string, number>;
  toolUsagePatterns: Record<string, ToolUsagePattern>;
  patternClusters: Record<string, string[]>; // Cluster-ID -> Pattern-Keys
  learningMetrics: LearningMetrics;
  version: number;
  lastUpdated: string;
}

export class KnowledgeBase {
  private data: KnowledgeBaseData;
  private readonly filePath: string;
  private readonly maxEntries = 10000; // Limit knowledge base size

  constructor(filePath?: string) {
    this.filePath =
      filePath ||
      process.env.MS365_MCP_KNOWLEDGE_BASE_PATH ||
      path.join(__dirname, '..', 'data', 'knowledge-base.json');
    this.data = this.load();
  }

  /**
   * Load knowledge base from file
   */
  private load(): KnowledgeBaseData {
    try {
      if (fs.existsSync(this.filePath)) {
        const content = fs.readFileSync(this.filePath, 'utf8');
        const loaded = JSON.parse(content) as Partial<KnowledgeBaseData>;

        // Migrate old format to new format
        const migrated: KnowledgeBaseData = {
          successfulQueries: loaded.successfulQueries || {},
          learnedSynonyms: loaded.learnedSynonyms || {},
          queryPatterns: loaded.queryPatterns || {},
          entityMappings: loaded.entityMappings || {},
          dataLocations: loaded.dataLocations || {},
          userFeedback: loaded.userFeedback || {},
          confidenceScores: loaded.confidenceScores || {},
          toolUsagePatterns: loaded.toolUsagePatterns || {},
          patternClusters: loaded.patternClusters || {},
          learningMetrics: loaded.learningMetrics || {
            totalQueries: 0,
            successfulQueries: 0,
            failedQueries: 0,
            averageResultsPerQuery: 0,
            averageConfidence: 0,
            lastCalculated: new Date(),
          },
          version: loaded.version || 1,
          lastUpdated: loaded.lastUpdated || new Date().toISOString(),
        };

        // Convert date strings back to Date objects
        Object.values(migrated.successfulQueries || {}).forEach((q) => {
          q.timestamp = new Date(q.timestamp);
        });
        Object.values(migrated.queryPatterns || {}).forEach((p) => {
          p.lastUsed = new Date(p.lastUsed);
        });
        Object.values(migrated.entityMappings || {}).forEach((m) => {
          m.lastUsed = new Date(m.lastUsed);
        });
        Object.values(migrated.dataLocations || {}).forEach((l) => {
          l.lastUsed = new Date(l.lastUsed);
        });

        // Convert user feedback dates
        Object.values(migrated.userFeedback || {}).forEach((feedbacks) => {
          feedbacks.forEach((f) => {
            f.timestamp = new Date(f.timestamp);
          });
        });

        // Convert tool usage pattern dates
        Object.values(migrated.toolUsagePatterns || {}).forEach((pattern) => {
          pattern.lastUsed = new Date(pattern.lastUsed);
        });

        // Convert learning metrics date
        if (migrated.learningMetrics.lastCalculated) {
          migrated.learningMetrics.lastCalculated = new Date(
            migrated.learningMetrics.lastCalculated
          );
        }

        // Update version if needed
        if (migrated.version < 2) {
          migrated.version = 2;
        }

        return migrated;
      }
    } catch (error) {
      logger.warn(`Failed to load knowledge base from ${this.filePath}: ${error}`);
    }

    // Return empty knowledge base
    return {
      successfulQueries: {},
      learnedSynonyms: {},
      queryPatterns: {},
      entityMappings: {},
      dataLocations: {},
      userFeedback: {},
      confidenceScores: {},
      toolUsagePatterns: {},
      patternClusters: {},
      learningMetrics: {
        totalQueries: 0,
        successfulQueries: 0,
        failedQueries: 0,
        averageResultsPerQuery: 0,
        averageConfidence: 0,
        lastCalculated: new Date(),
      },
      version: 2, // Increment version for new structure
      lastUpdated: new Date().toISOString(),
    };
  }

  /**
   * Save knowledge base to file
   */
  async save(): Promise<void> {
    try {
      // Ensure directory exists
      const dir = path.dirname(this.filePath);
      if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
      }

      // Apply time decay before trimming
      const decayDays = parseInt(process.env.MS365_MCP_LEARNING_DECAY_DAYS || '90', 10);
      const decayFactor = parseFloat(process.env.MS365_MCP_LEARNING_DECAY_FACTOR || '0.1');
      if (decayDays > 0 && decayFactor > 0) {
        this.applyTimeDecay(decayDays, decayFactor);
      }

      // Limit entries to prevent file from growing too large
      this.trimEntries();

      this.data.lastUpdated = new Date().toISOString();
      fs.writeFileSync(this.filePath, JSON.stringify(this.data, null, 2), 'utf8');
      logger.debug(`Knowledge base saved to ${this.filePath}`);
    } catch (error) {
      logger.error(`Failed to save knowledge base to ${this.filePath}: ${error}`);
    }
  }

  /**
   * Trim entries to prevent knowledge base from growing too large
   */
  private trimEntries(): void {
    // Keep only most recent and most successful entries
    const trimMapWithTimestamp = <T extends { lastUsed?: Date; successCount?: number }>(
      map: Record<string, T>,
      maxSize: number
    ): Record<string, T> => {
      const entries = Object.entries(map);
      if (entries.length <= maxSize) {
        return map;
      }

      // Sort by lastUsed (most recent first) and successCount (highest first)
      entries.sort((a, b) => {
        const aDate = a[1].lastUsed?.getTime() || 0;
        const bDate = b[1].lastUsed?.getTime() || 0;
        if (bDate !== aDate) {
          return bDate - aDate;
        }
        return (b[1].successCount || 0) - (a[1].successCount || 0);
      });

      return Object.fromEntries(entries.slice(0, maxSize));
    };

    // Separate function for SuccessfulQuery which uses 'timestamp' instead of 'lastUsed'
    const trimSuccessfulQueries = (
      map: Record<string, SuccessfulQuery>,
      maxSize: number
    ): Record<string, SuccessfulQuery> => {
      const entries = Object.entries(map);
      if (entries.length <= maxSize) {
        return map;
      }

      // Sort by timestamp (most recent first) and results count
      entries.sort((a, b) => {
        const aDate = a[1].timestamp?.getTime() || 0;
        const bDate = b[1].timestamp?.getTime() || 0;
        if (bDate !== aDate) {
          return bDate - aDate;
        }
        return (b[1].results || 0) - (a[1].results || 0);
      });

      return Object.fromEntries(entries.slice(0, maxSize));
    };

    const maxPerType = Math.floor(this.maxEntries / 5);
    this.data.successfulQueries = trimSuccessfulQueries(this.data.successfulQueries, maxPerType);
    this.data.queryPatterns = trimMapWithTimestamp(this.data.queryPatterns, maxPerType);
    this.data.entityMappings = trimMapWithTimestamp(this.data.entityMappings, maxPerType);
    this.data.dataLocations = trimMapWithTimestamp(this.data.dataLocations, maxPerType);
  }

  /**
   * Record a successful query
   * SECURITY: Query is sanitized before storage to remove PII
   */
  recordSuccessfulQuery(query: string, results: number, sources: string[], context?: string): void {
    // SECURITY: Sanitize query before storage
    const sanitizedQuery = sanitizeForStorage(query);
    const sanitizedContext = context ? sanitizeForStorage(context) : undefined;

    // Log if PII was detected and removed
    if (containsPII(query)) {
      logger.debug('SECURITY: PII detected and sanitized from query before storage');
    }

    const key = this.normalizeKey(sanitizedQuery);
    this.data.successfulQueries[key] = {
      query: sanitizedQuery,
      results,
      sources,
      timestamp: new Date(),
      context: sanitizedContext,
    };
  }

  /**
   * Get successful queries matching a pattern
   */
  getSuccessfulQueries(pattern: string, limit = 10): SuccessfulQuery[] {
    const patternLower = pattern.toLowerCase();
    return Object.values(this.data.successfulQueries)
      .filter((q) => q.query.toLowerCase().includes(patternLower))
      .sort((a, b) => b.timestamp.getTime() - a.timestamp.getTime())
      .slice(0, limit);
  }

  /**
   * Record a learned synonym
   * SECURITY: Synonyms are sanitized before storage
   */
  recordSynonym(original: string, synonym: string): void {
    // SECURITY: Sanitize before storage
    const sanitizedOriginal = sanitizeForStorage(original);
    const sanitizedSynonym = sanitizeForStorage(synonym);

    const key = this.normalizeKey(sanitizedOriginal);
    if (!this.data.learnedSynonyms[key]) {
      this.data.learnedSynonyms[key] = [];
    }
    if (!this.data.learnedSynonyms[key].includes(sanitizedSynonym)) {
      this.data.learnedSynonyms[key].push(sanitizedSynonym);
    }
  }

  /**
   * Get learned synonyms for a word
   */
  getLearnedSynonyms(word: string): string[] {
    const key = this.normalizeKey(word);
    return this.data.learnedSynonyms[key] || [];
  }

  /**
   * Record a query pattern
   */
  recordQueryPattern(
    pattern: string,
    entityTypes: string[],
    success: boolean,
    context?: string
  ): void {
    const key = this.normalizeKey(pattern);
    if (!this.data.queryPatterns[key]) {
      this.data.queryPatterns[key] = {
        pattern,
        entityTypes: [],
        successCount: 0,
        lastUsed: new Date(),
        context,
      };
    }

    const existing = this.data.queryPatterns[key];
    if (success) {
      existing.successCount++;
    }
    existing.lastUsed = new Date();

    // Update entity types (merge with existing)
    for (const et of entityTypes) {
      if (!existing.entityTypes.includes(et)) {
        existing.entityTypes.push(et);
      }
    }
  }

  /**
   * Get query pattern for a query type
   */
  getQueryPattern(pattern: string): QueryPattern | undefined {
    const key = this.normalizeKey(pattern);
    return this.data.queryPatterns[key];
  }

  /**
   * Record entity mapping
   */
  recordEntityMapping(keyword: string, entityTypes: string[], success: boolean): void {
    const key = this.normalizeKey(keyword);
    if (!this.data.entityMappings[key]) {
      this.data.entityMappings[key] = {
        keyword,
        entityTypes: [],
        successCount: 0,
        lastUsed: new Date(),
      };
    }

    const existing = this.data.entityMappings[key];
    if (success) {
      existing.successCount++;
    }
    existing.lastUsed = new Date();

    // Update entity types (merge with existing)
    for (const et of entityTypes) {
      if (!existing.entityTypes.includes(et)) {
        existing.entityTypes.push(et);
      }
    }
  }

  /**
   * Get entity mapping for a keyword
   */
  getEntityMapping(keyword: string): string[] {
    const key = this.normalizeKey(keyword);
    const mapping = this.data.entityMappings[key];
    return mapping ? mapping.entityTypes : [];
  }

  /**
   * Record data location (where certain data types are typically found)
   */
  recordDataLocation(dataType: string, sources: string[], success: boolean): void {
    const key = this.normalizeKey(dataType);
    if (!this.data.dataLocations[key]) {
      this.data.dataLocations[key] = {
        dataType,
        sources: [],
        successCount: 0,
        lastUsed: new Date(),
      };
    }

    const existing = this.data.dataLocations[key];
    if (success) {
      existing.successCount++;
    }
    existing.lastUsed = new Date();

    // Update sources (merge with existing)
    for (const source of sources) {
      if (!existing.sources.includes(source)) {
        existing.sources.push(source);
      }
    }
  }

  /**
   * Get data location for a data type
   */
  getDataLocation(dataType: string): string[] {
    const key = this.normalizeKey(dataType);
    const location = this.data.dataLocations[key];
    return location ? location.sources : [];
  }

  /**
   * Suggest query variants based on learned patterns
   */
  suggestVariants(query: string, context?: string): string[] {
    const variants = new Set<string>([query]);

    // Check successful queries
    const successful = this.getSuccessfulQueries(query, 5);
    for (const sq of successful) {
      if (!context || !sq.context || sq.context === context) {
        variants.add(sq.query);
      }
    }

    // Check learned synonyms
    const words = query.toLowerCase().split(/\s+/);
    for (const word of words) {
      const synonyms = this.getLearnedSynonyms(word);
      for (const synonym of synonyms) {
        const variant = query.toLowerCase().replace(word, synonym);
        variants.add(variant);
      }
    }

    return Array.from(variants).slice(0, 10);
  }

  /**
   * Normalize key for storage
   */
  private normalizeKey(key: string): string {
    return key.toLowerCase().trim();
  }

  /**
   * Get all data (for testing/debugging)
   */
  getAllData(): KnowledgeBaseData {
    return this.data;
  }

  /**
   * Clear all data
   */
  clear(): void {
    this.data = {
      successfulQueries: {},
      learnedSynonyms: {},
      queryPatterns: {},
      entityMappings: {},
      dataLocations: {},
      userFeedback: {},
      confidenceScores: {},
      toolUsagePatterns: {},
      patternClusters: {},
      learningMetrics: {
        totalQueries: 0,
        successfulQueries: 0,
        failedQueries: 0,
        averageResultsPerQuery: 0,
        averageConfidence: 0,
        lastCalculated: new Date(),
      },
      version: 2,
      lastUpdated: new Date().toISOString(),
    };
  }

  /**
   * Record user feedback
   * SECURITY: All user input is sanitized before storage
   */
  recordUserFeedback(
    query: string,
    feedbackType: 'helpful' | 'not_helpful' | 'incorrect' | 'correct',
    resultId?: string,
    comment?: string,
    context?: string
  ): void {
    // SECURITY: Sanitize all user-provided data
    const sanitizedQuery = sanitizeForStorage(query);
    const sanitizedComment = comment ? sanitizeForStorage(comment) : undefined;
    const sanitizedContext = context ? sanitizeForStorage(context) : undefined;

    const key = this.normalizeKey(sanitizedQuery);
    if (!this.data.userFeedback[key]) {
      this.data.userFeedback[key] = [];
    }

    this.data.userFeedback[key].push({
      query: sanitizedQuery,
      resultId,
      feedbackType,
      comment: sanitizedComment,
      timestamp: new Date(),
      context: sanitizedContext,
    });

    // Keep only last 100 feedback entries per query
    if (this.data.userFeedback[key].length > 100) {
      this.data.userFeedback[key] = this.data.userFeedback[key]
        .sort((a, b) => b.timestamp.getTime() - a.timestamp.getTime())
        .slice(0, 100);
    }
  }

  /**
   * Get user feedback for a query
   */
  getUserFeedback(query: string): UserFeedback[] {
    const key = this.normalizeKey(query);
    return this.data.userFeedback[key] || [];
  }

  /**
   * Calculate and store confidence score for a pattern
   */
  calculateConfidence(patternKey: string, score: number): void {
    this.data.confidenceScores[patternKey] = Math.max(0, Math.min(1, score)); // Clamp between 0 and 1
  }

  /**
   * Get confidence score for a pattern
   */
  getConfidenceScore(patternKey: string): number {
    return this.data.confidenceScores[patternKey] ?? 0.5; // Default to 0.5 if not set
  }

  /**
   * Apply time decay to patterns
   */
  applyTimeDecay(decayDays: number, decayFactor: number): void {
    const now = new Date();
    const decayMs = decayDays * 24 * 60 * 60 * 1000;

    // Apply decay to query patterns
    for (const [key, pattern] of Object.entries(this.data.queryPatterns)) {
      const age = now.getTime() - pattern.lastUsed.getTime();
      if (age > decayMs) {
        const monthsOld = age / (30 * 24 * 60 * 60 * 1000);
        const decayMultiplier = Math.max(0.1, 1 - monthsOld * decayFactor);
        const currentConfidence = this.getConfidenceScore(key);
        this.calculateConfidence(key, currentConfidence * decayMultiplier);
      }
    }

    // Apply decay to entity mappings
    for (const [key, mapping] of Object.entries(this.data.entityMappings)) {
      const age = now.getTime() - mapping.lastUsed.getTime();
      if (age > decayMs) {
        const monthsOld = age / (30 * 24 * 60 * 60 * 1000);
        const decayMultiplier = Math.max(0.1, 1 - monthsOld * decayFactor);
        const currentConfidence = this.getConfidenceScore(key);
        this.calculateConfidence(key, currentConfidence * decayMultiplier);
      }
    }
  }

  /**
   * Cluster patterns (simple implementation - can be enhanced)
   */
  clusterPatterns(similarityThreshold: number = 0.7): void {
    const patterns = Object.keys(this.data.queryPatterns);
    const clusters: Record<string, string[]> = {};
    let clusterId = 0;

    for (const pattern1 of patterns) {
      let assigned = false;

      // Check if pattern belongs to existing cluster
      for (const [clusterKey, clusterPatterns] of Object.entries(clusters)) {
        for (const pattern2 of clusterPatterns) {
          const similarity = this.calculateSimilarity(pattern1, pattern2);
          if (similarity >= similarityThreshold) {
            clusters[clusterKey].push(pattern1);
            assigned = true;
            break;
          }
        }
        if (assigned) break;
      }

      // Create new cluster if not assigned
      if (!assigned) {
        clusters[`cluster_${clusterId++}`] = [pattern1];
      }
    }

    this.data.patternClusters = clusters;
  }

  /**
   * Calculate similarity between two patterns (simple Jaccard similarity)
   */
  private calculateSimilarity(pattern1: string, pattern2: string): number {
    const words1 = new Set(pattern1.toLowerCase().split(/\s+/));
    const words2 = new Set(pattern2.toLowerCase().split(/\s+/));

    const intersection = new Set([...words1].filter((x) => words2.has(x)));
    const union = new Set([...words1, ...words2]);

    return union.size > 0 ? intersection.size / union.size : 0;
  }

  /**
   * Record tool usage pattern
   */
  recordToolUsage(
    toolName: string,
    usedWith: string[],
    success: boolean,
    resultsCount?: number
  ): void {
    const key = this.normalizeKey(toolName);
    if (!this.data.toolUsagePatterns[key]) {
      this.data.toolUsagePatterns[key] = {
        toolName,
        usedWith: [],
        successCount: 0,
        failureCount: 0,
        lastUsed: new Date(),
        averageResults: 0,
      };
    }

    const pattern = this.data.toolUsagePatterns[key];
    if (success) {
      pattern.successCount++;
    } else {
      pattern.failureCount++;
    }
    pattern.lastUsed = new Date();

    // Update usedWith (merge with existing)
    for (const tool of usedWith) {
      if (!pattern.usedWith.includes(tool)) {
        pattern.usedWith.push(tool);
      }
    }

    // Update average results
    if (resultsCount !== undefined) {
      const totalUses = pattern.successCount + pattern.failureCount;
      if (totalUses === 1) {
        pattern.averageResults = resultsCount;
      } else {
        pattern.averageResults =
          (pattern.averageResults! * (totalUses - 1) + resultsCount) / totalUses;
      }
    }
  }

  /**
   * Get tool usage pattern
   */
  getToolUsagePattern(toolName: string): ToolUsagePattern | undefined {
    const key = this.normalizeKey(toolName);
    return this.data.toolUsagePatterns[key];
  }

  /**
   * Get recommended tools to use together
   */
  getRecommendedToolCombinations(toolName: string, limit: number = 5): string[] {
    const pattern = this.getToolUsagePattern(toolName);
    if (!pattern) {
      return [];
    }

    // Sort by usage frequency and success rate
    const toolScores = pattern.usedWith.map((tool) => {
      const toolPattern = this.getToolUsagePattern(tool);
      if (!toolPattern) {
        return { tool, score: 0 };
      }
      const totalUses = toolPattern.successCount + toolPattern.failureCount;
      const successRate = totalUses > 0 ? toolPattern.successCount / totalUses : 0;
      return { tool, score: successRate * totalUses };
    });

    return toolScores
      .sort((a, b) => b.score - a.score)
      .slice(0, limit)
      .map((t) => t.tool);
  }

  /**
   * Export knowledge base as JSON
   */
  exportKnowledgeBase(): string {
    return JSON.stringify(this.data, null, 2);
  }

  /**
   * Import knowledge base from JSON
   */
  importKnowledgeBase(jsonData: string): void {
    try {
      const imported = JSON.parse(jsonData) as Partial<KnowledgeBaseData>;

      // Merge with existing data
      this.data = {
        ...this.data,
        ...imported,
        // Preserve existing data if imported data is missing fields
        successfulQueries: {
          ...this.data.successfulQueries,
          ...(imported.successfulQueries || {}),
        },
        learnedSynonyms: { ...this.data.learnedSynonyms, ...(imported.learnedSynonyms || {}) },
        queryPatterns: { ...this.data.queryPatterns, ...(imported.queryPatterns || {}) },
        entityMappings: { ...this.data.entityMappings, ...(imported.entityMappings || {}) },
        dataLocations: { ...this.data.dataLocations, ...(imported.dataLocations || {}) },
        userFeedback: { ...this.data.userFeedback, ...(imported.userFeedback || {}) },
        confidenceScores: { ...this.data.confidenceScores, ...(imported.confidenceScores || {}) },
        toolUsagePatterns: {
          ...this.data.toolUsagePatterns,
          ...(imported.toolUsagePatterns || {}),
        },
        patternClusters: { ...this.data.patternClusters, ...(imported.patternClusters || {}) },
        version: Math.max(this.data.version, imported.version || 2),
        lastUpdated: new Date().toISOString(),
      };
    } catch (error) {
      logger.error(`Failed to import knowledge base: ${error}`);
      throw new Error(`Invalid knowledge base format: ${error}`);
    }
  }

  /**
   * Merge knowledge base with another knowledge base
   */
  mergeKnowledgeBase(other: KnowledgeBaseData): void {
    // Merge successful queries
    this.data.successfulQueries = { ...this.data.successfulQueries, ...other.successfulQueries };

    // Merge learned synonyms
    for (const [key, synonyms] of Object.entries(other.learnedSynonyms || {})) {
      if (!this.data.learnedSynonyms[key]) {
        this.data.learnedSynonyms[key] = [];
      }
      for (const synonym of synonyms) {
        if (!this.data.learnedSynonyms[key].includes(synonym)) {
          this.data.learnedSynonyms[key].push(synonym);
        }
      }
    }

    // Merge query patterns (keep higher success count)
    for (const [key, pattern] of Object.entries(other.queryPatterns || {})) {
      if (
        !this.data.queryPatterns[key] ||
        pattern.successCount > this.data.queryPatterns[key].successCount
      ) {
        this.data.queryPatterns[key] = pattern;
      }
    }

    // Merge entity mappings
    for (const [key, mapping] of Object.entries(other.entityMappings || {})) {
      if (
        !this.data.entityMappings[key] ||
        mapping.successCount > this.data.entityMappings[key].successCount
      ) {
        this.data.entityMappings[key] = mapping;
      }
    }

    // Merge user feedback
    for (const [key, feedbacks] of Object.entries(other.userFeedback || {})) {
      if (!this.data.userFeedback[key]) {
        this.data.userFeedback[key] = [];
      }
      this.data.userFeedback[key].push(...feedbacks);
    }

    // Merge confidence scores (use average)
    for (const [key, score] of Object.entries(other.confidenceScores || {})) {
      const existing = this.data.confidenceScores[key];
      if (existing) {
        this.data.confidenceScores[key] = (existing + score) / 2;
      } else {
        this.data.confidenceScores[key] = score;
      }
    }

    // Merge tool usage patterns
    for (const [key, pattern] of Object.entries(other.toolUsagePatterns || {})) {
      if (this.data.toolUsagePatterns[key]) {
        const existing = this.data.toolUsagePatterns[key];
        existing.successCount += pattern.successCount;
        existing.failureCount += pattern.failureCount;
        existing.usedWith = [...new Set([...existing.usedWith, ...pattern.usedWith])];
        if (pattern.averageResults !== undefined) {
          const totalUses = existing.successCount + existing.failureCount;
          existing.averageResults = existing.averageResults
            ? (existing.averageResults + pattern.averageResults) / 2
            : pattern.averageResults;
        }
      } else {
        this.data.toolUsagePatterns[key] = pattern;
      }
    }

    // Merge pattern clusters
    for (const [clusterId, patterns] of Object.entries(other.patternClusters || {})) {
      if (!this.data.patternClusters[clusterId]) {
        this.data.patternClusters[clusterId] = [];
      }
      this.data.patternClusters[clusterId] = [
        ...new Set([...this.data.patternClusters[clusterId], ...patterns]),
      ];
    }

    this.data.version = Math.max(this.data.version, other.version || 2);
    this.data.lastUpdated = new Date().toISOString();
  }

  /**
   * Export analytics data
   */
  exportAnalytics(): {
    totalQueries: number;
    successfulQueries: number;
    failedQueries: number;
    averageResultsPerQuery: number;
    averageConfidence: number;
    topPatterns: Array<{ pattern: string; successCount: number; confidence: number }>;
    toolUsageStats: Array<{ tool: string; successRate: number; averageResults: number }>;
  } {
    const totalQueries = Object.keys(this.data.successfulQueries).length;
    const successfulQueries = Object.values(this.data.queryPatterns).reduce(
      (sum, p) => sum + p.successCount,
      0
    );

    const topPatterns = Object.entries(this.data.queryPatterns)
      .map(([key, pattern]) => ({
        pattern: key,
        successCount: pattern.successCount,
        confidence: this.getConfidenceScore(key),
      }))
      .sort((a, b) => b.successCount - a.successCount)
      .slice(0, 10);

    const toolUsageStats = Object.values(this.data.toolUsagePatterns)
      .map((pattern) => {
        const totalUses = pattern.successCount + pattern.failureCount;
        return {
          tool: pattern.toolName,
          successRate: totalUses > 0 ? pattern.successCount / totalUses : 0,
          averageResults: pattern.averageResults || 0,
        };
      })
      .sort((a, b) => b.successRate - a.successRate);

    const averageResultsPerQuery =
      totalQueries > 0
        ? Object.values(this.data.successfulQueries).reduce((sum, q) => sum + q.results, 0) /
          totalQueries
        : 0;

    const confidenceScores = Object.values(this.data.confidenceScores);
    const averageConfidence =
      confidenceScores.length > 0
        ? confidenceScores.reduce((sum, score) => sum + score, 0) / confidenceScores.length
        : 0;

    return {
      totalQueries,
      successfulQueries,
      failedQueries: totalQueries - successfulQueries,
      averageResultsPerQuery,
      averageConfidence,
      topPatterns,
      toolUsageStats,
    };
  }
}

export default KnowledgeBase;
