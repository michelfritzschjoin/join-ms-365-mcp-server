/**
 * Self-Learning System for improving search queries and data location discovery
 */

import logger from './logger.js';
import KnowledgeBase, { type KnowledgeBaseData } from './knowledge-base.js';
import SynonymExpander from './synonym-expander.js';
import type { RepairHistoryEntry } from './graph-api-repair.js';
import LearningAnalytics, { type PerformanceMetrics } from './learning-analytics.js';
import { getQueryStore, type QueryPattern as StoredQueryPattern } from './query-store.js';

export interface SearchResult {
  items: unknown[];
  sources: string[];
  query: string;
  entityTypes?: string[];
  totalResults?: number;
}

export interface QueryPattern {
  queryType: string;
  entityTypes: string[];
  sources: string[];
  successRate: number;
}

export class LearningSystem {
  private knowledgeBase: KnowledgeBase;
  private synonymExpander: SynonymExpander;
  private analytics: LearningAnalytics;
  private readonly learningEnabled: boolean;

  constructor(knowledgeBase: KnowledgeBase, synonymExpander: SynonymExpander) {
    this.knowledgeBase = knowledgeBase;
    this.synonymExpander = synonymExpander;
    this.analytics = new LearningAnalytics(knowledgeBase);
    this.learningEnabled =
      process.env.MS365_MCP_LEARNING_ENABLED !== 'false' &&
      process.env.MS365_MCP_LEARNING_ENABLED !== '0';
  }

  /**
   * Learn from a search operation
   */
  async learnFromSearch(
    query: string,
    results: SearchResult,
    userFeedback?: 'success' | 'failure',
    context?: string
  ): Promise<void> {
    if (!this.learningEnabled) {
      return;
    }

    try {
      const success = results.items.length > 0 || userFeedback === 'success';
      const failure = results.items.length === 0 || userFeedback === 'failure';

      // User feedback is weighted higher - if provided, use it directly
      const weight = userFeedback ? 2 : 1; // User feedback counts double

      // 1. Record successful query pattern
      if (success && (results.items.length > 0 || userFeedback === 'success')) {
        this.knowledgeBase.recordSuccessfulQuery(
          query,
          results.items.length,
          results.sources,
          context
        );

        // Record query pattern if entity types are known
        if (results.entityTypes && results.entityTypes.length > 0) {
          // Record multiple times if user feedback provided (weighted learning)
          for (let i = 0; i < weight; i++) {
            this.knowledgeBase.recordQueryPattern(query, results.entityTypes, true, context);
          }
        }
      }

      // 2. Learn synonyms from context
      if (success) {
        this.learnSynonymsFromResults(query, results);
      }

      // 3. Learn data locations (where certain data types are typically found)
      if (success && results.sources.length > 0) {
        this.learnDataLocations(query, results);
      }

      // 4. Update entity mappings (weighted if user feedback)
      if (success && results.entityTypes) {
        for (const et of results.entityTypes) {
          for (let i = 0; i < weight; i++) {
            this.knowledgeBase.recordEntityMapping(query, [et], true);
          }
        }
      }

      // 5. Apply time decay before saving
      this.applyTimeDecayIfEnabled();

      // 6. Persist knowledge (async, don't wait)
      this.knowledgeBase.save().catch((error) => {
        logger.warn(`Failed to persist knowledge base: ${error}`);
      });

      logger.debug(
        `Learning system: recorded ${success ? 'successful' : 'failed'} search for "${query}"${userFeedback ? ` (user feedback: ${userFeedback})` : ''}`
      );
    } catch (error) {
      logger.warn(`Learning system error: ${error}`);
    }
  }

  /**
   * Record explicit user feedback
   * This is weighted higher than implicit learning from search results
   */
  async recordUserFeedback(
    query: string,
    feedbackType: 'helpful' | 'not_helpful' | 'incorrect' | 'correct',
    resultId?: string,
    comment?: string,
    context?: string
  ): Promise<void> {
    if (!this.learningEnabled) {
      return;
    }

    try {
      // Record feedback in knowledge base
      this.knowledgeBase.recordUserFeedback(query, feedbackType, resultId, comment, context);

      // If feedback is positive, strengthen the pattern
      if (feedbackType === 'helpful' || feedbackType === 'correct') {
        // Get existing patterns for this query
        const pattern = this.knowledgeBase.getQueryPattern(query);
        if (pattern) {
          // Boost confidence and success count
          for (let i = 0; i < 3; i++) {
            // Record multiple times to boost weight
            this.knowledgeBase.recordQueryPattern(
              query,
              pattern.entityTypes,
              true,
              context || pattern.context
            );
          }

          // Increase confidence score
          const currentConfidence = this.knowledgeBase.getConfidenceScore(query);
          const newConfidence = Math.min(1.0, currentConfidence + 0.1);
          this.knowledgeBase.calculateConfidence(query, newConfidence);
        }
      } else if (feedbackType === 'not_helpful' || feedbackType === 'incorrect') {
        // Negative feedback - decrease confidence
        const currentConfidence = this.knowledgeBase.getConfidenceScore(query);
        const newConfidence = Math.max(0.0, currentConfidence - 0.1);
        this.knowledgeBase.calculateConfidence(query, newConfidence);
      }

      // Persist knowledge (async, don't wait)
      this.knowledgeBase.save().catch((error) => {
        logger.warn(`Failed to persist user feedback: ${error}`);
      });

      logger.info(`User feedback recorded: ${feedbackType} for query "${query}"`);
    } catch (error) {
      logger.warn(`Failed to record user feedback: ${error}`);
    }
  }

  /**
   * Learn synonyms from search results
   */
  private learnSynonymsFromResults(query: string, results: SearchResult): void {
    // Extract potential synonyms from result items
    // This is a simple heuristic - could be improved with NLP
    const queryWords = query.toLowerCase().split(/\s+/);

    // If results contain items with similar but different terms, learn them
    for (const item of results.items.slice(0, 10)) {
      // Try to extract text from item (simplified)
      const itemText = this.extractTextFromItem(item);
      if (itemText) {
        const itemWords = itemText.toLowerCase().split(/\s+/);
        // Find words that appear in results but not in query
        for (const word of itemWords) {
          if (word.length > 3 && !queryWords.includes(word)) {
            // Potential synonym - learn it with low confidence
            this.synonymExpander.learnSynonym(queryWords[0] || query, word, true);
          }
        }
      }
    }
  }

  /**
   * Learn data locations (where certain data types are typically found)
   */
  private learnDataLocations(query: string, results: SearchResult): void {
    // Infer data type from query
    const dataType = this.inferDataType(query);

    if (dataType && results.sources.length > 0) {
      const success = results.items.length > 0;
      this.knowledgeBase.recordDataLocation(dataType, results.sources, success);
    }
  }

  /**
   * Infer data type from query
   */
  private inferDataType(query: string): string | undefined {
    const queryLower = query.toLowerCase();

    if (queryLower.includes('projekt') || queryLower.includes('project')) {
      return 'project';
    }
    if (queryLower.includes('meeting') || queryLower.includes('besprechung')) {
      return 'meeting';
    }
    if (queryLower.includes('dokument') || queryLower.includes('document')) {
      return 'document';
    }
    if (queryLower.includes('kunde') || queryLower.includes('customer')) {
      return 'customer';
    }
    if (queryLower.includes('vertrag') || queryLower.includes('contract')) {
      return 'contract';
    }
    if (queryLower.includes('team') || queryLower.includes('gruppe')) {
      return 'team';
    }
    if (queryLower.includes('person') || queryLower.includes('user')) {
      return 'person';
    }

    return undefined;
  }

  /**
   * Extract text from result item (simplified)
   */
  private extractTextFromItem(item: unknown): string {
    if (typeof item === 'string') {
      return item;
    }
    if (typeof item === 'object' && item !== null) {
      const obj = item as Record<string, unknown>;
      // Try common text fields
      const textFields = ['name', 'title', 'subject', 'displayName', 'content', 'body'];
      for (const field of textFields) {
        if (typeof obj[field] === 'string') {
          return obj[field] as string;
        }
      }
      // Fallback: stringify and take first 200 chars
      return JSON.stringify(item).substring(0, 200);
    }
    return '';
  }

  /**
   * Suggest better queries based on learned patterns
   */
  async suggestQuery(query: string, context?: string): Promise<string[]> {
    if (!this.learningEnabled) {
      return [query];
    }

    // Get suggestions from knowledge base
    const suggestions = this.knowledgeBase.suggestVariants(query, context);

    // Also expand with synonyms
    const synonymVariants = this.synonymExpander.expandQuery(query, context);

    // Combine and deduplicate
    const allVariants = new Set([query, ...suggestions, ...synonymVariants]);
    return Array.from(allVariants).slice(0, 10);
  }

  /**
   * Get suggested query variants from past successful searches (learned patterns).
   * Returns queries that contained similar terms and previously returned results.
   */
  getSuggestedQueryVariants(query: string, limit = 5): string[] {
    if (!this.learningEnabled || !query || query.trim().length === 0) {
      return [];
    }
    const trimmed = query.trim().toLowerCase();
    const words = trimmed.split(/\s+/).filter((w) => w.length > 2);
    if (words.length === 0) return [];

    const seen = new Set<string>([trimmed]);
    const variants: string[] = [];
    const pattern = words.slice(0, 2).join(' ');
    const successful = this.knowledgeBase.getSuccessfulQueries(pattern, limit);
    for (const s of successful) {
      const q = s.query.trim().toLowerCase();
      if (q && !seen.has(q) && q !== trimmed) {
        seen.add(q);
        variants.push(s.query.trim());
      }
    }
    return variants.slice(0, limit);
  }

  /**
   * Get recommended entity types for a query
   */
  getRecommendedEntityTypes(query: string, context?: string): string[] {
    // Check entity mappings
    const mappings = this.knowledgeBase.getEntityMapping(query);
    if (mappings.length > 0) {
      return mappings;
    }

    // Check query patterns
    const pattern = this.knowledgeBase.getQueryPattern(query);
    if (pattern && (!context || !pattern.context || pattern.context === context)) {
      return pattern.entityTypes;
    }

    // Default entity types based on context - comprehensive set
    if (context === 'project') {
      return [
        'driveItem',
        'site',
        'listItem',
        'chatMessage',
        'event',
        'message',
        'acronym',
        'bookmark',
      ];
    }
    if (context === 'person') {
      return ['person', 'message', 'event', 'driveItem', 'chatMessage'];
    }
    if (context === 'document') {
      return ['driveItem', 'message'];
    }
    if (context === 'meeting') {
      return ['event', 'message', 'driveItem'];
    }

    // Default: comprehensive search across all available entity types
    return [
      'driveItem',
      'message',
      'event',
      'site',
      'list',
      'listItem',
      'chatMessage',
      'person',
      'acronym',
      'bookmark',
    ];
  }

  /**
   * Get recommended sources for a data type
   */
  getRecommendedSources(dataType: string): string[] {
    const locations = this.knowledgeBase.getDataLocation(dataType);
    if (locations.length > 0) {
      return locations;
    }

    // Default sources based on data type
    const defaultSources: Record<string, string[]> = {
      project: ['sharepoint', 'teams', 'planner', 'files', 'mail', 'calendar'],
      person: ['users', 'mail', 'calendar', 'teams', 'files', 'contacts'],
      document: ['sharepoint', 'onedrive', 'mail', 'teams'],
      meeting: ['calendar', 'teams', 'mail', 'onenote', 'files'],
      customer: ['contacts', 'mail', 'calendar', 'files', 'teams'],
      contract: ['files', 'mail', 'calendar'],
      team: ['teams', 'sharepoint', 'planner', 'mail'],
    };

    return defaultSources[dataType] || ['search', 'mail', 'files'];
  }

  /**
   * Load knowledge base data
   */
  loadKnowledgeBase(): KnowledgeBaseData {
    return this.knowledgeBase.getAllData();
  }

  /**
   * Clear all learned data
   */
  clearLearning(): void {
    this.knowledgeBase.clear();
    this.knowledgeBase.save().catch((error) => {
      logger.warn(`Failed to clear knowledge base: ${error}`);
    });
  }

  /**
   * Learn from repair history
   * Records successful repair patterns for future use
   */
  async learnFromRepair(repairEntry: RepairHistoryEntry): Promise<void> {
    if (!this.learningEnabled) {
      return;
    }

    try {
      // Extract endpoint pattern
      const endpointPattern = this.extractEndpointPattern(repairEntry.endpoint);

      // Record successful repair pattern
      if (repairEntry.success) {
        this.knowledgeBase.recordSuccessfulQuery(
          `repair:${repairEntry.strategy}:${endpointPattern}`,
          1,
          [repairEntry.strategy],
          `error:${repairEntry.error.errorCode || repairEntry.error.statusCode}`
        );

        logger.debug(`Learning system: recorded successful repair`, {
          strategy: repairEntry.strategy,
          endpoint: repairEntry.endpoint,
          errorCode: repairEntry.error.errorCode,
        });
      } else {
        // Record failed repair for analysis
        logger.debug(`Learning system: recorded failed repair`, {
          strategy: repairEntry.strategy,
          endpoint: repairEntry.endpoint,
          errorCode: repairEntry.error.errorCode,
        });
      }

      // Persist knowledge (async, don't wait)
      this.knowledgeBase.save().catch((error) => {
        logger.warn(`Failed to persist repair knowledge: ${error}`);
      });
    } catch (error) {
      logger.warn(`Learning system repair error: ${error}`);
    }
  }

  /**
   * Extract endpoint pattern from full endpoint path
   */
  private extractEndpointPattern(endpoint: string): string {
    // Remove IDs and specific values to get pattern
    // e.g., /users/123/messages -> /users/{id}/messages
    return endpoint
      .replace(/\/[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi, '/{id}')
      .replace(/\/[^/]+-id\//gi, '/{id}/')
      .replace(/\/[^/]+Id\//gi, '/{id}/')
      .replace(/\/\d+\//g, '/{id}/')
      .replace(/\/\d+$/g, '/{id}');
  }

  /**
   * Get recommended repair strategy for an endpoint and error
   */
  getRecommendedRepairStrategy(
    endpoint: string,
    errorCode?: string,
    statusCode?: number
  ): string | null {
    if (!this.learningEnabled) {
      return null;
    }

    const endpointPattern = this.extractEndpointPattern(endpoint);
    const query = `repair:${endpointPattern}`;
    const context = errorCode
      ? `error:${errorCode}`
      : statusCode
        ? `error:${statusCode}`
        : undefined;

    // Check knowledge base for successful repair patterns
    const patterns = this.knowledgeBase.getQueryPattern(query);
    if (patterns && patterns.entityTypes.length > 0) {
      // entityTypes in this context represent repair strategies
      return patterns.entityTypes[0];
    }

    return null;
  }

  /**
   * Get confidence score for a pattern
   */
  getConfidenceScore(patternKey: string): number {
    return this.knowledgeBase.getConfidenceScore(patternKey);
  }

  /**
   * Get performance metrics
   */
  getPerformanceMetrics(): PerformanceMetrics {
    // Update confidence scores before calculating metrics
    this.analytics.updateAllConfidenceScores();
    return this.analytics.calculatePerformanceMetrics();
  }

  /**
   * Update confidence scores for all patterns
   */
  updateConfidenceScores(): void {
    this.analytics.updateAllConfidenceScores();
  }

  /**
   * Apply time decay to patterns if enabled
   */
  private applyTimeDecayIfEnabled(): void {
    const decayDays = parseInt(process.env.MS365_MCP_LEARNING_DECAY_DAYS || '90', 10);
    const decayFactor = parseFloat(process.env.MS365_MCP_LEARNING_DECAY_FACTOR || '0.1');

    if (decayDays > 0 && decayFactor > 0) {
      this.knowledgeBase.applyTimeDecay(decayDays, decayFactor);
    }
  }

  /**
   * Learn from tool usage patterns
   */
  learnFromToolUsage(
    toolName: string,
    usedWith: string[],
    success: boolean,
    resultsCount?: number
  ): void {
    if (!this.learningEnabled) {
      return;
    }

    try {
      this.knowledgeBase.recordToolUsage(toolName, usedWith, success, resultsCount);

      // Persist knowledge (async, don't wait)
      this.knowledgeBase.save().catch((error) => {
        logger.warn(`Failed to persist tool usage: ${error}`);
      });

      logger.debug(`Learned from tool usage: ${toolName} with ${usedWith.length} other tools`);
    } catch (error) {
      logger.warn(`Failed to learn from tool usage: ${error}`);
    }
  }

  /**
   * Get recommended tool combinations
   */
  getRecommendedToolCombinations(toolName: string, limit: number = 5): string[] {
    return this.knowledgeBase.getRecommendedToolCombinations(toolName, limit);
  }

  // =========================================================================
  // QUERY HISTORY LEARNING (USER-SPECIFIC)
  // =========================================================================

  /**
   * Learn from query history for a specific user
   * Analyzes stored queries to improve future recommendations
   * @param userIdHash - Hashed user ID for isolation
   * @returns Summary of learned patterns
   */
  async learnFromQueryHistory(userIdHash: string): Promise<{
    patternsAnalyzed: number;
    newPatternsLearned: number;
    recommendations: string[];
  }> {
    if (!this.learningEnabled || !userIdHash) {
      return { patternsAnalyzed: 0, newPatternsLearned: 0, recommendations: [] };
    }

    try {
      const queryStore = getQueryStore();
      const patterns = queryStore.getQueryPatterns(userIdHash, 50);
      const recommendations: string[] = [];
      let newPatternsLearned = 0;

      // Analyze patterns and transfer to knowledge base
      for (const pattern of patterns) {
        // Skip patterns with low success rate
        if (pattern.successRate < 0.3) {
          continue;
        }

        // Record in knowledge base for cross-pattern learning
        if (pattern.optimalEntityTypes.length > 0) {
          this.knowledgeBase.recordQueryPattern(
            pattern.pattern,
            pattern.optimalEntityTypes,
            pattern.successRate >= 0.5,
            `user-history:${userIdHash.substring(0, 8)}`
          );
          newPatternsLearned++;
        }

        // Generate recommendations based on patterns
        if (pattern.successRate >= 0.8 && pattern.count >= 5) {
          recommendations.push(
            `Pattern "${pattern.pattern}" has ${Math.round(pattern.successRate * 100)}% success rate with entity types: ${pattern.optimalEntityTypes.join(', ')}`
          );
        }

        // Learn from slow queries
        if (pattern.avgDuration > 3000 && pattern.count >= 3) {
          recommendations.push(
            `Pattern "${pattern.pattern}" is slow (avg ${pattern.avgDuration}ms). Consider optimizing entity types.`
          );
        }
      }

      // Persist learned data
      await this.knowledgeBase.save();

      logger.debug('Learned from query history', {
        userIdHash: userIdHash.substring(0, 8) + '...',
        patternsAnalyzed: patterns.length,
        newPatternsLearned,
        recommendations: recommendations.length,
      });

      return {
        patternsAnalyzed: patterns.length,
        newPatternsLearned,
        recommendations: recommendations.slice(0, 10),
      };
    } catch (error) {
      logger.warn(`Failed to learn from query history: ${error}`);
      return { patternsAnalyzed: 0, newPatternsLearned: 0, recommendations: [] };
    }
  }

  /**
   * Get recommended entity types based on user history
   * @param query - Query string
   * @param userIdHash - Hashed user ID for isolation
   * @returns Recommended entity types or null
   */
  getHistoryBasedEntityTypes(
    query: string,
    userIdHash: string
  ): { entityTypes: string[]; confidence: number; reason: string } | null {
    if (!this.learningEnabled || !userIdHash || !query) {
      return null;
    }

    try {
      const queryStore = getQueryStore();
      return queryStore.getOptimalEntityTypes(query, userIdHash);
    } catch (error) {
      logger.warn(`Failed to get history-based entity types: ${error}`);
      return null;
    }
  }

  /**
   * Record a search query for future learning
   * @param userIdHash - Hashed user ID
   * @param query - Query string
   * @param entityTypes - Entity types used
   * @param success - Whether results were found
   * @param duration - Duration in milliseconds
   */
  recordSearchForLearning(
    userIdHash: string,
    query: string,
    entityTypes: string[],
    success: boolean,
    duration: number
  ): void {
    if (!this.learningEnabled || !userIdHash || !query) {
      return;
    }

    try {
      const queryStore = getQueryStore();
      queryStore.recordQueryPattern(userIdHash, query, entityTypes, success, duration);
    } catch (error) {
      logger.warn(`Failed to record search for learning: ${error}`);
    }
  }
}

export default LearningSystem;
