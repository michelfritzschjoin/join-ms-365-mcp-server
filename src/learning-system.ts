/**
 * Self-Learning System for improving search queries and data location discovery
 */

import logger from './logger.js';
import KnowledgeBase, { type KnowledgeBaseData } from './knowledge-base.js';
import SynonymExpander from './synonym-expander.js';

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
  private readonly learningEnabled: boolean;

  constructor(knowledgeBase: KnowledgeBase, synonymExpander: SynonymExpander) {
    this.knowledgeBase = knowledgeBase;
    this.synonymExpander = synonymExpander;
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

      // 1. Record successful query pattern
      if (success && results.items.length > 0) {
        this.knowledgeBase.recordSuccessfulQuery(
          query,
          results.items.length,
          results.sources,
          context
        );

        // Record query pattern if entity types are known
        if (results.entityTypes && results.entityTypes.length > 0) {
          this.knowledgeBase.recordQueryPattern(query, results.entityTypes, true, context);
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

      // 4. Update entity mappings
      if (success && results.entityTypes) {
        for (const et of results.entityTypes) {
          this.knowledgeBase.recordEntityMapping(query, [et], true);
        }
      }

      // 5. Persist knowledge (async, don't wait)
      this.knowledgeBase.save().catch((error) => {
        logger.warn(`Failed to persist knowledge base: ${error}`);
      });

      logger.debug(
        `Learning system: recorded ${success ? 'successful' : 'failed'} search for "${query}"`
      );
    } catch (error) {
      logger.warn(`Learning system error: ${error}`);
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

    // Default entity types based on context
    if (context === 'project') {
      return ['driveItem', 'site', 'chatMessage', 'event', 'message'];
    }
    if (context === 'person') {
      return ['person', 'message', 'event', 'driveItem'];
    }
    if (context === 'document') {
      return ['driveItem', 'message'];
    }
    if (context === 'meeting') {
      return ['event', 'message', 'driveItem'];
    }

    // Default: search all
    return ['driveItem', 'message', 'event', 'site', 'chatMessage', 'person'];
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
}

export default LearningSystem;
