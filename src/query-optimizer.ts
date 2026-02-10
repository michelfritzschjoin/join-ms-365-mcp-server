/**
 * Query Optimizer - Automatic query optimization for better search results
 *
 * Combines NLP optimization with history-based pattern learning to automatically
 * improve queries before execution. Applied generically across all query handlers
 * (Email, Calendar, Files, Search, Assistant).
 *
 * Features:
 * - Learned pattern transformations (e.g., "Projekt P2046" → "P2046")
 * - Synonym expansion from business terminology
 * - Stopword removal for cleaner queries
 * - Context-aware optimization (different strategies per tool)
 * - Project-specific pattern detection and optimization
 * - Confidence scoring for all optimizations
 */

import logger from './logger.js';
import NLPEnhancer from './nlp-enhancer.js';
import { getQueryStore } from './query-store.js';

// ============================================================================
// TYPES
// ============================================================================

/** Type of optimization applied to a query */
export type OptimizationType =
  | 'normalization'
  | 'synonym_expansion'
  | 'pattern_transformation'
  | 'stopword_removal'
  | 'project_extraction'
  | 'identifier_extraction'
  | 'abbreviation_expansion'
  | 'cross_language';

/** Single optimization step applied to a query */
export interface OptimizationStep {
  /** Type of optimization applied */
  type: OptimizationType;
  /** Human-readable description of what was done */
  description: string;
  /** Original query before optimization */
  original: string;
  /** Optimized query after optimization */
  optimized: string;
  /** Confidence score for this optimization (0.0 - 1.0) */
  confidence: number;
}

/** Result of query optimization */
export interface OptimizedQuery {
  /** Original query as provided */
  originalQuery: string;
  /** Primary optimized query to use */
  optimizedQuery: string;
  /** Alternative query variants for multi-query strategies */
  variants: string[];
  /** Overall confidence score (0.0 - 1.0) */
  confidence: number;
  /** List of optimization steps applied */
  optimizations: OptimizationStep[];
  /** Whether optimization used historical patterns */
  learnedFromHistory: boolean;
  /** NLP analysis data */
  nlpAnalysis: {
    intent?: string;
    service?: string;
    entities?: Array<{ value: string; type: string; confidence?: number }>;
    temporal?: { expression: string; type: string; relativeDays?: number } | null;
    confidence?: number;
  };
  /** Extracted temporal filters */
  filters?: Record<string, unknown>;
}

/** Context for tool-specific optimization */
export interface OptimizationContext {
  /** Tool name (e.g., 'email', 'files', 'search', 'calendar', 'assistant') */
  tool: string;
  /** Entity types being searched */
  entityTypes?: string[];
  /** User ID hash for personalized learning */
  userIdHash?: string;
}

/** Stored query transformation for learning */
export interface QueryTransformation {
  /** Original query pattern */
  originalPattern: string;
  /** Optimized query that worked */
  optimizedPattern: string;
  /** Tool context where this worked */
  toolContext: string;
  /** Success count */
  successCount: number;
  /** Failure count */
  failureCount: number;
  /** Last used timestamp */
  lastUsed: string;
}

// ============================================================================
// CONFIGURATION
// ============================================================================

/** Whether automatic query optimization is enabled */
const AUTO_OPTIMIZATION_ENABLED = process.env.MS365_MCP_AUTO_QUERY_OPTIMIZATION_ENABLED !== 'false';

/** Minimum confidence threshold to apply an optimization */
const CONFIDENCE_THRESHOLD = parseFloat(
  process.env.MS365_MCP_QUERY_OPTIMIZATION_CONFIDENCE_THRESHOLD || '0.6'
);

/** Minimum pattern occurrences before learning is applied */
const MIN_PATTERN_COUNT = parseInt(
  process.env.MS365_MCP_QUERY_OPTIMIZATION_MIN_PATTERN_COUNT || '2',
  10
);

// ============================================================================
// PATTERNS AND DICTIONARIES
// ============================================================================

/** German/English stopwords for removal */
const STOPWORDS = new Set([
  // German
  'der',
  'die',
  'das',
  'den',
  'dem',
  'des',
  'ein',
  'eine',
  'einer',
  'einem',
  'einen',
  'eines',
  'und',
  'oder',
  'aber',
  'in',
  'im',
  'auf',
  'an',
  'am',
  'zu',
  'zum',
  'zur',
  'für',
  'von',
  'vom',
  'mit',
  'durch',
  'über',
  'unter',
  'vor',
  'nach',
  'bei',
  'aus',
  'es',
  'noch',
  'schon',
  'auch',
  'nur',
  'doch',
  'mal',
  'ist',
  'sind',
  'war',
  'waren',
  'wird',
  'werden',
  'wurde',
  'wurden',
  'hat',
  'haben',
  'hatte',
  'hatten',
  'sein',
  'gewesen',
  'gibt',
  'gab',
  'geben',
  // German question/action words to strip for search optimization
  'suche',
  'suchen',
  'finde',
  'finden',
  'zeige',
  'zeigen',
  'wo',
  'wie',
  'was',
  'wer',
  'wann',
  'welche',
  'welcher',
  'welches',
  'welchen',
  'welchem',
  'bitte',
  'mir',
  'mich',
  'meine',
  'meinen',
  'meinem',
  'meiner',
  'alle',
  // English
  'the',
  'a',
  'an',
  'and',
  'or',
  'but',
  'in',
  'on',
  'at',
  'to',
  'for',
  'of',
  'with',
  'by',
  'from',
  'up',
  'about',
  'into',
  'over',
  'after',
  'is',
  'are',
  'was',
  'were',
  'be',
  'been',
  'being',
  'have',
  'has',
  'had',
  'do',
  'does',
  'did',
  'will',
  'would',
  'should',
  'could',
  'can',
  'may',
  'might',
  'must',
  'shall',
  'it',
  'this',
  'that',
  'these',
  'those',
  'there',
  // English action words to strip
  'find',
  'search',
  'show',
  'get',
  'list',
  'look',
  'where',
  'what',
  'who',
  'when',
  'which',
  'how',
  'please',
  'me',
  'my',
  'mine',
  'all',
]);

/** Project-related keywords (DE/EN) */
const PROJECT_KEYWORDS = new Set([
  'projekt',
  'project',
  'vorhaben',
  'initiative',
  'programm',
  'program',
]);

/** Identifier patterns for project codes, ticket numbers, etc. */
const IDENTIFIER_PATTERNS: Array<{
  pattern: RegExp;
  type: string;
  description: string;
}> = [
  {
    pattern: /\b([A-Z]{1,5}[-]?\d{2,6})\b/g,
    type: 'project_code',
    description: 'Project/ticket code (e.g., P2046, PROJ-123, AB1234)',
  },
  {
    pattern: /\b(PRJ|PROJ|TKT|INC|REQ|CR)[-_]?\d{3,8}\b/gi,
    type: 'ticket_number',
    description: 'Ticket/incident number',
  },
  {
    pattern: /\b\d{5,10}\b/g,
    type: 'numeric_id',
    description: 'Numeric identifier',
  },
];

/** Cross-language term mapping for bilingual search (DE ↔ EN) */
const CROSS_LANGUAGE_TERMS: Record<string, string> = {
  // DE → EN
  projekt: 'project',
  besprechung: 'meeting',
  termin: 'appointment',
  aufgabe: 'task',
  dokument: 'document',
  datei: 'file',
  nachricht: 'message',
  vertrag: 'contract',
  angebot: 'offer',
  rechnung: 'invoice',
  kunde: 'customer',
  bericht: 'report',
  praesentation: 'presentation',
  präsentation: 'presentation',
  kalender: 'calendar',
  notiz: 'note',
  ordner: 'folder',
  entwurf: 'draft',
  genehmigung: 'approval',
  budget: 'budget',
  // EN → DE
  project: 'projekt',
  meeting: 'besprechung',
  appointment: 'termin',
  task: 'aufgabe',
  document: 'dokument',
  file: 'datei',
  message: 'nachricht',
  contract: 'vertrag',
  offer: 'angebot',
  invoice: 'rechnung',
  customer: 'kunde',
  report: 'bericht',
  presentation: 'präsentation',
  calendar: 'kalender',
  note: 'notiz',
  folder: 'ordner',
  draft: 'entwurf',
  approval: 'genehmigung',
};

// ============================================================================
// QUERY OPTIMIZER CLASS
// ============================================================================

/**
 * QueryOptimizer - Automatically optimizes search queries for better results
 *
 * Applied generically across all query handlers.
 */
export class QueryOptimizer {
  private nlpEnhancer: NLPEnhancer;
  private readonly enabled: boolean;

  constructor() {
    this.nlpEnhancer = new NLPEnhancer();
    this.enabled = AUTO_OPTIMIZATION_ENABLED;
  }

  /**
   * Main optimization method - optimizes a query for search
   * @param query - Original query string
   * @param context - Tool context for optimization
   * @returns Optimized query with metadata
   */
  optimizeQuery(query: string, context: OptimizationContext = { tool: 'search' }): OptimizedQuery {
    if (!this.enabled || !query || query.trim().length === 0) {
      return this.createPassthroughResult(query);
    }

    const optimizations: OptimizationStep[] = [];
    const variants: string[] = [];
    let currentQuery = query.trim();
    let learnedFromHistory = false;
    let overallConfidence = 0.5;

    // =====================================================================
    // STEP 1: NLP Analysis (always performed for metadata)
    // =====================================================================
    const decomposed = this.nlpEnhancer.decomposeQuery(query);
    const nlpAnalysis = {
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

    // Extract temporal filters
    const filters: Record<string, unknown> = {};
    if (decomposed.temporal?.relativeDays !== undefined) {
      const now = new Date();
      const date = new Date(now);
      date.setDate(date.getDate() + decomposed.temporal.relativeDays);
      filters.dateFilter = date.toISOString();
    }

    // =====================================================================
    // STEP 2: History-based pattern transformation
    // =====================================================================
    const historyResult = this.applyLearnedTransformations(currentQuery, context);
    if (historyResult) {
      optimizations.push(historyResult.step);
      currentQuery = historyResult.optimized;
      learnedFromHistory = true;
      overallConfidence = Math.max(overallConfidence, historyResult.step.confidence);
      if (historyResult.variants.length > 0) {
        variants.push(...historyResult.variants);
      }
    }

    // =====================================================================
    // STEP 3: Project/Identifier extraction
    // =====================================================================
    const identifierResult = this.extractIdentifiers(currentQuery);
    if (identifierResult) {
      optimizations.push(identifierResult.step);
      // Add identifier-focused variant
      variants.push(identifierResult.identifier);
      // If query is primarily a project query, use identifier as main query
      if (
        identifierResult.isProjectQuery &&
        identifierResult.step.confidence >= CONFIDENCE_THRESHOLD
      ) {
        currentQuery = identifierResult.optimized;
        overallConfidence = Math.max(overallConfidence, identifierResult.step.confidence);
      }
    }

    // =====================================================================
    // STEP 4: Stopword removal (for cleaner search queries)
    // =====================================================================
    const cleanedQuery = this.removeStopwords(currentQuery);
    if (cleanedQuery && cleanedQuery !== currentQuery && cleanedQuery.length >= 2) {
      optimizations.push({
        type: 'stopword_removal',
        description: `Removed stopwords for cleaner search`,
        original: currentQuery,
        optimized: cleanedQuery,
        confidence: 0.7,
      });
      // Keep original as variant, use cleaned as primary
      if (currentQuery !== query) {
        variants.push(currentQuery);
      }
      currentQuery = cleanedQuery;
    }

    // =====================================================================
    // STEP 5: Cross-language expansion
    // =====================================================================
    const crossLangVariants = this.generateCrossLanguageVariants(currentQuery);
    if (crossLangVariants.length > 0) {
      optimizations.push({
        type: 'cross_language',
        description: `Added cross-language variants: ${crossLangVariants.join(', ')}`,
        original: currentQuery,
        optimized: currentQuery,
        confidence: 0.65,
      });
      variants.push(...crossLangVariants);
    }

    // =====================================================================
    // STEP 6: Synonym expansion
    // =====================================================================
    const synonymVariants = this.expandWithSynonyms(currentQuery, context);
    if (synonymVariants.length > 0) {
      optimizations.push({
        type: 'synonym_expansion',
        description: `Generated ${synonymVariants.length} synonym variant(s)`,
        original: currentQuery,
        optimized: currentQuery,
        confidence: 0.6,
      });
      variants.push(...synonymVariants);
    }

    // =====================================================================
    // STEP 7: Normalization (trim, collapse whitespace)
    // =====================================================================
    const normalized = currentQuery.replace(/\s+/g, ' ').trim();
    if (normalized !== currentQuery) {
      optimizations.push({
        type: 'normalization',
        description: 'Normalized whitespace',
        original: currentQuery,
        optimized: normalized,
        confidence: 1.0,
      });
      currentQuery = normalized;
    }

    // Deduplicate variants and remove current query from variants
    const uniqueVariants = [...new Set(variants)]
      .filter((v) => v !== currentQuery && v !== query && v.trim().length > 0)
      .slice(0, 8);

    // Calculate final confidence
    if (optimizations.length > 0) {
      const avgConfidence =
        optimizations.reduce((sum, o) => sum + o.confidence, 0) / optimizations.length;
      overallConfidence = Math.max(overallConfidence, avgConfidence);
    }

    return {
      originalQuery: query,
      optimizedQuery: currentQuery,
      variants: uniqueVariants,
      confidence: Math.min(1.0, overallConfidence),
      optimizations,
      learnedFromHistory,
      nlpAnalysis,
      ...(Object.keys(filters).length > 0 && { filters }),
    };
  }

  // ============================================================================
  // PRIVATE OPTIMIZATION METHODS
  // ============================================================================

  /**
   * Apply learned transformations from query history
   */
  private applyLearnedTransformations(
    query: string,
    context: OptimizationContext
  ): { optimized: string; step: OptimizationStep; variants: string[] } | null {
    if (!context.userIdHash) {
      return null;
    }

    try {
      const queryStore = getQueryStore();
      const transformations = queryStore.getQueryTransformationPatterns(
        context.userIdHash,
        context.tool
      );

      if (transformations.length === 0) {
        return null;
      }

      const queryLower = query.toLowerCase().trim();
      const variants: string[] = [];

      // Find matching transformation
      for (const transform of transformations) {
        const originalLower = transform.originalPattern.toLowerCase();
        if (queryLower === originalLower || queryLower.includes(originalLower)) {
          const successRate =
            transform.successCount + transform.failureCount > 0
              ? transform.successCount / (transform.successCount + transform.failureCount)
              : 0;

          if (successRate >= 0.5 && transform.successCount >= MIN_PATTERN_COUNT) {
            const optimized = query.replace(
              new RegExp(this.escapeRegExp(transform.originalPattern), 'gi'),
              transform.optimizedPattern
            );

            // Also add the optimized pattern as variant if different
            if (transform.optimizedPattern !== query) {
              variants.push(transform.optimizedPattern);
            }

            return {
              optimized,
              step: {
                type: 'pattern_transformation',
                description: `Applied learned pattern: "${transform.originalPattern}" → "${transform.optimizedPattern}" (${Math.round(successRate * 100)}% success, ${transform.successCount} uses)`,
                original: query,
                optimized,
                confidence: Math.min(0.95, successRate),
              },
              variants,
            };
          }
        }
      }
    } catch (error) {
      logger.debug('Failed to apply learned transformations', { error });
    }

    return null;
  }

  /**
   * Extract identifiers (project codes, ticket numbers) and optimize query
   */
  private extractIdentifiers(query: string): {
    identifier: string;
    optimized: string;
    isProjectQuery: boolean;
    step: OptimizationStep;
  } | null {
    const queryLower = query.toLowerCase();
    const words = queryLower.split(/\s+/);

    // Check if this is a project-type query
    const hasProjectKeyword = words.some((w) => PROJECT_KEYWORDS.has(w));

    // Try to extract identifiers
    for (const { pattern, type, description } of IDENTIFIER_PATTERNS) {
      pattern.lastIndex = 0;
      const match = pattern.exec(query);
      if (match) {
        const identifier = match[0];

        // If query is "Projekt P2046" or similar, optimize to just the identifier
        // but keep original context words as variant
        if (hasProjectKeyword) {
          // Remove project keyword and keep the rest
          const withoutKeyword = words
            .filter((w) => !PROJECT_KEYWORDS.has(w))
            .join(' ')
            .trim();

          return {
            identifier,
            optimized: withoutKeyword || identifier,
            isProjectQuery: true,
            step: {
              type: 'project_extraction',
              description: `Extracted ${type}: "${identifier}" from project query`,
              original: query,
              optimized: withoutKeyword || identifier,
              confidence: 0.85,
            },
          };
        }

        // For non-project queries with identifiers, extract the identifier as variant
        return {
          identifier,
          optimized: query,
          isProjectQuery: false,
          step: {
            type: 'identifier_extraction',
            description: `Found ${type}: "${identifier}" (${description})`,
            original: query,
            optimized: query,
            confidence: 0.7,
          },
        };
      }
    }

    return null;
  }

  /**
   * Remove stopwords from query for cleaner search
   */
  private removeStopwords(query: string): string | null {
    const words = query.split(/\s+/);
    if (words.length <= 1) {
      return null; // Don't remove stopwords from single-word queries
    }

    const significantWords = words.filter((w) => {
      const lower = w.toLowerCase();
      // Keep the word if it's not a stopword, or if it starts with uppercase (proper noun),
      // or if it matches an identifier pattern, or if it contains digits
      return !STOPWORDS.has(lower) || /^[A-ZÄÖÜ]/.test(w) || /\d/.test(w) || w.length <= 1;
    });

    if (significantWords.length === 0) {
      return null; // Don't return empty query
    }

    const result = significantWords.join(' ').trim();
    return result.length > 0 ? result : null;
  }

  /**
   * Generate cross-language variants (DE ↔ EN)
   */
  private generateCrossLanguageVariants(query: string): string[] {
    const variants: string[] = [];
    const words = query.toLowerCase().split(/\s+/);

    for (const word of words) {
      const translation = CROSS_LANGUAGE_TERMS[word];
      if (translation && !query.toLowerCase().includes(translation)) {
        // Create variant with translated word
        const variant = query.replace(
          new RegExp(`\\b${this.escapeRegExp(word)}\\b`, 'gi'),
          translation
        );
        if (variant !== query) {
          variants.push(variant);
        }
      }
    }

    return variants.slice(0, 3); // Limit to 3 cross-language variants
  }

  /**
   * Expand query with synonyms
   */
  private expandWithSynonyms(query: string, context: OptimizationContext): string[] {
    const variants: string[] = [];
    const decomposed = this.nlpEnhancer.decomposeQuery(query);

    // Use semantic variants from NLP decomposer
    if (decomposed.semanticVariants.length > 0) {
      variants.push(...decomposed.semanticVariants.slice(0, 3));
    }

    // Use compound parts as additional variants
    if (decomposed.compoundParts.length > 1) {
      variants.push(decomposed.compoundParts.join(' '));
    }

    return variants.slice(0, 5);
  }

  /**
   * Create passthrough result when optimization is disabled or not applicable
   */
  private createPassthroughResult(query: string): OptimizedQuery {
    const decomposed = this.nlpEnhancer.decomposeQuery(query || '');
    return {
      originalQuery: query,
      optimizedQuery: query,
      variants: [],
      confidence: 0.5,
      optimizations: [],
      learnedFromHistory: false,
      nlpAnalysis: {
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
      },
    };
  }

  /**
   * Escape special regex characters
   */
  private escapeRegExp(str: string): string {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
}

// ============================================================================
// SINGLETON
// ============================================================================

let queryOptimizerInstance: QueryOptimizer | null = null;

/**
 * Get the QueryOptimizer singleton instance
 */
export function getQueryOptimizer(): QueryOptimizer {
  if (!queryOptimizerInstance) {
    queryOptimizerInstance = new QueryOptimizer();
  }
  return queryOptimizerInstance;
}

export default QueryOptimizer;
