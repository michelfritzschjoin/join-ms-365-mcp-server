/**
 * UQAS Pro - Universal Question Answering System
 *
 * Bilingual (DE/EN) question answering for Microsoft 365 with:
 * - Automatic language detection
 * - Cross-language search optimization
 * - Multi-layer adaptive analysis
 * - Token-optimized responses
 * - Smart caching
 *
 * @module uqas
 */

// Re-export all i18n components
export * from './i18n/index.js';

// Re-export search components
export * from './search/index.js';

// Re-export core components
export * from './core/index.js';

// Re-export integration components
export * from './integration/index.js';

// Main UQAS orchestrator
import { LanguageDetector } from './i18n/language-detector.js';
import { BilingualThesaurus } from './i18n/bilingual-thesaurus.js';
import { BilingualIntentRecognizer } from './i18n/intent-patterns.js';
import { BilingualEntityRecognizer } from './i18n/entity-patterns.js';
import { BilingualResponseBuilder, type ResponseData } from './i18n/response-templates.js';
import { CrossLanguageSearchExpander } from './search/cross-language-expander.js';
import { AdaptiveLayerController, type AnalysisDepth } from './core/adaptive-layer.js';
import { EntityGraphBuilder } from './core/entity-graph.js';
import { TokenController } from './core/token-controller.js';
import { CacheManager } from './core/cache-manager.js';
import type { SupportedLanguage } from './i18n/index.js';

/**
 * Configuration for UQAS Pro
 */
export interface UQASConfig {
  /** Maximum analysis depth (1-5) */
  maxDepth?: number;
  /** Token budget per response */
  tokenBudget?: number;
  /** Enable caching */
  enableCache?: boolean;
  /** Cache TTL in seconds */
  cacheTTL?: number;
  /** Preferred output language (auto-detected if not set) */
  preferredLanguage?: SupportedLanguage;
  /** Confidence threshold to stop iteration */
  confidenceThreshold?: number;
}

/**
 * Result from UQAS query
 */
export interface UQASResult {
  /** The answer in user's language */
  answer: string;
  /** Detected input language */
  language: SupportedLanguage;
  /** Confidence score (0-1) */
  confidence: number;
  /** Analysis depth reached (1-5) */
  depth: number;
  /** Number of sources consulted */
  sourceCount: number;
  /** Token count of response */
  tokenCount: number;
  /** Whether result came from cache */
  cached: boolean;
  /** Structured response data */
  data: ResponseData;
}

/**
 * UQAS Pro - Main Orchestrator
 *
 * Coordinates all components to answer user questions using M365 data.
 */
export class UQASPro {
  private languageDetector: LanguageDetector;
  private thesaurus: BilingualThesaurus;
  private intentRecognizer: BilingualIntentRecognizer;
  private entityRecognizer: BilingualEntityRecognizer;
  private searchExpander: CrossLanguageSearchExpander;
  private layerController: AdaptiveLayerController;
  private entityGraph: EntityGraphBuilder;
  private tokenController: TokenController;
  private cacheManager: CacheManager;
  private config: Required<UQASConfig>;

  constructor(config: UQASConfig = {}) {
    this.config = {
      maxDepth: config.maxDepth ?? 3,
      tokenBudget: config.tokenBudget ?? 1500,
      enableCache: config.enableCache ?? true,
      cacheTTL: config.cacheTTL ?? 300,
      preferredLanguage: config.preferredLanguage ?? ('de' as SupportedLanguage),
      confidenceThreshold: config.confidenceThreshold ?? 0.8,
    };

    // Initialize components
    this.languageDetector = new LanguageDetector();
    this.thesaurus = new BilingualThesaurus();
    this.intentRecognizer = new BilingualIntentRecognizer();
    this.entityRecognizer = new BilingualEntityRecognizer();
    this.searchExpander = new CrossLanguageSearchExpander(this.thesaurus);
    this.layerController = new AdaptiveLayerController({
      maxDepth: this.config.maxDepth as AnalysisDepth,
      confidenceThreshold: this.config.confidenceThreshold,
    });
    this.entityGraph = new EntityGraphBuilder();
    this.tokenController = new TokenController({
      maxTokens: this.config.tokenBudget,
    });
    this.cacheManager = new CacheManager({
      enabled: this.config.enableCache,
      defaultTTL: this.config.cacheTTL,
    });
  }

  /**
   * Answer a user question using M365 data
   */
  async answer(question: string, context?: Record<string, unknown>): Promise<UQASResult> {
    // Check cache first
    const cacheKey = this.cacheManager.generateKey(question, context);
    const cached = this.cacheManager.get(cacheKey);
    if (cached) {
      return { ...cached, cached: true } as UQASResult;
    }

    // Detect language
    const langResult = this.languageDetector.detect(question);
    const lang = this.config.preferredLanguage || langResult.lang;

    // Recognize intent
    const intent = this.intentRecognizer.recognize(question);

    // Extract entities
    const entities = this.entityRecognizer.extractAll(question, lang);

    // Expand search queries
    const searchQueries = this.searchExpander.expand(question, lang);

    // Execute adaptive layer analysis
    const layerResult = await this.layerController.analyze({
      question,
      language: lang,
      intent,
      entities,
      searchQueries,
      context,
    });

    // Build entity graph from results
    const graph = this.entityGraph.build(layerResult.results);

    // Build response with token control
    const responseBuilder = new BilingualResponseBuilder(lang);
    const responseData: ResponseData = {
      summary: layerResult.summary,
      confidence: layerResult.confidence,
      sourceCount: layerResult.sourceCount,
      depth: layerResult.depth,
      facts: layerResult.keyFacts,
      timeline: layerResult.timeline,
      recommendations: layerResult.recommendations,
      entities: graph.nodes,
    };

    const compactResponse = this.tokenController.optimize(responseData);
    const answer = responseBuilder.buildResponse(compactResponse);

    const result: UQASResult = {
      answer,
      language: lang,
      confidence: layerResult.confidence,
      depth: layerResult.depth,
      sourceCount: layerResult.sourceCount,
      tokenCount: this.tokenController.estimateTokens(answer),
      cached: false,
      data: compactResponse,
    };

    // Cache the result
    this.cacheManager.set(cacheKey, result);

    return result;
  }

  /**
   * Get language detector for external use
   */
  getLanguageDetector(): LanguageDetector {
    return this.languageDetector;
  }

  /**
   * Get thesaurus for external use
   */
  getThesaurus(): BilingualThesaurus {
    return this.thesaurus;
  }

  /**
   * Clear the cache
   */
  clearCache(): void {
    this.cacheManager.clear();
  }
}

export default UQASPro;
