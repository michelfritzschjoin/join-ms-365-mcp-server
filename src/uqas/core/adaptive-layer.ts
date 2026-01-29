/**
 * Adaptive Layer Controller
 *
 * Multi-layer analysis system (L1-L5) that adapts depth based on
 * confidence and query complexity. Stops early when sufficient
 * information is gathered.
 */

import type { SupportedLanguage } from '../i18n/index.js';
import type { BilingualIntent } from '../i18n/intent-patterns.js';
import type { ExtractedEntity } from '../i18n/entity-patterns.js';
import type { SearchQuerySet } from '../search/cross-language-expander.js';
import type { TimelineEntry } from '../i18n/response-templates.js';

/**
 * Analysis depth levels
 */
export type AnalysisDepth = 1 | 2 | 3 | 4 | 5;

/**
 * Layer configuration
 */
export interface LayerConfig {
  maxDepth?: AnalysisDepth;
  confidenceThreshold?: number;
  timeout?: number;
  tokenBudgetPerLayer?: number[];
}

/**
 * Input for layer analysis
 */
export interface LayerInput {
  question: string;
  language: SupportedLanguage;
  intent: BilingualIntent;
  entities: ExtractedEntity[];
  searchQueries: SearchQuerySet;
  context?: Record<string, unknown>;
}

/**
 * Result from layer analysis
 */
export interface LayerResult {
  /** Analysis depth reached */
  depth: AnalysisDepth;
  /** Overall confidence (0-1) */
  confidence: number;
  /** Number of sources consulted */
  sourceCount: number;
  /** Summary text */
  summary: string;
  /** Key facts extracted */
  keyFacts: string[];
  /** Timeline of events */
  timeline?: TimelineEntry[];
  /** Recommended actions */
  recommendations?: string[];
  /** Raw results by layer */
  results: LayerResultSet[];
  /** Execution time in ms */
  executionTime: number;
  /** Whether confidence threshold was met */
  thresholdMet: boolean;
}

/**
 * Results from a single layer
 */
export interface LayerResultSet {
  layer: AnalysisDepth;
  sources: string[];
  items: unknown[];
  tokenCount: number;
  executionTime: number;
}

/**
 * Layer definitions
 */
const LAYER_DEFINITIONS: Record<
  AnalysisDepth,
  {
    name: string;
    description: string;
    defaultTokenBudget: number;
    sources: string[];
  }
> = {
  1: {
    name: 'Quick Scan',
    description: 'Unified Search, top 10 results',
    defaultTokenBudget: 500,
    sources: ['search'],
  },
  2: {
    name: 'Focused Search',
    description: 'Specific APIs based on intent',
    defaultTokenBudget: 1500,
    sources: ['email', 'calendar', 'files'],
  },
  3: {
    name: 'Deep Dive',
    description: 'All relevant sources',
    defaultTokenBudget: 3000,
    sources: ['email', 'calendar', 'files', 'teams', 'tasks', 'people'],
  },
  4: {
    name: 'Cross-Reference',
    description: 'Entity relationships',
    defaultTokenBudget: 2000,
    sources: ['entity-graph', 'relationships'],
  },
  5: {
    name: 'Historical Context',
    description: 'Temporal patterns and trends',
    defaultTokenBudget: 1500,
    sources: ['history', 'patterns'],
  },
};

/**
 * AdaptiveLayerController - Manages multi-layer analysis
 */
export class AdaptiveLayerController {
  private config: Required<LayerConfig>;
  private currentLayer: AnalysisDepth = 1;

  constructor(config: LayerConfig = {}) {
    this.config = {
      maxDepth: config.maxDepth ?? 3,
      confidenceThreshold: config.confidenceThreshold ?? 0.8,
      timeout: config.timeout ?? 30000,
      tokenBudgetPerLayer: config.tokenBudgetPerLayer ?? [500, 1500, 3000, 2000, 1500],
    };
  }

  /**
   * Analyze question through adaptive layers
   */
  async analyze(input: LayerInput): Promise<LayerResult> {
    const startTime = Date.now();
    const results: LayerResultSet[] = [];
    let totalConfidence = 0;
    let totalSources = 0;
    const allItems: unknown[] = [];
    const allFacts: string[] = [];

    // Process layers until confidence threshold or max depth
    for (let layer = 1 as AnalysisDepth; layer <= this.config.maxDepth; layer++) {
      this.currentLayer = layer;
      // Track layer execution time

      // Execute layer
      const layerResult = await this.executeLayer(layer, input, allItems);
      results.push(layerResult);

      // Accumulate results
      allItems.push(...layerResult.items);
      totalSources += layerResult.sources.length;

      // Calculate running confidence
      totalConfidence = this.calculateConfidence(allItems, input);

      // Extract facts from this layer
      const layerFacts = this.extractFacts(layerResult.items, layer, input.language);
      allFacts.push(...layerFacts);

      // Check if we should stop
      if (totalConfidence >= this.config.confidenceThreshold) {
        break;
      }

      // Check timeout
      if (Date.now() - startTime > this.config.timeout) {
        break;
      }
    }

    // Build timeline from results
    const timeline = this.buildTimeline(allItems, input.language);

    // Generate recommendations
    const recommendations = this.generateRecommendations(input, allItems);

    // Generate summary
    const summary = this.generateSummary(input, allFacts, allItems.length, input.language);

    return {
      depth: this.currentLayer,
      confidence: Math.min(1, totalConfidence),
      sourceCount: totalSources,
      summary,
      keyFacts: allFacts.slice(0, 5), // Top 5 facts
      timeline: timeline.slice(0, 10), // Top 10 timeline entries
      recommendations,
      results,
      executionTime: Date.now() - startTime,
      thresholdMet: totalConfidence >= this.config.confidenceThreshold,
    };
  }

  /**
   * Execute a single layer
   */
  private async executeLayer(
    layer: AnalysisDepth,
    input: LayerInput,
    _existingItems: unknown[]
  ): Promise<LayerResultSet> {
    const startTime = Date.now();

    // Simulate layer execution (in real implementation, this would call Graph API)
    // For now, return mock results that can be replaced with actual API calls
    const items = await this.executeLayerQueries(layer, input);
    const sources = this.determineSourcesForLayer(layer, input);

    return {
      layer,
      sources,
      items,
      tokenCount: this.estimateTokenCount(items),
      executionTime: Date.now() - startTime,
    };
  }

  /**
   * Execute queries for a specific layer
   * This is a placeholder - in real implementation, integrate with GraphClient
   */
  private async executeLayerQueries(_layer: AnalysisDepth, _input: LayerInput): Promise<unknown[]> {
    // This method should be overridden or extended to actually query M365
    // For now, return empty array - integration will add actual queries
    return [];
  }

  /**
   * Determine which sources to query for a layer
   */
  private determineSourcesForLayer(layer: AnalysisDepth, input: LayerInput): string[] {
    const layerDef = LAYER_DEFINITIONS[layer];
    const intentSources = input.intent.suggestedSources;

    // Layer 1-2: Use intent-suggested sources
    if (layer <= 2) {
      return intentSources.slice(0, layer === 1 ? 1 : 3);
    }

    // Layer 3+: Use all layer sources
    return layerDef.sources;
  }

  /**
   * Calculate confidence based on results
   */
  private calculateConfidence(items: unknown[], input: LayerInput): number {
    if (items.length === 0) return 0;

    let confidence = 0;

    // Base confidence from item count
    if (items.length >= 10) {
      confidence += 0.3;
    } else if (items.length >= 5) {
      confidence += 0.2;
    } else if (items.length >= 1) {
      confidence += 0.1;
    }

    // Bonus for matching entity mentions
    const entityMatches = this.countEntityMatches(items, input.entities);
    confidence += Math.min(0.3, entityMatches * 0.1);

    // Bonus for source variety
    const sources = this.extractSources(items);
    confidence += Math.min(0.2, sources.length * 0.05);

    // Bonus for recency
    const hasRecentItems = this.hasRecentItems(items);
    if (hasRecentItems) {
      confidence += 0.1;
    }

    // Bonus based on intent match
    confidence += input.intent.confidence * 0.1;

    return Math.min(1, confidence);
  }

  /**
   * Count how many entities are found in results
   */
  private countEntityMatches(items: unknown[], entities: ExtractedEntity[]): number {
    let matches = 0;
    const itemStr = JSON.stringify(items).toLowerCase();

    for (const entity of entities) {
      if (itemStr.includes(entity.value.toLowerCase())) {
        matches++;
      }
    }

    return matches;
  }

  /**
   * Extract unique sources from items
   */
  private extractSources(items: unknown[]): string[] {
    const sources = new Set<string>();

    for (const item of items) {
      if (typeof item === 'object' && item !== null) {
        const obj = item as Record<string, unknown>;
        if (obj['@odata.type']) {
          sources.add(String(obj['@odata.type']));
        }
      }
    }

    return Array.from(sources);
  }

  /**
   * Check if items contain recent entries
   */
  private hasRecentItems(items: unknown[]): boolean {
    const now = Date.now();
    const oneWeekAgo = now - 7 * 24 * 60 * 60 * 1000;

    for (const item of items) {
      if (typeof item === 'object' && item !== null) {
        const obj = item as Record<string, unknown>;
        const dateFields = ['createdDateTime', 'lastModifiedDateTime', 'start', 'receivedDateTime'];

        for (const field of dateFields) {
          if (obj[field]) {
            const date = new Date(String(obj[field]));
            if (date.getTime() > oneWeekAgo) {
              return true;
            }
          }
        }
      }
    }

    return false;
  }

  /**
   * Extract facts from layer results
   */
  private extractFacts(items: unknown[], layer: AnalysisDepth, lang: SupportedLanguage): string[] {
    const facts: string[] = [];

    for (const item of items.slice(0, 5)) {
      if (typeof item === 'object' && item !== null) {
        const obj = item as Record<string, unknown>;
        const fact = this.itemToFact(obj, lang);
        if (fact) {
          facts.push(fact);
        }
      }
    }

    return facts;
  }

  /**
   * Convert item to fact string
   */
  private itemToFact(item: Record<string, unknown>, lang: SupportedLanguage): string | null {
    // Email
    if (item.subject && item.from) {
      const from = this.extractName(item.from);
      return `📧 ${item.subject} (${lang === 'de' ? 'von' : 'from'} ${from})`;
    }

    // Event
    if (item.subject && item.start) {
      const date = this.formatDateShort(String(item.start), lang);
      return `📅 ${item.subject} (${date})`;
    }

    // File
    if (item.name && item.webUrl) {
      return `📁 ${item.name}`;
    }

    // Task
    if (item.title && item.status) {
      return `✅ ${item.title}`;
    }

    return null;
  }

  /**
   * Extract name from complex object
   */
  private extractName(obj: unknown): string {
    if (typeof obj === 'string') return obj;
    if (typeof obj === 'object' && obj !== null) {
      const o = obj as Record<string, unknown>;
      if (o.emailAddress && typeof o.emailAddress === 'object') {
        return ((o.emailAddress as Record<string, unknown>).name as string) || 'Unknown';
      }
      return (o.name as string) || (o.displayName as string) || 'Unknown';
    }
    return 'Unknown';
  }

  /**
   * Format date in short form
   */
  private formatDateShort(dateStr: string, lang: SupportedLanguage): string {
    const date = new Date(dateStr);
    if (lang === 'de') {
      return `${date.getDate()}.${date.getMonth() + 1}.`;
    } else {
      return `${date.getMonth() + 1}/${date.getDate()}`;
    }
  }

  /**
   * Build timeline from results
   */
  private buildTimeline(items: unknown[], lang: SupportedLanguage): TimelineEntry[] {
    const timeline: TimelineEntry[] = [];

    for (const item of items) {
      if (typeof item === 'object' && item !== null) {
        const obj = item as Record<string, unknown>;
        const entry = this.itemToTimelineEntry(obj);
        if (entry) {
          timeline.push(entry);
        }
      }
    }

    // Sort by date (most recent first)
    timeline.sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime());

    return timeline;
  }

  /**
   * Convert item to timeline entry
   */
  private itemToTimelineEntry(item: Record<string, unknown>): TimelineEntry | null {
    // Event
    if (item.start) {
      const startObj = item.start as Record<string, unknown>;
      return {
        date: String(startObj.dateTime || item.start),
        title: String(item.subject || 'Event'),
        type: 'meeting',
      };
    }

    // Email
    if (item.receivedDateTime) {
      return {
        date: String(item.receivedDateTime),
        title: String(item.subject || 'Email'),
        type: 'email',
      };
    }

    // File
    if (item.lastModifiedDateTime) {
      return {
        date: String(item.lastModifiedDateTime),
        title: String(item.name || 'File'),
        type: 'file',
      };
    }

    return null;
  }

  /**
   * Generate recommendations based on results
   */
  private generateRecommendations(input: LayerInput, items: unknown[]): string[] {
    const recommendations: string[] = [];

    // Based on intent
    if (input.intent.type === 'find' && items.length === 0) {
      recommendations.push(
        input.language === 'de' ? 'Versuchen Sie andere Suchbegriffe' : 'Try different search terms'
      );
    }

    // Based on results
    if (items.length > 20) {
      recommendations.push(
        input.language === 'de'
          ? 'Grenzen Sie die Suche mit Zeitraum oder Person ein'
          : 'Narrow search with date range or person'
      );
    }

    return recommendations;
  }

  /**
   * Generate summary from facts
   */
  private generateSummary(
    input: LayerInput,
    facts: string[],
    itemCount: number,
    lang: SupportedLanguage
  ): string {
    if (itemCount === 0) {
      return lang === 'de'
        ? `Keine Ergebnisse gefunden für "${input.question}".`
        : `No results found for "${input.question}".`;
    }

    const countStr =
      lang === 'de'
        ? `${itemCount} Ergebnis${itemCount > 1 ? 'se' : ''}`
        : `${itemCount} result${itemCount > 1 ? 's' : ''}`;

    if (facts.length > 0) {
      return lang === 'de' ? `${countStr} gefunden. ${facts[0]}` : `Found ${countStr}. ${facts[0]}`;
    }

    return lang === 'de' ? `${countStr} gefunden.` : `Found ${countStr}.`;
  }

  /**
   * Estimate token count for items
   */
  private estimateTokenCount(items: unknown[]): number {
    const jsonStr = JSON.stringify(items);
    // Rough estimate: 4 chars per token
    return Math.ceil(jsonStr.length / 4);
  }

  /**
   * Get current layer
   */
  getCurrentLayer(): AnalysisDepth {
    return this.currentLayer;
  }

  /**
   * Get layer definition
   */
  getLayerDefinition(layer: AnalysisDepth) {
    return LAYER_DEFINITIONS[layer];
  }

  /**
   * Get configuration
   */
  getConfig(): Required<LayerConfig> {
    return this.config;
  }
}

export default AdaptiveLayerController;
