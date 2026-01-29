/**
 * UQAS Pro - Super-Tools Integration
 *
 * Integrates UQAS bilingual support into the existing Super-Tools system:
 * - Language detection for user queries
 * - Cross-language search expansion
 * - Bilingual intent recognition
 * - Token-optimized response formatting
 * - Result caching
 */

import { LanguageDetector, type LanguageDetectionResult } from '../i18n/language-detector.js';
import { BilingualThesaurus } from '../i18n/bilingual-thesaurus.js';
import { BilingualIntentRecognizer, type BilingualIntent } from '../i18n/intent-patterns.js';
import {
  BilingualEntityRecognizer,
  type ExtractedEntity,
  type TemporalExpression,
} from '../i18n/entity-patterns.js';
import {
  BilingualResponseBuilder,
  type ResponseData,
  type TimelineEntry,
} from '../i18n/response-templates.js';
import {
  CrossLanguageSearchExpander,
  type SearchQuerySet,
} from '../search/cross-language-expander.js';
import { TokenController } from '../core/token-controller.js';
import { CacheManager } from '../core/cache-manager.js';
import type { SupportedLanguage } from '../i18n/index.js';
import type { SourceInfo, DocumentLink } from '../i18n/response-templates.js';
import DownloadLinkGenerator from '../../download-link-generator.js';
import { getRequestTokens } from '../../request-context.js';
import DataAggregator from '../../data-aggregator.js';
import logger from '../../logger.js';

/**
 * UQAS analysis result
 */
export interface UQASAnalysis {
  /** Detected language */
  language: SupportedLanguage;
  /** Language detection confidence */
  languageConfidence: number;
  /** Has code-switching (mixed language) */
  hasCodeSwitch: boolean;
  /** Recognized intent */
  intent: BilingualIntent;
  /** Extracted entities */
  entities: ExtractedEntity[];
  /** Temporal expression if found */
  temporal: TemporalExpression | null;
  /** Expanded search queries */
  searchQueries: SearchQuerySet;
  /** Original query */
  originalQuery: string;
  /** Processing time in ms */
  processingTime: number;
}

/**
 * Formatted response options
 */
export interface FormatOptions {
  /** Target language (defaults to detected) */
  language?: SupportedLanguage;
  /** Maximum tokens */
  maxTokens?: number;
  /** Use compact format */
  compact?: boolean;
}

/**
 * UQASIntegration - Singleton for Super-Tools integration
 */
export class UQASIntegration {
  private static instance: UQASIntegration;

  private languageDetector: LanguageDetector;
  private thesaurus: BilingualThesaurus;
  private intentRecognizer: BilingualIntentRecognizer;
  private entityRecognizer: BilingualEntityRecognizer;
  private searchExpander: CrossLanguageSearchExpander;
  private tokenController: TokenController;
  private cache: CacheManager<UQASAnalysis>;
  private dataAggregator: DataAggregator;
  private downloadLinkGenerator: DownloadLinkGenerator | null = null;

  private constructor() {
    this.languageDetector = new LanguageDetector();
    this.thesaurus = new BilingualThesaurus();
    this.intentRecognizer = new BilingualIntentRecognizer();
    this.entityRecognizer = new BilingualEntityRecognizer();
    this.searchExpander = new CrossLanguageSearchExpander(this.thesaurus);
    this.tokenController = new TokenController({ maxTokens: 1500 });
    this.cache = new CacheManager<UQASAnalysis>({ defaultTTL: 300 });
    this.dataAggregator = new DataAggregator();
  }

  /**
   * Set DownloadLinkGenerator (called from super-tools integration)
   */
  setDownloadLinkGenerator(generator: DownloadLinkGenerator | null): void {
    this.downloadLinkGenerator = generator;
  }

  /**
   * Get singleton instance
   */
  static getInstance(): UQASIntegration {
    if (!UQASIntegration.instance) {
      UQASIntegration.instance = new UQASIntegration();
    }
    return UQASIntegration.instance;
  }

  /**
   * Analyze a user query with full UQAS processing
   */
  analyze(query: string): UQASAnalysis {
    const startTime = Date.now();

    // Check cache
    const cacheKey = this.cache.generateKey(query);
    const cached = this.cache.get(cacheKey);
    if (cached) {
      return cached;
    }

    // Detect language
    const langResult = this.languageDetector.detect(query);
    const lang = langResult.lang;

    // Recognize intent
    const intent = this.intentRecognizer.recognize(query);

    // Extract entities
    const entities = this.entityRecognizer.extractAll(query, lang);

    // Extract temporal expression
    const temporal = this.entityRecognizer.extractTemporal(query, lang);

    // Expand search queries
    const searchQueries = this.searchExpander.expand(query, lang);

    const analysis: UQASAnalysis = {
      language: lang,
      languageConfidence: langResult.confidence,
      hasCodeSwitch: langResult.hasCodeSwitch,
      intent,
      entities,
      temporal,
      searchQueries,
      originalQuery: query,
      processingTime: Date.now() - startTime,
    };

    // Cache result
    this.cache.set(cacheKey, analysis, { language: lang, query });

    return analysis;
  }

  /**
   * Detect language only (fast path)
   */
  detectLanguage(text: string): LanguageDetectionResult {
    return this.languageDetector.detect(text);
  }

  /**
   * Get search query variants for both languages
   */
  getSearchVariants(query: string, sourceLang?: SupportedLanguage): SearchQuerySet {
    const lang = sourceLang ?? this.languageDetector.detect(query).lang;
    return this.searchExpander.expand(query, lang);
  }

  /**
   * Get intent for a query
   */
  getIntent(query: string): BilingualIntent {
    return this.intentRecognizer.recognize(query);
  }

  /**
   * Extract entities from query
   */
  getEntities(query: string, lang?: SupportedLanguage): ExtractedEntity[] {
    const detectedLang = lang ?? this.languageDetector.detect(query).lang;
    return this.entityRecognizer.extractAll(query, detectedLang);
  }

  /**
   * Get synonyms for a word in both languages
   */
  getSynonyms(word: string): string[] {
    return this.thesaurus.getAllVariants(word);
  }

  /**
   * Format response for user in their language
   */
  formatResponse(data: ResponseData, options: FormatOptions = {}): string {
    const lang = options.language ?? 'de';
    const builder = new BilingualResponseBuilder(lang);

    // Apply token optimization if needed
    if (options.maxTokens) {
      const controller = new TokenController({ maxTokens: options.maxTokens });
      data = controller.optimize(data);
    }

    if (options.compact) {
      return builder.buildCompactResponse(data);
    }

    return builder.buildResponse(data);
  }

  /**
   * Create response data from search results
   */
  async createResponseData(
    results: unknown[],
    analysis: UQASAnalysis,
    sourceCount: number = 1
  ): Promise<ResponseData> {
    const facts = this.extractFacts(results, analysis.language);
    const timeline = this.extractTimeline(results);
    const confidence = this.calculateConfidence(results, analysis);

    // Generate summary
    const summary = this.generateSummary(results, analysis);

    // Extract sources and document links
    const sources = await this.extractSources(results);
    const documentLinks = await this.extractDocumentLinks(results);
    const importantDocuments = await this.identifyImportantDocuments(results);

    return {
      summary,
      confidence,
      sourceCount,
      depth: 1,
      facts,
      timeline,
      recommendations: this.generateRecommendations(analysis),
      sources,
      documentLinks,
      importantDocuments,
    };
  }

  /**
   * Extract sources from results
   */
  private async extractSources(results: unknown[]): Promise<SourceInfo[]> {
    const sources: SourceInfo[] = [];
    const seen = new Set<string>();

    for (const item of results) {
      if (typeof item !== 'object' || item === null) continue;
      const obj = item as Record<string, unknown>;

      const entityType = (obj['@odata.type'] as string) || 'unknown';
      const name =
        (obj.name as string) ||
        (obj.subject as string) ||
        (obj.title as string) ||
        (obj.displayName as string) ||
        'Unknown';
      const webUrl = (obj.webUrl as string) || (obj.webLink as string) || undefined;
      const rank = (obj.rank as number) || 0.5;

      // Create unique key
      const key = webUrl || `${entityType}:${name}`;
      if (seen.has(key)) continue;
      seen.add(key);

      sources.push({
        type: entityType,
        name,
        webUrl,
        relevance: rank,
      });
    }

    // Sort by relevance
    sources.sort((a, b) => b.relevance - a.relevance);

    return sources;
  }

  /**
   * Extract document links from results
   */
  private async extractDocumentLinks(results: unknown[]): Promise<DocumentLink[]> {
    const links: DocumentLink[] = [];
    const driveItems: unknown[] = [];

    // Collect all driveItems
    for (const item of results) {
      if (typeof item === 'object' && item !== null) {
        const obj = item as Record<string, unknown>;
        const entityType = obj['@odata.type'] as string | undefined;
        const isFile =
          entityType?.includes('driveItem') ||
          entityType?.includes('listItem') ||
          obj['file'] !== undefined ||
          obj['webUrl']?.toString().includes('/sites/') ||
          obj['webUrl']?.toString().includes('/drives/');

        if (isFile) {
          driveItems.push(item);
        }
      }
    }

    // Generate download links if generator available
    if (this.downloadLinkGenerator && driveItems.length > 0) {
      try {
        const requestTokens = getRequestTokens();
        const accessToken = requestTokens?.accessToken;

        const enrichedResults = await this.downloadLinkGenerator.addDownloadLinksToResults(
          driveItems,
          accessToken
        );

        for (const item of enrichedResults) {
          if (typeof item === 'object' && item !== null) {
            const obj = item as Record<string, unknown>;
            const downloadLink = obj.downloadLink as
              | { fileName: string; downloadUrl: string; webUrl?: string }
              | undefined;

            if (downloadLink) {
              links.push({
                fileName: downloadLink.fileName,
                webUrl: downloadLink.webUrl || (obj.webUrl as string) || '',
                downloadUrl: downloadLink.downloadUrl,
                type: (obj['@odata.type'] as string) || 'driveItem',
              });
            } else if (obj.webUrl) {
              const name = (obj.name as string) || (obj.title as string) || 'Unknown';
              links.push({
                fileName: name,
                webUrl: obj.webUrl as string,
                type: (obj['@odata.type'] as string) || 'unknown',
              });
            }
          }
        }
      } catch (error) {
        logger.warn(`Failed to generate document links: ${error}`);
      }
    } else {
      // Fallback: extract webUrls without download links
      for (const item of driveItems) {
        if (typeof item === 'object' && item !== null) {
          const obj = item as Record<string, unknown>;
          if (obj.webUrl) {
            const name = (obj.name as string) || (obj.title as string) || 'Unknown';
            links.push({
              fileName: name,
              webUrl: obj.webUrl as string,
              type: (obj['@odata.type'] as string) || 'unknown',
            });
          }
        }
      }
    }

    return links;
  }

  /**
   * Identify important documents based on relevance
   */
  private async identifyImportantDocuments(results: unknown[]): Promise<SourceInfo[]> {
    // Convert results to AggregatedItem format for DataAggregator
    const aggregatedItems = results.map((item) => {
      if (typeof item === 'object' && item !== null) {
        const obj = item as Record<string, unknown>;
        const id = (obj.id as string) || String(Math.random());
        const relevanceScore = (obj.rank as number) || 0.5;
        const source = (obj['@odata.type'] as string) || 'unknown';
        const timestamp = this.extractTimestamp(obj);

        return {
          id,
          data: item,
          relevanceScore,
          source,
          timestamp,
        };
      }
      return {
        id: String(Math.random()),
        data: item,
        relevanceScore: 0.5,
        source: 'unknown',
      };
    });

    // Use DataAggregator to identify important documents
    const important = this.dataAggregator.identifyImportantDocuments(aggregatedItems, 0.7, 10);

    // Convert back to SourceInfo format
    const sources: SourceInfo[] = [];
    for (const item of important) {
      if (typeof item.data === 'object' && item.data !== null) {
        const obj = item.data as Record<string, unknown>;
        const name =
          (obj.name as string) ||
          (obj.subject as string) ||
          (obj.title as string) ||
          (obj.displayName as string) ||
          'Unknown';
        const webUrl = (obj.webUrl as string) || (obj.webLink as string) || undefined;

        sources.push({
          type: item.source,
          name,
          webUrl,
          relevance: item.relevanceScore,
        });
      }
    }

    return sources;
  }

  /**
   * Extract timestamp from object
   */
  private extractTimestamp(obj: Record<string, unknown>): Date | undefined {
    const timeFields = [
      'createdDateTime',
      'lastModifiedDateTime',
      'start',
      'end',
      'sentDateTime',
      'receivedDateTime',
    ];

    for (const field of timeFields) {
      if (typeof obj[field] === 'string') {
        const date = new Date(obj[field] as string);
        if (!isNaN(date.getTime())) {
          return date;
        }
      } else if (typeof obj[field] === 'object' && obj[field] !== null) {
        const dateObj = obj[field] as Record<string, unknown>;
        if (typeof dateObj.dateTime === 'string') {
          const date = new Date(dateObj.dateTime);
          if (!isNaN(date.getTime())) {
            return date;
          }
        }
      }
    }

    return undefined;
  }

  /**
   * Extract facts from results
   */
  private extractFacts(results: unknown[], lang: SupportedLanguage): string[] {
    const facts: string[] = [];

    for (const item of results.slice(0, 5)) {
      if (typeof item !== 'object' || item === null) continue;
      const obj = item as Record<string, unknown>;

      // Email
      if (obj.subject && obj.from) {
        const from = this.extractName(obj.from);
        const label = lang === 'de' ? 'von' : 'from';
        facts.push(`📧 ${obj.subject} (${label} ${from})`);
      }
      // Event
      else if (obj.subject && obj.start) {
        const date = this.formatDateShort(obj.start, lang);
        facts.push(`📅 ${obj.subject} (${date})`);
      }
      // File
      else if (obj.name && obj.webUrl) {
        facts.push(`📁 ${obj.name}`);
      }
      // Task
      else if (obj.title) {
        facts.push(`✅ ${obj.title}`);
      }
    }

    return facts;
  }

  /**
   * Extract timeline from results
   */
  private extractTimeline(results: unknown[]): TimelineEntry[] {
    const entries: TimelineEntry[] = [];

    for (const item of results) {
      if (typeof item !== 'object' || item === null) continue;
      const obj = item as Record<string, unknown>;

      let date: string | undefined;
      let title: string | undefined;
      let type: TimelineEntry['type'] = 'other';

      if (obj.start) {
        const startObj = obj.start as Record<string, unknown>;
        date = String(startObj.dateTime || obj.start);
        title = String(obj.subject || 'Event');
        type = 'meeting';
      } else if (obj.receivedDateTime) {
        date = String(obj.receivedDateTime);
        title = String(obj.subject || 'Email');
        type = 'email';
      } else if (obj.lastModifiedDateTime) {
        date = String(obj.lastModifiedDateTime);
        title = String(obj.name || 'File');
        type = 'file';
      }

      if (date && title) {
        entries.push({ date, title, type });
      }
    }

    // Sort by date
    entries.sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime());

    return entries.slice(0, 10);
  }

  /**
   * Calculate confidence score
   */
  private calculateConfidence(results: unknown[], analysis: UQASAnalysis): number {
    let confidence = 0;

    // Base confidence from result count
    if (results.length >= 10) confidence += 0.4;
    else if (results.length >= 5) confidence += 0.3;
    else if (results.length >= 1) confidence += 0.2;

    // Bonus for intent match
    confidence += analysis.intent.confidence * 0.3;

    // Bonus for entity matches
    const entityMatches = this.countEntityMatches(results, analysis.entities);
    confidence += Math.min(0.2, entityMatches * 0.05);

    // Language confidence bonus
    confidence += analysis.languageConfidence * 0.1;

    return Math.min(1, confidence);
  }

  /**
   * Count entity matches in results
   */
  private countEntityMatches(results: unknown[], entities: ExtractedEntity[]): number {
    const resultStr = JSON.stringify(results).toLowerCase();
    return entities.filter((e) => resultStr.includes(e.value.toLowerCase())).length;
  }

  /**
   * Generate summary text
   */
  private generateSummary(results: unknown[], analysis: UQASAnalysis): string {
    const count = results.length;
    const lang = analysis.language;

    if (count === 0) {
      return lang === 'de'
        ? `Keine Ergebnisse gefunden für "${analysis.originalQuery}".`
        : `No results found for "${analysis.originalQuery}".`;
    }

    const countText =
      lang === 'de'
        ? `${count} Ergebnis${count > 1 ? 'se' : ''}`
        : `${count} result${count > 1 ? 's' : ''}`;

    return lang === 'de' ? `${countText} gefunden.` : `Found ${countText}.`;
  }

  /**
   * Generate recommendations
   */
  private generateRecommendations(analysis: UQASAnalysis): string[] {
    const recs: string[] = [];
    const lang = analysis.language;

    // Based on intent
    const actions = analysis.intent.suggestedActions;
    if (actions.length > 0) {
      const toolMap: Record<string, { de: string; en: string }> = {
        'list-calendar-events': {
          de: 'Nutzen Sie "calendar" mit action "list" für mehr Termine',
          en: 'Use "calendar" with action "list" for more events',
        },
        'get-mail-message': {
          de: 'Nutzen Sie "mail" mit action "get" für E-Mail-Details',
          en: 'Use "mail" with action "get" for email details',
        },
        'search-files': {
          de: 'Nutzen Sie "files" mit action "search" für Dateisuche',
          en: 'Use "files" with action "search" for file search',
        },
      };

      for (const action of actions.slice(0, 2)) {
        const rec = toolMap[action];
        if (rec) {
          recs.push(rec[lang]);
        }
      }
    }

    return recs;
  }

  /**
   * Extract name from object
   */
  private extractName(obj: unknown): string {
    if (typeof obj === 'string') return obj;
    if (typeof obj === 'object' && obj !== null) {
      const o = obj as Record<string, unknown>;
      if (o.emailAddress && typeof o.emailAddress === 'object') {
        return String((o.emailAddress as Record<string, unknown>).name || 'Unknown');
      }
      return String(o.name || o.displayName || 'Unknown');
    }
    return 'Unknown';
  }

  /**
   * Format date in short form
   */
  private formatDateShort(dateObj: unknown, lang: SupportedLanguage): string {
    let dateStr: string;
    if (typeof dateObj === 'string') {
      dateStr = dateObj;
    } else if (typeof dateObj === 'object' && dateObj !== null) {
      dateStr = String((dateObj as Record<string, unknown>).dateTime || dateObj);
    } else {
      return '';
    }

    const date = new Date(dateStr);
    if (lang === 'de') {
      return `${date.getDate()}.${date.getMonth() + 1}.`;
    }
    return `${date.getMonth() + 1}/${date.getDate()}`;
  }

  /**
   * Get cache statistics
   */
  getCacheStats() {
    return this.cache.getStats();
  }

  /**
   * Clear cache
   */
  clearCache() {
    this.cache.clear();
  }

  /**
   * Create thinking steps for analysis
   */
  createThinkingSteps(analysis: UQASAnalysis): string[] {
    const steps: string[] = [];
    const langLabel = analysis.language === 'de' ? 'Deutsch' : 'English';

    steps.push(`🌐 Language: ${langLabel} (${Math.round(analysis.languageConfidence * 100)}%)`);
    steps.push(
      `🎯 Intent: ${analysis.intent.type} (${Math.round(analysis.intent.confidence * 100)}%)`
    );

    if (analysis.entities.length > 0) {
      const entityStr = analysis.entities.map((e) => `${e.value} (${e.type})`).join(', ');
      steps.push(`🏷️ Entities: ${entityStr}`);
    }

    if (analysis.temporal) {
      steps.push(`📅 Temporal: ${analysis.temporal.original} → ${analysis.temporal.normalized}`);
    }

    if (analysis.searchQueries.crossLangVariants.length > 0) {
      steps.push(
        `🔄 Cross-lang queries: ${analysis.searchQueries.crossLangVariants.slice(0, 2).join(', ')}`
      );
    }

    return steps;
  }
}

// Export singleton getter
export const getUQAS = () => UQASIntegration.getInstance();

// Export convenience functions
export function analyzeQuery(query: string): UQASAnalysis {
  return getUQAS().analyze(query);
}

export function detectLanguage(text: string): LanguageDetectionResult {
  return getUQAS().detectLanguage(text);
}

export function getSearchVariants(query: string, lang?: SupportedLanguage): SearchQuerySet {
  return getUQAS().getSearchVariants(query, lang);
}

export function formatBilingualResponse(data: ResponseData, options?: FormatOptions): string {
  return getUQAS().formatResponse(data, options);
}

export default UQASIntegration;
