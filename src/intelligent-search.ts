/**
 * Search-First Strategy Engine - always start with search-query, then query specific products
 */

import GraphClient from './graph-client.js';
import logger from './logger.js';
import EntityExtractor, { type ExtractedInfo } from './entity-extractor.js';
import SynonymExpander from './synonym-expander.js';
import QueryRefiner from './query-refiner.js';
import LearningSystem, { type SearchResult } from './learning-system.js';
import { validateEntityTypeCombinations } from './utils/entity-type-validator.js';
import { getQueryOptimizer } from './query-optimizer.js';
import { CrossLanguageSearchExpander } from './uqas/search/index.js';
import { LanguageDetector } from './uqas/i18n/index.js';
import type { SupportedLanguage } from './uqas/i18n/index.js';
import { getMaxResults } from './perf-config.js';

export interface SearchContext {
  entityTypes?: string[];
  sources?: string[];
  timeRange?: string;
  maxResults?: number;
}

export interface SearchFirstResult {
  searchResults: SearchResult;
  extractedInfo: ExtractedInfo;
  specificResults: Record<string, unknown[]>;
  totalItems: number;
}

export class SearchFirstStrategy {
  private graphClient: GraphClient;
  private entityExtractor: EntityExtractor;
  private synonymExpander: SynonymExpander;
  private queryRefiner: QueryRefiner;
  private learningSystem: LearningSystem;
  private crossLanguageExpander: CrossLanguageSearchExpander | null = null;
  private languageDetector: LanguageDetector | null = null;

  constructor(
    graphClient: GraphClient,
    entityExtractor: EntityExtractor,
    synonymExpander: SynonymExpander,
    queryRefiner: QueryRefiner,
    learningSystem: LearningSystem
  ) {
    this.graphClient = graphClient;
    this.entityExtractor = entityExtractor;
    this.synonymExpander = synonymExpander;
    this.queryRefiner = queryRefiner;
    this.learningSystem = learningSystem;
  }

  /**
   * Lazy-init cross-language expander and language detector for bilingual query variants.
   */
  private getCrossLanguageExpander(): CrossLanguageSearchExpander {
    if (!this.crossLanguageExpander) {
      this.crossLanguageExpander = new CrossLanguageSearchExpander();
    }
    return this.crossLanguageExpander;
  }

  private getLanguageDetector(): LanguageDetector {
    if (!this.languageDetector) {
      this.languageDetector = new LanguageDetector();
    }
    return this.languageDetector;
  }

  /**
   * Generate simplified query variants (drop one word, key terms only) for fallback.
   */
  private buildSimplifiedQueries(query: string): string[] {
    const words = query
      .trim()
      .split(/\s+/)
      .filter((w) => w.length > 0);
    if (words.length < 3) return [];

    const simplified: string[] = [];
    for (let i = 0; i < words.length; i++) {
      const without = words.filter((_, j) => j !== i).join(' ');
      if (without.length >= 2) simplified.push(without);
    }
    if (words.length >= 3) {
      simplified.push(words.slice(0, 2).join(' '));
      if (words.length >= 4) simplified.push(words.slice(0, 3).join(' '));
    }
    return [...new Set(simplified)].slice(0, 5);
  }

  /**
   * Build all query candidates: learned variants, cross-language, optimizer, synonyms, simplified.
   */
  private buildQueryCandidates(
    primaryQuery: string,
    originalQuery: string,
    optimizerVariants: string[]
  ): string[] {
    const seen = new Set<string>([
      primaryQuery.toLowerCase().trim(),
      originalQuery.toLowerCase().trim(),
    ]);
    const candidates: string[] = [];

    const add = (v: string) => {
      const key = v.trim().toLowerCase();
      if (key && !seen.has(key)) {
        seen.add(key);
        candidates.push(v.trim());
      }
    };

    // 1. Learned variants from past successful searches (highest priority)
    const learnedLimit = parseInt(process.env.MS365_MCP_SEARCH_LEARNED_VARIANTS || '3', 10);
    if (learnedLimit > 0) {
      const learned = this.learningSystem.getSuggestedQueryVariants(primaryQuery, learnedLimit);
      for (const v of learned) add(v);
    }

    // 2. Cross-language variants (DE/EN thesaurus)
    if (process.env.MS365_MCP_SEARCH_CROSS_LANGUAGE !== 'false') {
      try {
        const detector = this.getLanguageDetector();
        const expander = this.getCrossLanguageExpander();
        const lang = detector.detect(primaryQuery).lang as SupportedLanguage;
        const set = expander.expand(primaryQuery, lang);
        for (const v of set.sourceLangVariants.slice(0, 3)) add(v);
        for (const v of set.crossLangVariants.slice(0, 3)) add(v);
      } catch (err) {
        logger.debug(`Cross-language expansion skipped: ${err}`);
      }
    }

    // 3. Optimizer variants (cross-lang, typos, abbreviations, NLP)
    for (const v of optimizerVariants) add(v);

    // 4. Synonym expander (business terms, learned synonyms)
    const maxSynonym = parseInt(process.env.MS365_MCP_SEARCH_MAX_SYNONYM_VARIANTS || '5', 10);
    const synonymVariants = this.synonymExpander.expandQuery(primaryQuery, 'search');
    for (const v of synonymVariants.slice(0, maxSynonym)) add(v);

    // 5. Simplified sub-queries (drop-one-word, key terms) for fallback
    if (process.env.MS365_MCP_SEARCH_SIMPLIFIED_FALLBACK !== 'false') {
      for (const v of this.buildSimplifiedQueries(primaryQuery)) add(v);
    }

    const maxCandidates = parseInt(process.env.MS365_MCP_SEARCH_MAX_CANDIDATES || '18', 10);
    return candidates.slice(0, maxCandidates);
  }

  /**
   * Build a single OR-query string from primary + top variants (for one broad API call).
   */
  private buildOrQuery(primaryQuery: string, variants: string[], maxTerms = 3): string {
    const terms = [primaryQuery, ...variants].slice(0, maxTerms);
    const escaped = terms.map((q) => {
      const t = q.trim();
      if (t.includes(' ') || t.includes('"')) return `"${t.replace(/"/g, '\\"')}"`;
      return t;
    });
    return escaped.join(' OR ');
  }

  /**
   * Returns a stable key for a search hit resource (for deduplication).
   */
  private getResourceKey(resource: Record<string, unknown>): string {
    const id = resource['id'];
    if (typeof id === 'string') return id;
    const webUrl = resource['webUrl'];
    if (typeof webUrl === 'string') return webUrl;
    return JSON.stringify(resource);
  }

  /**
   * Merge search results from multiple queries and deduplicate by resource key.
   */
  private mergeSearchResults(
    base: SearchResult,
    additional: SearchResult[],
    maxItems: number
  ): void {
    const keySet = new Set<string>();
    for (const item of base.items) {
      if (typeof item === 'object' && item !== null) {
        keySet.add(this.getResourceKey(item as Record<string, unknown>));
      }
    }

    for (const result of additional) {
      for (const item of result.items) {
        if (base.items.length >= maxItems) break;
        if (typeof item !== 'object' || item === null) continue;
        const key = this.getResourceKey(item as Record<string, unknown>);
        if (!keySet.has(key)) {
          keySet.add(key);
          base.items.push(item);
        }
      }
      for (const source of result.sources) {
        if (source && !base.sources.includes(source)) {
          base.sources.push(source);
        }
      }
    }

    base.totalResults = base.items.length;
  }

  /**
   * Execute search-first strategy
   * Uses automatic query optimization, then tries multiple query variants (optimizer + synonyms).
   * Optionally enriches results by running synonym queries in parallel when primary returns few results.
   */
  async execute(query: string, context?: SearchContext): Promise<SearchFirstResult> {
    logger.info(`Executing search-first strategy for query: "${query}"`);

    // 0. Auto-optimize query (NLP, synonyms, identifiers, stopwords, abbreviations, typos)
    const optimizer = getQueryOptimizer();
    const optimized = optimizer.optimizeQuery(query, {
      tool: 'search',
      entityTypes: context?.entityTypes,
    });
    const searchQuery = optimized.optimizedQuery;

    // 1. ALWAYS start with search-query using optimized query
    const searchResults = await this.executeSearchQuery(searchQuery, context);

    // 2. Extract entities and keywords from search results
    const extractedInfo = this.entityExtractor.extractFromResults(searchResults.items);

    // 3a. Build intelligent query candidates (optimizer + synonym variants)
    const variantsToTry = this.buildQueryCandidates(searchQuery, query, optimized.variants);

    // 3b. If no results, try OR-query first (one broad request), then variant queries
    if (searchResults.items.length === 0) {
      const useOrQuery = process.env.MS365_MCP_SEARCH_USE_OR_QUERY !== 'false';
      if (useOrQuery && variantsToTry.length >= 1) {
        const orQuery = this.buildOrQuery(searchQuery, variantsToTry, 3);
        if (orQuery !== searchQuery) {
          const orResults = await this.executeSearchQuery(orQuery, context);
          if (orResults.items.length > 0) {
            logger.info(`OR-query returned results: "${orQuery.slice(0, 80)}..."`);
            searchResults.items.push(...orResults.items);
            searchResults.sources.push(...orResults.sources);
            searchResults.query = orQuery;
            searchResults.totalResults = searchResults.items.length;
            const newExtracted = this.entityExtractor.extractFromResults(orResults.items);
            this.mergeExtractedInfo(extractedInfo, newExtracted);
          }
        }
      }

      if (searchResults.items.length === 0) {
        const refinerQueries = (await this.queryRefiner.refineQuery(query, true, context)).slice(1);
        const allCandidates = [...variantsToTry];
        for (const r of refinerQueries) {
          const key = r.trim().toLowerCase();
          if (key && key !== searchQuery.toLowerCase() && key !== query.toLowerCase()) {
            allCandidates.push(r.trim());
          }
        }

        for (const variantQuery of allCandidates) {
          if (variantQuery === searchQuery || variantQuery === query) continue;
          const refinedResults = await this.executeSearchQuery(variantQuery, context);
          if (refinedResults.items.length > 0) {
            logger.info(`Variant query returned results: "${variantQuery}"`);
            searchResults.items.push(...refinedResults.items);
            searchResults.sources.push(...refinedResults.sources);
            searchResults.query = variantQuery;

            const newExtracted = this.entityExtractor.extractFromResults(refinedResults.items);
            this.mergeExtractedInfo(extractedInfo, newExtracted);
            break;
          }
        }
      }
    } else if (
      process.env.MS365_MCP_SEARCH_INTELLIGENT_MERGE !== 'false' &&
      variantsToTry.length > 0
    ) {
      // 3c. Optional: enrich with synonym/variant queries when primary returned few results
      const mergeThreshold = parseInt(
        process.env.MS365_MCP_SEARCH_SYNONYM_MERGE_THRESHOLD || '5',
        10
      );
      const maxResults = context?.maxResults ?? getMaxResults();

      if (searchResults.items.length < mergeThreshold) {
        const parallelMergeCount = parseInt(
          process.env.MS365_MCP_SEARCH_PARALLEL_MERGE_COUNT || '3',
          10
        );
        const parallelQueries = variantsToTry.slice(0, Math.max(1, parallelMergeCount));
        const extraResults = await Promise.all(
          parallelQueries.map((q) => this.executeSearchQuery(q, context))
        );
        const withResults = extraResults.filter((r) => r.items.length > 0);
        if (withResults.length > 0) {
          logger.info(
            `Intelligent merge: primary had ${searchResults.items.length} items, merging ${withResults.length} synonym query result(s)`
          );
          this.mergeSearchResults(searchResults, withResults, maxResults);
          for (const r of withResults) {
            const newExtracted = this.entityExtractor.extractFromResults(r.items);
            this.mergeExtractedInfo(extractedInfo, newExtracted);
          }
        }
      }
    }

    // 4. Query specific products based on extracted info
    const specificResults = await this.querySpecificProducts(extractedInfo, context);

    // 5. Learn from results
    await this.learningSystem.learnFromSearch(
      query,
      searchResults,
      undefined,
      context?.entityTypes?.join(',')
    );

    // 6. Aggregate and return
    const totalItems = searchResults.items.length + Object.values(specificResults).flat().length;

    return {
      searchResults,
      extractedInfo,
      specificResults,
      totalItems,
    };
  }

  /**
   * Execute search-query via Graph API
   */
  async executeSearchQuery(query: string, context?: SearchContext): Promise<SearchResult> {
    try {
      // Get recommended entity types from learning system
      let recommendedEntityTypes =
        context?.entityTypes ||
        this.learningSystem.getRecommendedEntityTypes(query, context?.sources?.join(','));

      // If no specific types recommended, use default (configurable via MS365_MCP_SEARCH_DEFAULT_ENTITY_TYPES)
      if (!recommendedEntityTypes || recommendedEntityTypes.length === 0) {
        const envDefault = process.env.MS365_MCP_SEARCH_DEFAULT_ENTITY_TYPES;
        recommendedEntityTypes = envDefault
          ? envDefault
              .split(',')
              .map((s) => s.trim())
              .filter(Boolean)
          : ['message', 'event', 'driveItem', 'site', 'list', 'listItem'];
        logger.debug('Using default entity types for search', {
          types: recommendedEntityTypes,
        });
      }

      // Validate entity types to ensure compatibility
      recommendedEntityTypes = validateEntityTypeCombinations(recommendedEntityTypes);

      // Build search request body
      const now = new Date();
      const startDate = new Date(
        now.getFullYear() - 1,
        now.getMonth(),
        now.getDate()
      ).toISOString();
      const endDate = new Date(now.getFullYear() + 1, now.getMonth(), now.getDate()).toISOString();

      const requestBody = {
        requests: [
          {
            entityTypes: recommendedEntityTypes,
            query: {
              queryString: query,
            },
            from: 0,
            size: context?.maxResults || getMaxResults(),
            // Mandatory for 'event' entityType
            ...(recommendedEntityTypes.includes('event') && {
              timeContext: {
                startDateTime: startDate,
                endDateTime: endDate,
              },
            }),
            ...(context?.timeRange && { timeRange: context.timeRange }),
          },
        ],
      };

      // Execute search
      const response = await this.graphClient.makeRequest('/search/query', {
        method: 'POST',
        body: JSON.stringify(requestBody),
      });

      // Parse results
      const items: unknown[] = [];
      const sources: string[] = [];

      if (response && typeof response === 'object' && 'value' in response) {
        const searchResponse = response as {
          value: Array<{ hitsContainers?: Array<{ hits?: unknown[]; total?: number }> }>;
        };
        if (Array.isArray(searchResponse.value)) {
          for (const container of searchResponse.value) {
            if (container.hitsContainers && Array.isArray(container.hitsContainers)) {
              for (const hitsContainer of container.hitsContainers) {
                if (hitsContainer.hits && Array.isArray(hitsContainer.hits)) {
                  for (const hit of hitsContainer.hits) {
                    if (typeof hit === 'object' && hit !== null) {
                      const hitObj = hit as Record<string, unknown>;
                      if (hitObj['resource']) {
                        items.push(hitObj['resource']);
                        // Extract source
                        const resource = hitObj['resource'] as Record<string, unknown>;
                        const source = this.extractSourceFromResource(resource);
                        if (source && !sources.includes(source)) {
                          sources.push(source);
                        }
                      }
                    }
                  }
                }
              }
            }
          }
        }
      }

      return {
        items,
        sources,
        query,
        entityTypes: recommendedEntityTypes,
        totalResults: items.length,
      };
    } catch (error) {
      logger.error(`Search query failed: ${error}`);
      return {
        items: [],
        sources: [],
        query,
        entityTypes: context?.entityTypes,
        totalResults: 0,
      };
    }
  }

  /**
   * Query specific products based on extracted information
   */
  private async querySpecificProducts(
    extractedInfo: ExtractedInfo,
    context?: SearchContext
  ): Promise<Record<string, unknown[]>> {
    const results: Record<string, unknown[]> = {};

    // Query SharePoint sites
    if (extractedInfo.sites.length > 0) {
      try {
        for (const siteId of extractedInfo.sites.slice(0, 5)) {
          // Limit to 5 sites
          const siteResponse = await this.graphClient.makeRequest(`/sites/${siteId}`);
          if (siteResponse) {
            if (!results['sites']) {
              results['sites'] = [];
            }
            results['sites'].push(siteResponse);
          }
        }
      } catch (error) {
        logger.warn(`Failed to query sites: ${error}`);
      }
    }

    // Query drives/files
    if (extractedInfo.drives.length > 0 || extractedInfo.files.length > 0) {
      try {
        for (const driveId of extractedInfo.drives.slice(0, 3)) {
          // Limit to 3 drives
          const driveResponse = await this.graphClient.makeRequest(`/drives/${driveId}`);
          if (driveResponse) {
            if (!results['drives']) {
              results['drives'] = [];
            }
            results['drives'].push(driveResponse);
          }
        }

        for (const fileId of extractedInfo.files.slice(0, 10)) {
          // Limit to 10 files
          try {
            const fileResponse = await this.graphClient.makeRequest(
              `/drives/${extractedInfo.drives[0]}/items/${fileId}`
            );
            if (fileResponse) {
              if (!results['files']) {
                results['files'] = [];
              }
              results['files'].push(fileResponse);
            }
          } catch (error) {
            // Skip if file query fails
          }
        }
      } catch (error) {
        logger.warn(`Failed to query drives/files: ${error}`);
      }
    }

    // Query teams
    if (extractedInfo.teams.length > 0) {
      try {
        for (const teamId of extractedInfo.teams.slice(0, 3)) {
          // Limit to 3 teams
          const teamResponse = await this.graphClient.makeRequest(`/teams/${teamId}`);
          if (teamResponse) {
            if (!results['teams']) {
              results['teams'] = [];
            }
            results['teams'].push(teamResponse);
          }
        }
      } catch (error) {
        logger.warn(`Failed to query teams: ${error}`);
      }
    }

    // Query users
    if (extractedInfo.users.length > 0) {
      try {
        for (const userId of extractedInfo.users.slice(0, 5)) {
          // Limit to 5 users
          const userResponse = await this.graphClient.makeRequest(`/users/${userId}`);
          if (userResponse) {
            if (!results['users']) {
              results['users'] = [];
            }
            results['users'].push(userResponse);
          }
        }
      } catch (error) {
        logger.warn(`Failed to query users: ${error}`);
      }
    }

    return results;
  }

  /**
   * Extract source from resource
   */
  private extractSourceFromResource(resource: Record<string, unknown>): string | undefined {
    if (resource['@odata.type']) {
      const type = resource['@odata.type'] as string;
      if (type.includes('site')) return 'sharepoint';
      if (type.includes('driveItem')) return 'onedrive';
      if (type.includes('message')) return 'mail';
      if (type.includes('chatMessage')) return 'teams';
      if (type.includes('person')) return 'people';
      if (type.includes('event')) return 'calendar';
    }

    if (resource['webUrl']) {
      const url = resource['webUrl'] as string;
      // SECURITY: Use exact hostname matching with whitelist instead of substring matching
      try {
        const urlObj = new URL(url);
        const hostname = urlObj.hostname.toLowerCase();
        // Use exact matching or endsWith for specific domains to prevent bypass
        if (hostname === 'sharepoint.com' || hostname.endsWith('.sharepoint.com')) {
          return 'sharepoint';
        }
        if (hostname === 'teams.microsoft.com' || hostname.endsWith('.teams.microsoft.com')) {
          return 'teams';
        }
        if (hostname === 'outlook.office.com' || hostname.endsWith('.outlook.office.com')) {
          return 'outlook';
        }
      } catch {
        // Invalid URL, skip
      }
    }

    return undefined;
  }

  /**
   * Merge extracted info
   */
  private mergeExtractedInfo(target: ExtractedInfo, source: ExtractedInfo): void {
    // Merge entities
    target.entities.push(...source.entities);

    // Merge keywords
    for (const keyword of source.keywords) {
      if (!target.keywords.includes(keyword)) {
        target.keywords.push(keyword);
      }
    }

    // Merge IDs
    for (const [type, ids] of Object.entries(source.ids)) {
      if (!target.ids[type]) {
        target.ids[type] = [];
      }
      for (const id of ids) {
        if (!target.ids[type].includes(id)) {
          target.ids[type].push(id);
        }
      }
    }

    // Merge sites, drives, teams, users, files
    for (const site of source.sites) {
      if (!target.sites.includes(site)) {
        target.sites.push(site);
      }
    }
    for (const drive of source.drives) {
      if (!target.drives.includes(drive)) {
        target.drives.push(drive);
      }
    }
    for (const team of source.teams) {
      if (!target.teams.includes(team)) {
        target.teams.push(team);
      }
    }
    for (const user of source.users) {
      if (!target.users.includes(user)) {
        target.users.push(user);
      }
    }
    for (const file of source.files) {
      if (!target.files.includes(file)) {
        target.files.push(file);
      }
    }
  }
}

export default SearchFirstStrategy;
