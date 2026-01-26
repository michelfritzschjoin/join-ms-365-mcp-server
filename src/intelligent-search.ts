/**
 * Search-First Strategy Engine - always start with search-query, then query specific products
 */

import GraphClient from './graph-client.js';
import logger from './logger.js';
import EntityExtractor, { type ExtractedInfo } from './entity-extractor.js';
import SynonymExpander from './synonym-expander.js';
import QueryRefiner from './query-refiner.js';
import LearningSystem, { type SearchResult } from './learning-system.js';

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
   * Execute search-first strategy
   */
  async execute(query: string, context?: SearchContext): Promise<SearchFirstResult> {
    logger.info(`Executing search-first strategy for query: "${query}"`);

    // 1. ALWAYS start with search-query
    const searchResults = await this.executeSearchQuery(query, context);

    // 2. Extract entities and keywords from search results
    const extractedInfo = this.entityExtractor.extractFromResults(searchResults.items);

    // 3. If no results, try query refinement
    if (searchResults.items.length === 0) {
      logger.info('No results from initial search, trying query refinement');
      const refinedQueries = await this.queryRefiner.refineQuery(query, true, context);

      // Try refined queries
      for (const refinedQuery of refinedQueries.slice(1)) {
        // Skip first as it's the original
        const refinedResults = await this.executeSearchQuery(refinedQuery, context);
        if (refinedResults.items.length > 0) {
          // Found results with refined query
          searchResults.items.push(...refinedResults.items);
          searchResults.sources.push(...refinedResults.sources);
          searchResults.query = refinedQuery;

          // Re-extract with new results
          const newExtracted = this.entityExtractor.extractFromResults(refinedResults.items);
          this.mergeExtractedInfo(extractedInfo, newExtracted);
          break; // Stop after first successful refinement
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
      const recommendedEntityTypes =
        context?.entityTypes ||
        this.learningSystem.getRecommendedEntityTypes(query, context?.sources?.join(','));

      // Build search request body
      const now = new Date();
      const startDate = new Date(now.getFullYear() - 1, now.getMonth(), now.getDate()).toISOString();
      const endDate = new Date(now.getFullYear() + 1, now.getMonth(), now.getDate()).toISOString();

      const requestBody = {
        requests: [
          {
            entityTypes: recommendedEntityTypes,
            query: {
              queryString: query,
            },
            from: 0,
            size: context?.maxResults || parseInt(process.env.MS365_MCP_MAX_RESULTS || '500', 10),
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
      if (url.includes('sharepoint.com')) return 'sharepoint';
      if (url.includes('teams.microsoft.com')) return 'teams';
      if (url.includes('outlook.office.com')) return 'outlook';
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
