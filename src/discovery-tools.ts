/**
 * Discovery Tools for intelligent business question answering
 */

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import logger from './logger.js';
import GraphClient from './graph-client.js';
import SearchFirstStrategy from './intelligent-search.js';
import EntityExtractor from './entity-extractor.js';
import SynonymExpander from './synonym-expander.js';
import QueryRefiner from './query-refiner.js';
import LearningSystem from './learning-system.js';
import KnowledgeBase from './knowledge-base.js';
import ToolCombiner from './tool-combiner.js';
import DataAggregator from './data-aggregator.js';
import DownloadLinkGenerator from './download-link-generator.js';
import DeepResearchEngine from './deep-research-engine.js';
import type { AppSecrets } from './secrets.js';

let searchStrategy: SearchFirstStrategy | null = null;
let deepResearchEngine: DeepResearchEngine | null = null;
let dataAggregator: DataAggregator | null = null;
let downloadLinkGenerator: DownloadLinkGenerator | null = null;

/**
 * Initialize discovery components
 */
function initializeDiscoveryComponents(graphClient: GraphClient, secrets: AppSecrets): void {
  const knowledgeBase = new KnowledgeBase();
  const synonymExpander = new SynonymExpander();
  const learningSystem = new LearningSystem(knowledgeBase, synonymExpander);
  const entityExtractor = new EntityExtractor();
  const queryRefiner = new QueryRefiner(synonymExpander);
  const toolCombiner = new ToolCombiner(graphClient);

  searchStrategy = new SearchFirstStrategy(
    graphClient,
    entityExtractor,
    synonymExpander,
    queryRefiner,
    learningSystem
  );

  dataAggregator = new DataAggregator();
  downloadLinkGenerator = new DownloadLinkGenerator(graphClient, secrets);

  deepResearchEngine = new DeepResearchEngine(
    graphClient,
    searchStrategy,
    entityExtractor,
    toolCombiner,
    dataAggregator,
    downloadLinkGenerator,
    secrets
  );
}

/**
 * Register all discovery tools
 */
export function registerDiscoveryTools(
  server: McpServer,
  graphClient: GraphClient,
  secrets: AppSecrets
): void {
  // Initialize components
  initializeDiscoveryComponents(graphClient, secrets);

  // ============================================================================
  // PRIMARY SEARCH TOOL - MUST BE CALLED FOR ANY USER QUESTION
  // ============================================================================

  server.tool(
    'ms365-search',
    `🔍 UNIVERSAL MICROSOFT 365 SEARCH - ALWAYS USE THIS TOOL FIRST!

This is the primary search tool for finding ANY information in Microsoft 365.
CRITICAL: You MUST call this tool for EVERY user question before answering from general knowledge!

This tool searches across ALL Microsoft 365 data sources:
- Emails and mail messages
- Files (OneDrive, SharePoint)
- Teams chat messages and channels
- Calendar events and meetings
- People and contacts
- OneNote pages
- SharePoint sites and lists

Examples of when to use this tool:
- "What do you know about DZBANK?" → Search for "DZBANK"
- "Tell me about Project Alpha" → Search for "Project Alpha"
- "What did we discuss about budget?" → Search for "budget"
- "Find info about customer ABC" → Search for "customer ABC"

The tool will ALWAYS return results or explicitly state that no information was found in Microsoft 365.
Even if no results are found, this confirms the information is not in the user's Microsoft 365 data.`,
    {
      query: z
        .string()
        .describe(
          'The search query - can be any topic, person, company, project, keyword, or phrase the user is asking about'
        ),
      entityTypes: z
        .array(z.string())
        .optional()
        .describe(
          'Optional: Specific entity types to search (message, driveItem, event, chatMessage, site, person). Defaults to all types.'
        ),
      maxResults: z
        .number()
        .optional()
        .describe('Maximum number of results to return (default: 50, max: 500)'),
    },
    async ({ query, entityTypes, maxResults }) => {
      if (!searchStrategy || !dataAggregator) {
        throw new Error('Search components not initialized');
      }

      logger.info(`MS365 Universal Search: "${query}"`);

      const configuredMaxResults = parseInt(process.env.MS365_MCP_MAX_RESULTS || '500', 10);
      const maxAggregateItems = parseInt(process.env.MS365_MCP_MAX_AGGREGATE_ITEMS || '500', 10);
      const effectiveMaxResults = Math.min(maxResults || 50, configuredMaxResults);

      // Default to all entity types for comprehensive search
      const searchEntityTypes = entityTypes || [
        'message',
        'driveItem',
        'event',
        'chatMessage',
        'site',
        'person',
        'listItem',
      ];

      const result = await searchStrategy.execute(query, {
        entityTypes: searchEntityTypes,
        maxResults: effectiveMaxResults,
      });

      // Aggregate results
      const aggregated = dataAggregator.aggregate(
        [
          {
            source: 'search',
            items: result.searchResults.items,
          },
          ...Object.entries(result.specificResults).map(([source, items]) => ({
            source,
            items: items as unknown[],
          })),
        ],
        {
          sortBy: 'relevance',
          sortOrder: 'desc',
          maxItems: Math.min(effectiveMaxResults, maxAggregateItems),
          deduplicate: true,
          formatForLLM: true,
        }
      );

      // Build response with explicit messaging
      const hasResults = result.totalItems > 0;

      const response = {
        query,
        searchPerformed: true,
        searchedSources: searchEntityTypes,
        totalItemsFound: result.totalItems,
        hasResults,
        message: hasResults
          ? `Found ${result.totalItems} items in Microsoft 365 matching "${query}".`
          : `No items found in Microsoft 365 matching "${query}". This query was searched across all connected Microsoft 365 data sources (emails, files, chats, calendar, etc.) but no matching content was found.`,
        sources: aggregated.sources,
        items: aggregated.items,
        formattedSummary: hasResults
          ? aggregated.formattedForLLM
          : `Microsoft 365 Search Results for "${query}": No matching items found in any data source. The user's Microsoft 365 tenant does not contain any emails, files, messages, events, or other content matching this search query.`,
        tip: hasResults
          ? 'Use the specific items above to answer the user question with information from their Microsoft 365 data.'
          : 'No Microsoft 365 data found. You may inform the user that no information about this topic exists in their Microsoft 365 environment, or answer from your general knowledge while noting that no specific company data was found.',
      };

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(response, null, 2),
          },
        ],
      };
    }
  );

  // Deep Research Tool
  server.tool(
    'deep-research',
    'Conduct comprehensive deep research on a question with multi-step reasoning, iterative refinement, and comprehensive data gathering. Can process more than 100 items and provides download links for files.',
    {
      question: z.string().describe('The research question to answer'),
      context: z.string().optional().describe('Additional context for the research'),
      maxDepth: z.number().optional().describe('Maximum research depth (default: 3)'),
      maxIterations: z.number().optional().describe('Maximum research iterations (default: 5)'),
      includeDownloadLinks: z
        .boolean()
        .optional()
        .describe('Include download links for files (default: false)'),
    },
    async ({ question, context, maxDepth, maxIterations, includeDownloadLinks }) => {
      if (!deepResearchEngine) {
        throw new Error('Deep research engine not initialized');
      }

      logger.info(`Deep research requested: "${question}"`);

      const result = await deepResearchEngine.research({
        question,
        context,
        maxDepth,
        maxIterations,
        includeDownloadLinks: includeDownloadLinks || false,
      });

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );

  // Discover Project Tool
  server.tool(
    'discover-project',
    'Discover comprehensive information about a project: files, sites, teams, tasks, meetings, and communications. Uses intelligent search-first strategy with synonym expansion.',
    {
      projectName: z.string().describe('Project name or identifier'),
      includeFiles: z.boolean().optional().describe('Include project files'),
      includeTasks: z.boolean().optional().describe('Include planner tasks'),
      includeMeetings: z.boolean().optional().describe('Include related meetings'),
      includeDownloadLinks: z.boolean().optional().describe('Include download links for files'),
    },
    async ({ projectName, includeFiles, includeTasks, includeMeetings, includeDownloadLinks }) => {
      if (!searchStrategy || !dataAggregator) {
        throw new Error('Discovery components not initialized');
      }

      logger.info(`Discovering project: "${projectName}"`);

      const maxResults = parseInt(process.env.MS365_MCP_MAX_RESULTS || '500', 10);
      const maxAggregateItems = parseInt(process.env.MS365_MCP_MAX_AGGREGATE_ITEMS || '500', 10);

      const result = await searchStrategy.execute(projectName, {
        entityTypes: ['driveItem', 'site', 'chatMessage', 'event', 'message'],
        maxResults,
      });

      // Aggregate results
      const aggregated = dataAggregator.aggregate(
        [
          {
            source: 'search',
            items: result.searchResults.items,
          },
          ...Object.entries(result.specificResults).map(([source, items]) => ({
            source,
            items: items as unknown[],
          })),
        ],
        {
          sortBy: 'relevance',
          sortOrder: 'desc',
          maxItems: maxAggregateItems,
          deduplicate: true,
          formatForLLM: true,
        }
      );

      // Add download links if requested
      let downloadLinks: Array<{ fileName: string; downloadUrl: string }> | undefined;
      if (includeDownloadLinks && downloadLinkGenerator) {
        const links = await downloadLinkGenerator.addDownloadLinksToResults(
          aggregated.items.map((i) => i.data)
        );
        downloadLinks = links
          .filter((l) => l.downloadLink)
          .map((l) => ({
            fileName: l.downloadLink!.fileName,
            downloadUrl: l.downloadLink!.downloadUrl,
          }));
      }

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                project: projectName,
                totalItems: result.totalItems,
                sources: aggregated.sources,
                items: aggregated.items,
                downloadLinks,
                formattedSummary: aggregated.formattedForLLM,
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );

  // Discover Person Tool
  server.tool(
    'discover-person',
    'Discover comprehensive information about a person: profile, communications, meetings, files, and team memberships. Uses intelligent search-first strategy.',
    {
      personName: z.string().describe('Person name or email'),
      includeCommunications: z.boolean().optional().describe('Include mail and chat messages'),
      includeMeetings: z.boolean().optional().describe('Include calendar events'),
      includeFiles: z.boolean().optional().describe('Include shared files'),
    },
    async ({ personName, includeCommunications, includeMeetings, includeFiles }) => {
      if (!searchStrategy || !dataAggregator) {
        throw new Error('Discovery components not initialized');
      }

      logger.info(`Discovering person: "${personName}"`);

      const maxResults = parseInt(process.env.MS365_MCP_MAX_RESULTS || '500', 10);
      const maxAggregateItems = parseInt(process.env.MS365_MCP_MAX_AGGREGATE_ITEMS || '500', 10);

      const result = await searchStrategy.execute(personName, {
        entityTypes: ['person', 'message', 'event', 'driveItem'],
        maxResults,
      });

      // Aggregate results
      const aggregated = dataAggregator.aggregate(
        [
          {
            source: 'search',
            items: result.searchResults.items,
          },
          ...Object.entries(result.specificResults).map(([source, items]) => ({
            source,
            items: items as unknown[],
          })),
        ],
        {
          sortBy: 'relevance',
          sortOrder: 'desc',
          maxItems: maxAggregateItems,
          deduplicate: true,
          formatForLLM: true,
        }
      );

      // Add download links for files
      let downloadLinks: Array<{ fileName: string; downloadUrl: string }> | undefined;
      if (includeFiles && downloadLinkGenerator) {
        const links = await downloadLinkGenerator.addDownloadLinksToResults(
          aggregated.items.map((i) => i.data)
        );
        downloadLinks = links
          .filter((l) => l.downloadLink)
          .map((l) => ({
            fileName: l.downloadLink!.fileName,
            downloadUrl: l.downloadLink!.downloadUrl,
          }));
      }

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                person: personName,
                totalItems: result.totalItems,
                sources: aggregated.sources,
                items: aggregated.items,
                downloadLinks,
                formattedSummary: aggregated.formattedForLLM,
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );

  // Discover Meeting Tool
  server.tool(
    'discover-meeting',
    'Discover comprehensive meeting information: participants, agenda, notes, recordings, related files, and follow-up actions.',
    {
      meetingQuery: z.string().describe('Meeting name, topic, or date'),
      includeRecordings: z.boolean().optional().describe('Include meeting recordings'),
      includeNotes: z.boolean().optional().describe('Include OneNote meeting notes'),
      includeFiles: z.boolean().optional().describe('Include related files'),
    },
    async ({ meetingQuery, includeRecordings, includeNotes, includeFiles }) => {
      if (!searchStrategy || !dataAggregator) {
        throw new Error('Discovery components not initialized');
      }

      logger.info(`Discovering meeting: "${meetingQuery}"`);

      const maxResults = parseInt(process.env.MS365_MCP_MAX_RESULTS || '500', 10);
      const maxAggregateItems = parseInt(process.env.MS365_MCP_MAX_AGGREGATE_ITEMS || '500', 10);

      const result = await searchStrategy.execute(meetingQuery, {
        entityTypes: ['event', 'message', 'driveItem'],
        maxResults,
      });

      // Aggregate results
      const aggregated = dataAggregator.aggregate(
        [
          {
            source: 'search',
            items: result.searchResults.items,
          },
          ...Object.entries(result.specificResults).map(([source, items]) => ({
            source,
            items: items as unknown[],
          })),
        ],
        {
          sortBy: 'timestamp',
          sortOrder: 'desc',
          maxItems: maxAggregateItems,
          deduplicate: true,
          formatForLLM: true,
        }
      );

      // Add download links for files
      let downloadLinks: Array<{ fileName: string; downloadUrl: string }> | undefined;
      if (includeFiles && downloadLinkGenerator) {
        const links = await downloadLinkGenerator.addDownloadLinksToResults(
          aggregated.items.map((i) => i.data)
        );
        downloadLinks = links
          .filter((l) => l.downloadLink)
          .map((l) => ({
            fileName: l.downloadLink!.fileName,
            downloadUrl: l.downloadLink!.downloadUrl,
          }));
      }

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                meeting: meetingQuery,
                totalItems: result.totalItems,
                sources: aggregated.sources,
                items: aggregated.items,
                downloadLinks,
                formattedSummary: aggregated.formattedForLLM,
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );

  // Discover Document Tool
  server.tool(
    'discover-document',
    'Discover a document across all sources: location, versions, authors, sharing, related emails, and discussions.',
    {
      documentName: z.string().describe('Document name or keywords'),
      includeVersions: z.boolean().optional().describe('Include version history'),
      includeSharing: z.boolean().optional().describe('Include sharing information'),
      includeDownloadLink: z.boolean().optional().describe('Include download link'),
    },
    async ({ documentName, includeVersions, includeSharing, includeDownloadLink }) => {
      if (!searchStrategy || !dataAggregator) {
        throw new Error('Discovery components not initialized');
      }

      logger.info(`Discovering document: "${documentName}"`);

      const maxResults = parseInt(process.env.MS365_MCP_MAX_RESULTS || '500', 10);
      const maxAggregateItems = parseInt(process.env.MS365_MCP_MAX_AGGREGATE_ITEMS || '500', 10);

      const result = await searchStrategy.execute(documentName, {
        entityTypes: ['driveItem'],
        maxResults,
      });

      // Aggregate results
      const aggregated = dataAggregator.aggregate(
        [
          {
            source: 'search',
            items: result.searchResults.items,
          },
          ...Object.entries(result.specificResults).map(([source, items]) => ({
            source,
            items: items as unknown[],
          })),
        ],
        {
          sortBy: 'relevance',
          sortOrder: 'desc',
          maxItems: maxAggregateItems,
          deduplicate: true,
          formatForLLM: true,
        }
      );

      // Add download links if requested (default true for documents)
      let downloadLinks: Array<{ fileName: string; downloadUrl: string }> | undefined;
      if (includeDownloadLink !== false && downloadLinkGenerator) {
        const links = await downloadLinkGenerator.addDownloadLinksToResults(
          aggregated.items.map((i) => i.data)
        );
        downloadLinks = links
          .filter((l) => l.downloadLink)
          .map((l) => ({
            fileName: l.downloadLink!.fileName,
            downloadUrl: l.downloadLink!.downloadUrl,
          }));
      }

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                document: documentName,
                totalItems: result.totalItems,
                sources: aggregated.sources,
                items: aggregated.items,
                downloadLinks,
                formattedSummary: aggregated.formattedForLLM,
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );

  // Discover Team Tool
  server.tool(
    'discover-team',
    'Discover comprehensive team information: members, channels, projects, files, meetings, and recent activity.',
    {
      teamName: z.string().describe('Team name or identifier'),
      includeActivity: z.boolean().optional().describe('Include recent activity'),
      timeRange: z.string().optional().describe('Time range for activity (e.g., "7d", "30d")'),
    },
    async ({ teamName, includeActivity, timeRange }) => {
      if (!searchStrategy || !dataAggregator) {
        throw new Error('Discovery components not initialized');
      }

      logger.info(`Discovering team: "${teamName}"`);

      const maxResults = parseInt(process.env.MS365_MCP_MAX_RESULTS || '500', 10);
      const maxAggregateItems = parseInt(process.env.MS365_MCP_MAX_AGGREGATE_ITEMS || '500', 10);

      const result = await searchStrategy.execute(teamName, {
        entityTypes: ['chatMessage', 'driveItem', 'event'],
        maxResults,
        timeRange,
      });

      // Aggregate results
      const aggregated = dataAggregator.aggregate(
        [
          {
            source: 'search',
            items: result.searchResults.items,
          },
          ...Object.entries(result.specificResults).map(([source, items]) => ({
            source,
            items: items as unknown[],
          })),
        ],
        {
          sortBy: 'timestamp',
          sortOrder: 'desc',
          maxItems: maxAggregateItems,
          deduplicate: true,
          formatForLLM: true,
        }
      );

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                team: teamName,
                totalItems: result.totalItems,
                sources: aggregated.sources,
                items: aggregated.items,
                formattedSummary: aggregated.formattedForLLM,
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );

  // Discover Customer Tool
  server.tool(
    'discover-customer',
    'Discover all information about a customer: contact details, communication history, meetings, files, and notes.',
    {
      customerName: z.string().describe('Customer name or company'),
      includeHistory: z.boolean().optional().describe('Include communication history'),
      timeRange: z.string().optional().describe('Time range for history'),
    },
    async ({ customerName, includeHistory, timeRange }) => {
      if (!searchStrategy || !dataAggregator) {
        throw new Error('Discovery components not initialized');
      }

      logger.info(`Discovering customer: "${customerName}"`);

      const maxResults = parseInt(process.env.MS365_MCP_MAX_RESULTS || '500', 10);
      const maxAggregateItems = parseInt(process.env.MS365_MCP_MAX_AGGREGATE_ITEMS || '500', 10);

      const result = await searchStrategy.execute(customerName, {
        entityTypes: ['person', 'message', 'event', 'driveItem'],
        maxResults,
        timeRange,
      });

      // Aggregate results
      const aggregated = dataAggregator.aggregate(
        [
          {
            source: 'search',
            items: result.searchResults.items,
          },
          ...Object.entries(result.specificResults).map(([source, items]) => ({
            source,
            items: items as unknown[],
          })),
        ],
        {
          sortBy: 'timestamp',
          sortOrder: 'desc',
          maxItems: maxAggregateItems,
          deduplicate: true,
          formatForLLM: true,
        }
      );

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                customer: customerName,
                totalItems: result.totalItems,
                sources: aggregated.sources,
                items: aggregated.items,
                formattedSummary: aggregated.formattedForLLM,
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );

  // Discover Contract Tool
  server.tool(
    'discover-contract',
    'Discover contract information: location, parties, dates, related communications, and renewal information.',
    {
      contractQuery: z.string().describe('Contract name, number, or party'),
      includeRenewals: z.boolean().optional().describe('Include renewal information'),
      includeDownloadLink: z
        .boolean()
        .optional()
        .describe('Include download link for contract file'),
    },
    async ({ contractQuery, includeRenewals, includeDownloadLink }) => {
      if (!searchStrategy || !dataAggregator) {
        throw new Error('Discovery components not initialized');
      }

      logger.info(`Discovering contract: "${contractQuery}"`);

      const maxResults = parseInt(process.env.MS365_MCP_MAX_RESULTS || '500', 10);
      const maxAggregateItems = parseInt(process.env.MS365_MCP_MAX_AGGREGATE_ITEMS || '500', 10);

      const result = await searchStrategy.execute(contractQuery, {
        entityTypes: ['driveItem', 'message'],
        maxResults,
      });

      // Aggregate results
      const aggregated = dataAggregator.aggregate(
        [
          {
            source: 'search',
            items: result.searchResults.items,
          },
          ...Object.entries(result.specificResults).map(([source, items]) => ({
            source,
            items: items as unknown[],
          })),
        ],
        {
          sortBy: 'relevance',
          sortOrder: 'desc',
          maxItems: maxAggregateItems,
          deduplicate: true,
          formatForLLM: true,
        }
      );

      // Add download links if requested (default true for contracts)
      let downloadLinks: Array<{ fileName: string; downloadUrl: string }> | undefined;
      if (includeDownloadLink !== false && downloadLinkGenerator) {
        const links = await downloadLinkGenerator.addDownloadLinksToResults(
          aggregated.items.map((i) => i.data)
        );
        downloadLinks = links
          .filter((l) => l.downloadLink)
          .map((l) => ({
            fileName: l.downloadLink!.fileName,
            downloadUrl: l.downloadLink!.downloadUrl,
          }));
      }

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                contract: contractQuery,
                totalItems: result.totalItems,
                sources: aggregated.sources,
                items: aggregated.items,
                downloadLinks,
                formattedSummary: aggregated.formattedForLLM,
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );

  // Discover Decision Tool
  server.tool(
    'discover-decision',
    'Discover decisions made on a topic: who decided, when, rationale, related discussions, and implementation.',
    {
      topic: z.string().describe('Topic or decision keyword'),
      includeRationale: z.boolean().optional().describe('Include decision rationale'),
    },
    async ({ topic, includeRationale }) => {
      if (!searchStrategy || !dataAggregator) {
        throw new Error('Discovery components not initialized');
      }

      logger.info(`Discovering decision: "${topic}"`);

      const maxResults = parseInt(process.env.MS365_MCP_MAX_RESULTS || '500', 10);
      const maxAggregateItems = parseInt(process.env.MS365_MCP_MAX_AGGREGATE_ITEMS || '500', 10);

      const result = await searchStrategy.execute(topic, {
        entityTypes: ['message', 'chatMessage', 'driveItem'],
        maxResults,
      });

      // Aggregate results
      const aggregated = dataAggregator.aggregate(
        [
          {
            source: 'search',
            items: result.searchResults.items,
          },
          ...Object.entries(result.specificResults).map(([source, items]) => ({
            source,
            items: items as unknown[],
          })),
        ],
        {
          sortBy: 'timestamp',
          sortOrder: 'desc',
          maxItems: maxAggregateItems,
          deduplicate: true,
          formatForLLM: true,
        }
      );

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                topic,
                totalItems: result.totalItems,
                sources: aggregated.sources,
                items: aggregated.items,
                formattedSummary: aggregated.formattedForLLM,
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );

  logger.info('Discovery tools registered successfully');
}
