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
import LearningDashboard from './learning-dashboard.js';
import NLPEnhancer from './nlp-enhancer.js';
import ToolCombiner from './tool-combiner.js';
import DataAggregator from './data-aggregator.js';
import DownloadLinkGenerator from './download-link-generator.js';
import DeepResearchEngine from './deep-research-engine.js';
import type { AppSecrets } from './secrets.js';
import { getMaxResults, getMaxAggregateItems } from './perf-config.js';

let searchStrategy: SearchFirstStrategy | null = null;
let deepResearchEngine: DeepResearchEngine | null = null;
let dataAggregator: DataAggregator | null = null;
let downloadLinkGenerator: DownloadLinkGenerator | null = null;
let learningDashboard: LearningDashboard | null = null;

// Store references for status checks
let knowledgeBaseInstance: KnowledgeBase | null = null;
let learningSystemInstance: LearningSystem | null = null;

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

  // Store references for status checks
  knowledgeBaseInstance = knowledgeBase;
  learningSystemInstance = learningSystem;

  // Log Learning System initialization status
  const learningEnabled =
    process.env.MS365_MCP_LEARNING_ENABLED !== 'false' &&
    process.env.MS365_MCP_LEARNING_ENABLED !== '0';
  const clusterEnabled =
    process.env.MS365_MCP_LEARNING_CLUSTER_ENABLED === 'true' ||
    process.env.MS365_MCP_LEARNING_CLUSTER_ENABLED !== 'false';
  const nlpEnabled =
    process.env.MS365_MCP_LEARNING_NLP_ENABLED === 'true' ||
    process.env.MS365_MCP_LEARNING_NLP_ENABLED !== 'false';

  logger.info('Learning System initialized', {
    learningEnabled,
    clusterEnabled,
    nlpEnabled,
    knowledgeBasePath: process.env.MS365_MCP_KNOWLEDGE_BASE_PATH || './data/knowledge-base.json',
    decayDays: process.env.MS365_MCP_LEARNING_DECAY_DAYS || '90',
    decayFactor: process.env.MS365_MCP_LEARNING_DECAY_FACTOR || '0.1',
  });

  // Log knowledge base status
  const kbData = knowledgeBase.getAllData();
  const totalPatterns = Object.keys(kbData.queryPatterns).length;
  const totalQueries = Object.keys(kbData.successfulQueries).length;
  const totalSynonyms = Object.keys(kbData.learnedSynonyms).length;
  const totalFeedback = Object.values(kbData.userFeedback).reduce(
    (sum, arr) => sum + arr.length,
    0
  );

  logger.info('Knowledge Base status', {
    totalPatterns,
    totalQueries,
    totalSynonyms,
    totalFeedback,
    lastUpdated: kbData.lastUpdated,
    version: kbData.version,
  });

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

  // Initialize learning dashboard
  const nlpEnhancer = new NLPEnhancer();
  learningDashboard = new LearningDashboard(learningSystem, knowledgeBase, nlpEnhancer);
}

/**
 * Initialize only the Learning System (and dependencies) for use without Discovery Tools.
 * Call when Super-Tools or Classic mode is used and MS365_MCP_LEARNING_ENABLED is true.
 * Idempotent: if already initialized, does nothing.
 */
export function ensureLearningSystemInitialized(): void {
  if (learningSystemInstance !== null) {
    return;
  }
  const knowledgeBase = new KnowledgeBase();
  const synonymExpander = new SynonymExpander();
  const learningSystem = new LearningSystem(knowledgeBase, synonymExpander);
  knowledgeBaseInstance = knowledgeBase;
  learningSystemInstance = learningSystem;
  logger.info(
    'Learning System initialized (standalone, without Discovery Tools). Pattern and entity-type learning will be used by Search/Assistant.'
  );
}

/**
 * Return the Learning System instance if it has been initialized (Discovery or standalone).
 * Use from Super-Tools search handler to record learnFromSearch and getRecommendedEntityTypes.
 */
export function getLearningSystem(): LearningSystem | null {
  return learningSystemInstance;
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

      const configuredMaxResults = getMaxResults();
      const maxAggregateItems = getMaxAggregateItems();
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

      const maxResults = getMaxResults();
      const maxAggregateItems = getMaxAggregateItems();

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

      const maxResults = getMaxResults();
      const maxAggregateItems = getMaxAggregateItems();

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

      const maxResults = getMaxResults();
      const maxAggregateItems = getMaxAggregateItems();

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

      const maxResults = getMaxResults();
      const maxAggregateItems = getMaxAggregateItems();

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

      const maxResults = getMaxResults();
      const maxAggregateItems = getMaxAggregateItems();

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

      const maxResults = getMaxResults();
      const maxAggregateItems = getMaxAggregateItems();

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

      const maxResults = getMaxResults();
      const maxAggregateItems = getMaxAggregateItems();

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

      const maxResults = getMaxResults();
      const maxAggregateItems = getMaxAggregateItems();

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

  // ============================================================================
  // USER FEEDBACK TOOL - For explicit user feedback on search results
  // ============================================================================

  server.tool(
    'provide-feedback',
    `Provide explicit feedback on search results to help the learning system improve.

This tool allows users to give feedback on search results, which helps the system learn
which queries and patterns work well. User feedback is weighted higher than implicit
learning from search results.

Feedback types:
- helpful: The search results were helpful and relevant
- not_helpful: The search results were not helpful or relevant
- correct: The search results were correct and accurate
- incorrect: The search results were incorrect or inaccurate

The learning system will use this feedback to:
- Improve query pattern confidence scores
- Adjust entity type recommendations
- Enhance synonym learning
- Refine data location mappings`,
    {
      query: z.string().describe('The original search query that was used').optional(),
      resultId: z
        .string()
        .describe('Optional identifier for the specific result being feedback on')
        .optional(),
      feedbackType: z
        .enum(['helpful', 'not_helpful', 'correct', 'incorrect'])
        .describe('Type of feedback: helpful, not_helpful, correct, or incorrect'),
      comment: z.string().describe('Optional comment explaining the feedback').optional(),
      context: z
        .string()
        .describe('Optional context about the search (e.g., entity types, sources)')
        .optional(),
    },
    {
      title: 'provide-feedback',
      readOnlyHint: false,
      openWorldHint: false,
    },
    async ({ query, resultId, feedbackType, comment, context }) => {
      if (!searchStrategy) {
        throw new Error('Discovery components not initialized');
      }

      // Get learning system from search strategy
      const learningSystem = (searchStrategy as any).learningSystem as LearningSystem;
      if (!learningSystem) {
        throw new Error('Learning system not available');
      }

      if (!query) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                {
                  success: false,
                  error: 'Query is required for feedback',
                },
                null,
                2
              ),
            },
          ],
          isError: true,
        };
      }

      try {
        await learningSystem.recordUserFeedback(query, feedbackType, resultId, comment, context);

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                {
                  success: true,
                  message: `Feedback recorded: ${feedbackType} for query "${query}"`,
                  feedbackType,
                  query,
                  resultId,
                  comment,
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (error) {
        logger.error(`Failed to record user feedback: ${error}`);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                {
                  success: false,
                  error: `Failed to record feedback: ${error}`,
                },
                null,
                2
              ),
            },
          ],
          isError: true,
        };
      }
    }
  );

  // ============================================================================
  // EXPORT/IMPORT TOOLS - For knowledge base management
  // ============================================================================

  server.tool(
    'export-knowledge-base',
    `Export the knowledge base as JSON for backup or sharing.

The knowledge base contains all learned patterns, synonyms, entity mappings, and
user feedback. This tool allows you to export it for backup purposes or to share
learned patterns with other instances.`,
    {},
    {
      title: 'export-knowledge-base',
      readOnlyHint: true,
      openWorldHint: false,
    },
    async () => {
      if (!searchStrategy) {
        throw new Error('Discovery components not initialized');
      }

      const knowledgeBase = (searchStrategy as any).knowledgeBase as KnowledgeBase;
      if (!knowledgeBase) {
        throw new Error('Knowledge base not available');
      }

      try {
        const exported = knowledgeBase.exportKnowledgeBase();
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                {
                  success: true,
                  message: 'Knowledge base exported successfully',
                  data: JSON.parse(exported),
                  size: exported.length,
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (error) {
        logger.error(`Failed to export knowledge base: ${error}`);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                {
                  success: false,
                  error: `Failed to export knowledge base: ${error}`,
                },
                null,
                2
              ),
            },
          ],
          isError: true,
        };
      }
    }
  );

  server.tool(
    'import-knowledge-base',
    `Import a knowledge base from JSON.

This tool allows you to import a previously exported knowledge base. The imported
data will be merged with the existing knowledge base, preserving existing patterns
while adding new ones.`,
    {
      data: z.string().describe('JSON string of the knowledge base to import'),
      merge: z
        .boolean()
        .describe('If true, merge with existing knowledge base. If false, replace it.')
        .optional()
        .default(true),
    },
    {
      title: 'import-knowledge-base',
      readOnlyHint: false,
      openWorldHint: false,
    },
    async ({ data, merge = true }) => {
      if (!searchStrategy) {
        throw new Error('Discovery components not initialized');
      }

      const knowledgeBase = (searchStrategy as any).knowledgeBase as KnowledgeBase;
      if (!knowledgeBase) {
        throw new Error('Knowledge base not available');
      }

      try {
        if (merge) {
          // Parse and merge
          const imported = JSON.parse(data);
          knowledgeBase.mergeKnowledgeBase(imported);
        } else {
          // Replace
          knowledgeBase.importKnowledgeBase(data);
        }

        // Save after import
        await knowledgeBase.save();

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                {
                  success: true,
                  message: `Knowledge base ${merge ? 'merged' : 'imported'} successfully`,
                  merged: merge,
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (error) {
        logger.error(`Failed to import knowledge base: ${error}`);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                {
                  success: false,
                  error: `Failed to import knowledge base: ${error}`,
                },
                null,
                2
              ),
            },
          ],
          isError: true,
        };
      }
    }
  );

  // ============================================================================
  // LEARNING STATUS TOOL - Check if Learning System is active and working
  // ============================================================================

  server.tool(
    'learning-status',
    `Check the status of the MCP Learning System.

This tool provides a comprehensive status check of the learning system including:
- Whether learning is enabled (MS365_MCP_LEARNING_ENABLED)
- Knowledge base path and accessibility
- Current learning statistics
- Pattern clustering status
- NLP enhancement status

Use this tool to verify the learning system is active and functioning correctly.`,
    {},
    {
      title: 'learning-status',
      readOnlyHint: true,
      openWorldHint: false,
    },
    async () => {
      // Check environment configuration
      const learningEnabled =
        process.env.MS365_MCP_LEARNING_ENABLED !== 'false' &&
        process.env.MS365_MCP_LEARNING_ENABLED !== '0';
      const clusterEnabled =
        process.env.MS365_MCP_LEARNING_CLUSTER_ENABLED === 'true' ||
        process.env.MS365_MCP_LEARNING_CLUSTER_ENABLED !== 'false';
      const nlpEnabled =
        process.env.MS365_MCP_LEARNING_NLP_ENABLED === 'true' ||
        process.env.MS365_MCP_LEARNING_NLP_ENABLED !== 'false';

      // Get knowledge base data
      let kbStatus = {
        accessible: false,
        totalPatterns: 0,
        totalQueries: 0,
        totalSynonyms: 0,
        totalEntityMappings: 0,
        totalDataLocations: 0,
        totalUserFeedback: 0,
        totalToolUsagePatterns: 0,
        totalConfidenceScores: 0,
        lastUpdated: 'unknown',
        version: 0,
      };

      if (knowledgeBaseInstance) {
        try {
          const kbData = knowledgeBaseInstance.getAllData();
          kbStatus = {
            accessible: true,
            totalPatterns: Object.keys(kbData.queryPatterns).length,
            totalQueries: Object.keys(kbData.successfulQueries).length,
            totalSynonyms: Object.keys(kbData.learnedSynonyms).length,
            totalEntityMappings: Object.keys(kbData.entityMappings).length,
            totalDataLocations: Object.keys(kbData.dataLocations).length,
            totalUserFeedback: Object.values(kbData.userFeedback).reduce(
              (sum, arr) => sum + arr.length,
              0
            ),
            totalToolUsagePatterns: Object.keys(kbData.toolUsagePatterns).length,
            totalConfidenceScores: Object.keys(kbData.confidenceScores).length,
            lastUpdated: kbData.lastUpdated,
            version: kbData.version,
          };
        } catch (error) {
          logger.warn(`Failed to read knowledge base: ${error}`);
        }
      }

      // Check learning system availability
      const learningSystemAvailable = !!learningSystemInstance;

      // Get performance metrics if available
      let performanceMetrics = null;
      if (learningDashboard) {
        try {
          performanceMetrics = learningDashboard.getLearningStats();
        } catch (error) {
          logger.warn(`Failed to get performance metrics: ${error}`);
        }
      }

      const status = {
        learningSystem: {
          enabled: learningEnabled,
          environmentVariable:
            process.env.MS365_MCP_LEARNING_ENABLED || '(not set - defaults to true)',
          available: learningSystemAvailable,
          clusteringEnabled: clusterEnabled,
          nlpEnabled: nlpEnabled,
        },
        configuration: {
          knowledgeBasePath:
            process.env.MS365_MCP_KNOWLEDGE_BASE_PATH || './data/knowledge-base.json',
          decayDays: parseInt(process.env.MS365_MCP_LEARNING_DECAY_DAYS || '90', 10),
          decayFactor: parseFloat(process.env.MS365_MCP_LEARNING_DECAY_FACTOR || '0.1'),
          clusterThreshold: parseFloat(process.env.MS365_MCP_LEARNING_CLUSTER_THRESHOLD || '0.7'),
        },
        knowledgeBase: kbStatus,
        performanceMetrics: performanceMetrics
          ? {
              totalQueries: performanceMetrics.totalQueries,
              successfulQueries: performanceMetrics.successfulQueries,
              failedQueries: performanceMetrics.failedQueries,
              successRate: performanceMetrics.successRate,
              averageResultsPerQuery: performanceMetrics.averageResultsPerQuery,
              averageConfidence: performanceMetrics.averageConfidence,
            }
          : null,
        recommendations: [] as string[],
      };

      // Add recommendations based on status
      if (!learningEnabled) {
        status.recommendations.push(
          'Learning is disabled. Set MS365_MCP_LEARNING_ENABLED=true to enable.'
        );
      }
      if (kbStatus.totalPatterns === 0 && kbStatus.totalQueries === 0) {
        status.recommendations.push(
          'Knowledge base is empty. Use ms365-search or deep-research tools to start learning.'
        );
      }
      if (kbStatus.totalUserFeedback === 0) {
        status.recommendations.push(
          'No user feedback recorded. Use provide-feedback tool to improve learning quality.'
        );
      }
      if (!clusterEnabled) {
        status.recommendations.push(
          'Pattern clustering is disabled. Set MS365_MCP_LEARNING_CLUSTER_ENABLED=true for better pattern grouping.'
        );
      }

      logger.info('Learning status check completed', {
        enabled: learningEnabled,
        available: learningSystemAvailable,
        kbAccessible: kbStatus.accessible,
        totalPatterns: kbStatus.totalPatterns,
      });

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(status, null, 2),
          },
        ],
      };
    }
  );

  // ============================================================================
  // LEARNING DASHBOARD TOOL - For learning insights and statistics
  // ============================================================================

  server.tool(
    'get-learning-insights',
    `Get insights and statistics about the learning system.

This tool provides comprehensive analytics about the learning system including:
- Overall performance metrics (success rate, average results, confidence scores)
- Top performing query patterns
- Confidence score distribution
- Tool usage statistics
- Performance trends over time
- Pattern clusters

Use this to understand how well the learning system is performing and which
patterns are most effective.`,
    {
      days: z
        .number()
        .describe('Number of days for performance trends (default: 30)')
        .optional()
        .default(30),
      includeClusters: z
        .boolean()
        .describe('Include pattern cluster information (default: true)')
        .optional()
        .default(true),
    },
    {
      title: 'get-learning-insights',
      readOnlyHint: true,
      openWorldHint: false,
    },
    async ({ days = 30, includeClusters = true }) => {
      if (!learningDashboard) {
        throw new Error('Learning dashboard not initialized');
      }

      try {
        const insights = learningDashboard.getLearningInsights();

        // Optionally exclude clusters
        if (!includeClusters) {
          delete insights.clusters;
        }

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(insights, null, 2),
            },
          ],
        };
      } catch (error) {
        logger.error(`Failed to get learning insights: ${error}`);
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify(
                {
                  success: false,
                  error: `Failed to get learning insights: ${error}`,
                },
                null,
                2
              ),
            },
          ],
          isError: true,
        };
      }
    }
  );

  logger.info('Discovery tools registered successfully');
}
