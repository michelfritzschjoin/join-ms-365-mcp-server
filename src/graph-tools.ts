import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import logger from './logger.js';
import GraphClient from './graph-client.js';
import { api } from './generated/client.js';
import { z } from 'zod';
import { readFileSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { TOOL_CATEGORIES } from './tool-categories.js';
import type KnowledgeBase from './knowledge-base.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

interface EndpointConfig {
  pathPattern: string;
  method: string;
  toolName: string;
  scopes?: string[];
  workScopes?: string[];
  returnDownloadUrl?: boolean;
  supportsTimezone?: boolean;
  llmTip?: string;
}

const endpointsData = JSON.parse(
  readFileSync(path.join(__dirname, 'endpoints.json'), 'utf8')
) as EndpointConfig[];

type TextContent = {
  type: 'text';
  text: string;
  [key: string]: unknown;
};

type ImageContent = {
  type: 'image';
  data: string;
  mimeType: string;
  [key: string]: unknown;
};

type AudioContent = {
  type: 'audio';
  data: string;
  mimeType: string;
  [key: string]: unknown;
};

type ResourceTextContent = {
  type: 'resource';
  resource: {
    text: string;
    uri: string;
    mimeType?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
};

type ResourceBlobContent = {
  type: 'resource';
  resource: {
    blob: string;
    uri: string;
    mimeType?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
};

type ResourceContent = ResourceTextContent | ResourceBlobContent;

type ContentItem = TextContent | ImageContent | AudioContent | ResourceContent;

interface CallToolResult {
  content: ContentItem[];
  _meta?: Record<string, unknown>;
  isError?: boolean;

  [key: string]: unknown;
}

// Global tool usage tracker (session-based)
let toolUsageTracker: {
  knowledgeBase: KnowledgeBase | null;
  currentSession: string[];
  sessionStartTime: number;
} = {
  knowledgeBase: null,
  currentSession: [],
  sessionStartTime: Date.now(),
};

// Reset session after 5 minutes of inactivity
const SESSION_TIMEOUT = 5 * 60 * 1000; // 5 minutes

function resetSessionIfNeeded(): void {
  const now = Date.now();
  if (now - toolUsageTracker.sessionStartTime > SESSION_TIMEOUT) {
    toolUsageTracker.currentSession = [];
    toolUsageTracker.sessionStartTime = now;
  }
}

async function executeGraphTool(
  tool: (typeof api.endpoints)[0],
  config: EndpointConfig | undefined,
  graphClient: GraphClient,
  params: Record<string, unknown>
): Promise<CallToolResult> {
  logger.info(`Tool ${tool.alias} called with params: ${JSON.stringify(params)}`);

  resetSessionIfNeeded();
  const startTime = Date.now();

  try {
    const parameterDefinitions = tool.parameters || [];

    // Apply default $select for detailed content - no date filter by default
    // Date filters are only applied when user explicitly specifies a time range
    const isCalendarTool = tool.path.includes('/events') || tool.path.includes('calendar');

    // Apply default $top for calendar events to get more results (Microsoft Graph defaults to only 10)
    if (isCalendarTool && !params['$top'] && !params['top']) {
      params['$top'] = 100;
      logger.info('Applied default $top=100 for calendar events (MS Graph default is only 10)');
    }

    // For calendarView, add default date range if not provided (required parameters)
    if (tool.path.includes('/calendarView') || tool.alias === 'get-calendar-view') {
      if (!params['startDateTime']) {
        // Default to 30 days in the past
        const startDate = new Date();
        startDate.setDate(startDate.getDate() - 30);
        params['startDateTime'] = startDate.toISOString();
        logger.info('Applied default startDateTime (30 days ago) for calendarView');
      }
      if (!params['endDateTime']) {
        // Default to 90 days in the future
        const endDate = new Date();
        endDate.setDate(endDate.getDate() + 90);
        params['endDateTime'] = endDate.toISOString();
        logger.info('Applied default endDateTime (90 days future) for calendarView');
      }
    }

    if (!params['$select']) {
      // Calendar events - return detailed content
      if (isCalendarTool) {
        params['$select'] = [
          'id',
          'subject',
          'bodyPreview',
          'body',
          'start',
          'end',
          'location',
          'attendees',
          'organizer',
          'isOnlineMeeting',
          'onlineMeeting',
          'webLink',
          'isAllDay',
          'isCancelled',
          'importance',
          'sensitivity',
          'showAs',
          'responseStatus',
          'categories',
        ];
        logger.info('Applied default $select for detailed calendar content');
      }
      // Mail messages - return detailed content
      else if (tool.path.includes('/messages') || tool.path.includes('/mail')) {
        params['$select'] = [
          'id',
          'subject',
          'bodyPreview',
          'body',
          'from',
          'toRecipients',
          'ccRecipients',
          'receivedDateTime',
          'sentDateTime',
          'hasAttachments',
          'importance',
          'isRead',
          'isDraft',
          'webLink',
          'categories',
          'flag',
        ];
        logger.info('Applied default $select for detailed mail content');
      }
      // Files/Drive items - return detailed content
      else if (tool.path.includes('/drive') || tool.path.includes('/items')) {
        params['$select'] = [
          'id',
          'name',
          'size',
          'createdDateTime',
          'lastModifiedDateTime',
          'webUrl',
          'createdBy',
          'lastModifiedBy',
          'file',
          'folder',
          'parentReference',
        ];
        logger.info('Applied default $select for detailed file content');
      }
      // Tasks - return detailed content
      else if (tool.path.includes('/tasks') || tool.path.includes('/todo')) {
        params['$select'] = [
          'id',
          'title',
          'body',
          'importance',
          'status',
          'createdDateTime',
          'lastModifiedDateTime',
          'dueDateTime',
          'completedDateTime',
          'reminderDateTime',
          'categories',
        ];
        logger.info('Applied default $select for detailed task content');
      }
      // Contacts - return detailed content
      else if (tool.path.includes('/contacts')) {
        params['$select'] = [
          'id',
          'displayName',
          'givenName',
          'surname',
          'emailAddresses',
          'businessPhones',
          'mobilePhone',
          'companyName',
          'jobTitle',
          'department',
          'officeLocation',
        ];
        logger.info('Applied default $select for detailed contact content');
      }
      // Users - return detailed content
      else if (tool.path.includes('/users')) {
        params['$select'] = [
          'id',
          'displayName',
          'givenName',
          'surname',
          'mail',
          'userPrincipalName',
          'jobTitle',
          'department',
          'officeLocation',
          'mobilePhone',
          'businessPhones',
        ];
        logger.info('Applied default $select for detailed user content');
      }
    }

    let path = tool.path;
    const queryParams: Record<string, string> = {};
    const headers: Record<string, string> = {};
    let body: unknown = null;

    for (const [paramName, paramValue] of Object.entries(params)) {
      // Skip control parameters - not part of the Microsoft Graph API
      if (['fetchAllPages', 'includeHeaders', 'excludeResponse', 'timezone'].includes(paramName)) {
        continue;
      }

      // Ok, so, MCP clients (such as claude code) doesn't support $ in parameter names,
      // and others might not support __, so we strip them in hack.ts and restore them here
      const odataParams = [
        'filter',
        'select',
        'expand',
        'orderby',
        'skip',
        'top',
        'count',
        'search',
        'format',
      ];
      // Handle both "top" and "$top" formats - strip $ if present, then re-add it
      const normalizedParamName = paramName.startsWith('$') ? paramName.slice(1) : paramName;
      const isOdataParam = odataParams.includes(normalizedParamName.toLowerCase());
      const fixedParamName = isOdataParam ? `$${normalizedParamName.toLowerCase()}` : paramName;
      // Look up param definition using normalized name (without $) for OData params
      const paramDef = parameterDefinitions.find(
        (p) => p.name === paramName || (isOdataParam && p.name === normalizedParamName)
      );

      if (paramDef) {
        switch (paramDef.type) {
          case 'Path':
            path = path
              .replace(`{${paramName}}`, encodeURIComponent(paramValue as string))
              .replace(`:${paramName}`, encodeURIComponent(paramValue as string));
            break;

          case 'Query':
            queryParams[fixedParamName] = `${paramValue}`;
            break;

          case 'Body':
            if (paramDef.schema) {
              const parseResult = paramDef.schema.safeParse(paramValue);
              if (!parseResult.success) {
                const wrapped = { [paramName]: paramValue };
                const wrappedResult = paramDef.schema.safeParse(wrapped);
                if (wrappedResult.success) {
                  logger.info(
                    `Auto-corrected parameter '${paramName}': AI passed nested field directly, wrapped it as {${paramName}: ...}`
                  );
                  body = wrapped;
                } else {
                  body = paramValue;
                }
              } else {
                body = paramValue;
              }
            } else {
              body = paramValue;
            }
            break;

          case 'Header':
            headers[fixedParamName] = `${paramValue}`;
            break;
        }
      } else if (paramName === 'body') {
        body = paramValue;
        logger.info(`Set body param: ${JSON.stringify(body)}`);
      }
    }

    // Handle search parameter for directory endpoints that require special formatting
    // Microsoft Graph API requirements:
    // - /users: requires property:value format (e.g., "displayName:John") + ConsistencyLevel: eventual
    // - /groups: requires property:value format (e.g., "displayName:Team") + ConsistencyLevel: eventual
    // - /sites: requires ConsistencyLevel: eventual (but supports free-text search)

    // Endpoints that require ConsistencyLevel: eventual header for $search
    const requiresConsistencyLevelHeader = ['list-users', 'search-sharepoint-sites'];

    // Endpoints that require property:value format for $search
    const requiresPropertyValueFormat = ['list-users'];

    if (queryParams['$search']) {
      let searchValue = queryParams['$search'];

      // Remove surrounding quotes if present (handles both single and double quotes)
      const quotePattern = /^(["'])(.*)\1$/;
      const match = searchValue.match(quotePattern);
      const cleanSearchValue = match ? match[2] : searchValue;

      // Check if this endpoint requires property:value format
      if (requiresPropertyValueFormat.includes(tool.alias)) {
        // Check if search value is already in property:value format
        const propertyValuePattern = /^[a-zA-Z]+:/i;

        if (!propertyValuePattern.test(cleanSearchValue)) {
          // Auto-format: prepend displayName: if not already formatted
          queryParams['$search'] = `"displayName:${cleanSearchValue}"`;
          logger.info(
            `Auto-formatted search query for ${tool.alias}: "${searchValue}" -> "${queryParams['$search']}"`
          );
        } else {
          // Already in property:value format, ensure it's wrapped in double quotes
          queryParams['$search'] = `"${cleanSearchValue}"`;
          logger.info(
            `Search query already formatted, ensuring quotes: "${searchValue}" -> "${queryParams['$search']}"`
          );
        }
      } else {
        // For endpoints that don't require property:value format (like SharePoint sites)
        // Just ensure the value is wrapped in double quotes
        if (!searchValue.startsWith('"') || !searchValue.endsWith('"')) {
          queryParams['$search'] = `"${cleanSearchValue}"`;
          logger.info(
            `Wrapped search query in quotes for ${tool.alias}: "${searchValue}" -> "${queryParams['$search']}"`
          );
        }
      }

      // Set ConsistencyLevel header if required for this endpoint
      if (requiresConsistencyLevelHeader.includes(tool.alias)) {
        headers['ConsistencyLevel'] = 'eventual';
        logger.info(`Setting ConsistencyLevel header to "eventual" for ${tool.alias} search`);
      }

      // Microsoft Graph API limitation: $orderby is NOT supported with $search
      // When both are present, remove $orderby (search results use relevance ranking)
      if (queryParams['$orderby']) {
        logger.warn(
          `Removing $orderby parameter for ${tool.alias}: Microsoft Graph API does not support $orderby with $search. ` +
            `Search results will be ordered by relevance instead of "${queryParams['$orderby']}".`
        );
        delete queryParams['$orderby'];
      }
    }

    // Handle timezone parameter for calendar endpoints
    if (config?.supportsTimezone && params.timezone) {
      headers['Prefer'] = `outlook.timezone="${params.timezone}"`;
      logger.info(`Setting timezone header: Prefer: outlook.timezone="${params.timezone}"`);
    }

    if (Object.keys(queryParams).length > 0) {
      const queryString = Object.entries(queryParams)
        .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
        .join('&');
      path = `${path}${path.includes('?') ? '&' : '?'}${queryString}`;
    }

    const options: {
      method: string;
      headers: Record<string, string>;
      body?: string;
      rawResponse?: boolean;
      includeHeaders?: boolean;
      excludeResponse?: boolean;
      queryParams?: Record<string, string>;
    } = {
      method: tool.method.toUpperCase(),
      headers,
    };

    if (options.method !== 'GET' && body) {
      options.body = typeof body === 'string' ? body : JSON.stringify(body);
    }

    const isProbablyMediaContent =
      tool.errors?.some((error) => error.description === 'Retrieved media content') ||
      path.endsWith('/content');

    if (config?.returnDownloadUrl && path.endsWith('/content')) {
      path = path.replace(/\/content$/, '');
      logger.info(
        `Auto-returning download URL for ${tool.alias} (returnDownloadUrl=true in endpoints.json)`
      );
    } else if (isProbablyMediaContent) {
      options.rawResponse = true;
    }

    // Set includeHeaders if requested
    if (params.includeHeaders === true) {
      options.includeHeaders = true;
    }

    // Set excludeResponse if requested
    if (params.excludeResponse === true) {
      options.excludeResponse = true;
    }

    logger.info(`Making graph request to ${path} with options: ${JSON.stringify(options)}`);
    let response = await graphClient.graphRequest(path, options);

    const fetchAllPages = params.fetchAllPages === true;
    if (fetchAllPages && response?.content?.[0]?.text) {
      try {
        let combinedResponse = JSON.parse(response.content[0].text);
        let allItems = combinedResponse.value || [];
        let nextLink = combinedResponse['@odata.nextLink'];
        let pageCount = 1;

        const maxPages = parseInt(process.env.MS365_MCP_MAX_PAGES || '500', 10); // Default 500, configurable via ENV

        while (nextLink && pageCount < maxPages) {
          logger.info(`Fetching page ${pageCount + 1} from: ${nextLink}`);

          const url = new URL(nextLink);
          const nextPath = url.pathname.replace('/v1.0', '');
          const nextOptions = { ...options };

          const nextQueryParams: Record<string, string> = {};
          for (const [key, value] of url.searchParams.entries()) {
            nextQueryParams[key] = value;
          }
          nextOptions.queryParams = nextQueryParams;

          const nextResponse = await graphClient.graphRequest(nextPath, nextOptions);
          if (nextResponse?.content?.[0]?.text) {
            const nextJsonResponse = JSON.parse(nextResponse.content[0].text);
            if (nextJsonResponse.value && Array.isArray(nextJsonResponse.value)) {
              allItems = allItems.concat(nextJsonResponse.value);
            }
            nextLink = nextJsonResponse['@odata.nextLink'];
            pageCount++;
          } else {
            break;
          }
        }

        if (pageCount >= maxPages) {
          logger.warn(`Reached maximum page limit (${maxPages}) for pagination`);
        }

        combinedResponse.value = allItems;
        if (combinedResponse['@odata.count']) {
          combinedResponse['@odata.count'] = allItems.length;
        }
        delete combinedResponse['@odata.nextLink'];

        response.content[0].text = JSON.stringify(combinedResponse);

        logger.info(
          `Pagination complete: collected ${allItems.length} items across ${pageCount} pages`
        );
      } catch (e) {
        logger.error(`Error during pagination: ${e}`);
      }
    }

    if (response?.content?.[0]?.text) {
      const responseText = response.content[0].text;
      logger.info(`Response size: ${responseText.length} characters`);

      try {
        const jsonResponse = JSON.parse(responseText);
        if (jsonResponse.value && Array.isArray(jsonResponse.value)) {
          logger.info(`Response contains ${jsonResponse.value.length} items`);
        }
        if (jsonResponse['@odata.nextLink']) {
          logger.info(`Response has pagination nextLink: ${jsonResponse['@odata.nextLink']}`);
        }
      } catch {
        // Non-JSON response
      }
    }

    // Convert McpResponse to CallToolResult with the correct structure
    const content: ContentItem[] = response.content.map((item) => ({
      type: 'text' as const,
      text: item.text,
    }));

    // Track tool usage for learning
    if (toolUsageTracker.knowledgeBase) {
      const executionTime = Date.now() - startTime;
      const success = !response.isError;

      // Count results if available
      let resultsCount = 0;
      try {
        const responseText = response.content[0]?.text;
        if (responseText) {
          const jsonResponse = JSON.parse(responseText);
          if (Array.isArray(jsonResponse.value)) {
            resultsCount = jsonResponse.value.length;
          } else if (jsonResponse.value && typeof jsonResponse.value === 'object') {
            resultsCount = 1;
          }
        }
      } catch {
        // Non-JSON or parse error - ignore
      }

      // Record tool usage with other tools in current session
      const usedWith = toolUsageTracker.currentSession.filter((t) => t !== tool.alias);
      toolUsageTracker.knowledgeBase.recordToolUsage(tool.alias, usedWith, success, resultsCount);

      // Add to current session
      if (!toolUsageTracker.currentSession.includes(tool.alias)) {
        toolUsageTracker.currentSession.push(tool.alias);
      }
      toolUsageTracker.sessionStartTime = Date.now();
    }

    return {
      content,
      _meta: response._meta,
      isError: response.isError,
    };
  } catch (error) {
    const errorMessage = (error as Error).message;
    logger.error(`Error in tool ${tool.alias}: ${errorMessage}`);

    // Check if this is an authentication error and provide a clear message
    const isAuthError =
      errorMessage.includes('AUTHENTICATION REQUIRED') ||
      errorMessage.includes('No access token') ||
      errorMessage.includes('not logged in') ||
      (error as { name?: string }).name === 'AuthenticationError';

    if (isAuthError) {
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify({
              error: 'AUTHENTICATION REQUIRED',
              message:
                'You must log in to Microsoft 365 before using this tool. ' +
                'Please call the "login" tool first and follow the device code instructions.',
              action_required: 'Call the "login" tool to authenticate',
              tool_to_call: 'login',
            }),
          },
        ],
        isError: true,
      };
    }

    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            error: `Error in tool ${tool.alias}: ${errorMessage}`,
          }),
        },
      ],
      isError: true,
    };
  }
}

export function registerGraphTools(
  server: McpServer,
  graphClient: GraphClient,
  readOnly: boolean = false,
  enabledToolsPattern?: string,
  orgMode: boolean = false,
  knowledgeBase?: KnowledgeBase
): number {
  // Set knowledge base for tool usage tracking
  if (knowledgeBase) {
    toolUsageTracker.knowledgeBase = knowledgeBase;
  }
  let enabledToolsRegex: RegExp | undefined;
  if (enabledToolsPattern) {
    try {
      enabledToolsRegex = new RegExp(enabledToolsPattern, 'i');
      logger.info(`Tool filtering enabled with pattern: ${enabledToolsPattern}`);
    } catch {
      logger.error(`Invalid tool filter regex pattern: ${enabledToolsPattern}. Ignoring filter.`);
    }
  }

  let registeredCount = 0;
  let skippedCount = 0;
  let failedCount = 0;

  for (const tool of api.endpoints) {
    const endpointConfig = endpointsData.find((e) => e.toolName === tool.alias);
    if (!orgMode && endpointConfig && !endpointConfig.scopes && endpointConfig.workScopes) {
      logger.info(`Skipping work account tool ${tool.alias} - not in org mode`);
      skippedCount++;
      continue;
    }

    if (readOnly && tool.method.toUpperCase() !== 'GET') {
      logger.info(`Skipping write operation ${tool.alias} in read-only mode`);
      skippedCount++;
      continue;
    }

    if (enabledToolsRegex && !enabledToolsRegex.test(tool.alias)) {
      logger.info(`Skipping tool ${tool.alias} - doesn't match filter pattern`);
      skippedCount++;
      continue;
    }

    const paramSchema: Record<string, z.ZodTypeAny> = {};
    if (tool.parameters && tool.parameters.length > 0) {
      for (const param of tool.parameters) {
        // Make startDateTime and endDateTime optional for calendar tools - no default date filter
        if (
          (param.name === 'startDateTime' || param.name === 'endDateTime') &&
          tool.path.includes('calendar')
        ) {
          const defaultDescription =
            param.name === 'startDateTime'
              ? 'The start date and time in ISO 8601 format. Optional - if not provided, no date filter is applied and all events are returned.'
              : 'The end date and time in ISO 8601 format. Optional - if not provided, no date filter is applied and all events are returned.';
          paramSchema[param.name] = z.string().describe(defaultDescription).optional();
        } else {
          paramSchema[param.name] = param.schema || z.any();
        }
      }
    }

    if (tool.method.toUpperCase() === 'GET' && tool.path.includes('/')) {
      paramSchema['fetchAllPages'] = z
        .boolean()
        .describe('Automatically fetch all pages of results')
        .optional();
    }

    // Add includeHeaders parameter for all tools to capture ETags and other headers
    paramSchema['includeHeaders'] = z
      .boolean()
      .describe('Include response headers (including ETag) in the response metadata')
      .optional();

    // Add excludeResponse parameter to only return success/failure indication
    paramSchema['excludeResponse'] = z
      .boolean()
      .describe('Exclude the full response body and only return success or failure indication')
      .optional();

    // Add timezone parameter for calendar endpoints that support it
    if (endpointConfig?.supportsTimezone) {
      paramSchema['timezone'] = z
        .string()
        .describe(
          'IANA timezone name (e.g., "America/New_York", "Europe/London", "Asia/Tokyo") for calendar event times. If not specified, times are returned in UTC.'
        )
        .optional();
    }

    // Build the tool description, optionally appending LLM tips
    let toolDescription =
      tool.description || `Execute ${tool.method.toUpperCase()} request to ${tool.path}`;
    if (endpointConfig?.llmTip) {
      toolDescription += `\n\n💡 TIP: ${endpointConfig.llmTip}`;
    }

    // Add authentication reminder to tool description
    toolDescription +=
      '\n\n⚠️ REQUIRES AUTHENTICATION: You must call the "login" tool first if not already authenticated.';

    try {
      server.tool(
        tool.alias,
        toolDescription,
        paramSchema,
        {
          title: tool.alias,
          readOnlyHint: tool.method.toUpperCase() === 'GET',
          destructiveHint: ['POST', 'PATCH', 'DELETE'].includes(tool.method.toUpperCase()),
          openWorldHint: true, // All tools call Microsoft Graph API
        },
        async (params) => executeGraphTool(tool, endpointConfig, graphClient, params)
      );
      registeredCount++;
    } catch (error) {
      logger.error(`Failed to register tool ${tool.alias}: ${(error as Error).message}`);
      failedCount++;
    }
  }

  logger.info(
    `Tool registration complete: ${registeredCount} registered, ${skippedCount} skipped, ${failedCount} failed`
  );
  return registeredCount;
}

function buildToolsRegistry(
  readOnly: boolean,
  orgMode: boolean
): Map<string, { tool: (typeof api.endpoints)[0]; config: EndpointConfig | undefined }> {
  const toolsMap = new Map<
    string,
    { tool: (typeof api.endpoints)[0]; config: EndpointConfig | undefined }
  >();

  for (const tool of api.endpoints) {
    const endpointConfig = endpointsData.find((e) => e.toolName === tool.alias);

    if (!orgMode && endpointConfig && !endpointConfig.scopes && endpointConfig.workScopes) {
      continue;
    }

    if (readOnly && tool.method.toUpperCase() !== 'GET') {
      continue;
    }

    toolsMap.set(tool.alias, { tool, config: endpointConfig });
  }

  return toolsMap;
}

export function registerDiscoveryTools(
  server: McpServer,
  graphClient: GraphClient,
  readOnly: boolean = false,
  orgMode: boolean = false
): void {
  const toolsRegistry = buildToolsRegistry(readOnly, orgMode);
  logger.info(`Discovery mode: ${toolsRegistry.size} tools available in registry`);

  server.tool(
    'search-tools',
    `Search through ${toolsRegistry.size} available Microsoft Graph API tools. Use this to find tools by name, path, or description before executing them.`,
    {
      query: z
        .string()
        .describe('Search query to filter tools (searches name, path, and description)')
        .optional(),
      category: z
        .string()
        .describe(
          'Filter by category: mail, calendar, files, contacts, tasks, onenote, search, users, excel'
        )
        .optional(),
      limit: z.number().describe('Maximum results to return (default: 20, max: 50)').optional(),
    },
    {
      title: 'search-tools',
      readOnlyHint: true,
      openWorldHint: true, // Searches Microsoft Graph API tools
    },
    async ({ query, category, limit = 20 }) => {
      const maxLimit = Math.min(limit, 50);
      const results: Array<{
        name: string;
        method: string;
        path: string;
        description: string;
      }> = [];

      const queryLower = query?.toLowerCase();
      const categoryDef = category ? TOOL_CATEGORIES[category] : undefined;

      for (const [name, { tool, config }] of toolsRegistry) {
        if (categoryDef && !categoryDef.pattern.test(name)) {
          continue;
        }

        if (queryLower) {
          const searchText =
            `${name} ${tool.path} ${tool.description || ''} ${config?.llmTip || ''}`.toLowerCase();
          if (!searchText.includes(queryLower)) {
            continue;
          }
        }

        results.push({
          name,
          method: tool.method.toUpperCase(),
          path: tool.path,
          description: tool.description || `${tool.method.toUpperCase()} ${tool.path}`,
        });

        if (results.length >= maxLimit) break;
      }

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                found: results.length,
                total: toolsRegistry.size,
                tools: results,
                tip: 'Use execute-tool with the tool name and required parameters to call any of these tools.',
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );

  server.tool(
    'execute-tool',
    'Execute a Microsoft Graph API tool by name. Use search-tools first to find available tools and their parameters.',
    {
      tool_name: z.string().describe('Name of the tool to execute (e.g., "list-mail-messages")'),
      parameters: z
        .record(z.any())
        .describe('Parameters to pass to the tool as key-value pairs')
        .optional(),
    },
    {
      title: 'execute-tool',
      readOnlyHint: false,
      destructiveHint: true, // Can execute any tool, including write operations
      openWorldHint: true, // Executes against Microsoft Graph API
    },
    async ({ tool_name, parameters = {} }) => {
      const toolData = toolsRegistry.get(tool_name);
      if (!toolData) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error: `Tool not found: ${tool_name}`,
                tip: 'Use search-tools to find available tools.',
              }),
            },
          ],
          isError: true,
        };
      }

      return executeGraphTool(toolData.tool, toolData.config, graphClient, parameters);
    }
  );
}
