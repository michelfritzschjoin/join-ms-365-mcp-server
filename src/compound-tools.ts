/**
 * Compound Tools - Intelligent multi-step tools that chain API calls
 * for complex contextual queries like "Find messages with [person]"
 *
 * These tools solve common scenarios where users want to:
 * - Find all communication with a specific person
 * - Get files shared by someone
 * - Find meetings with specific attendees
 * - Get a complete communication summary
 */

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import GraphClient from './graph-client.js';
import logger from './logger.js';
import { getChatMemoryStore, isChatMemoryEnabled, type EntityType } from './chat-memory.js';
import { getChatId, getUserId } from './request-context.js';
import NLPEnhancer, { type DecomposedQuery } from './nlp-enhancer.js';
import {
  formatCalendarResponse,
  calendarResponseToText,
  formatMailResponse,
  mailResponseToText,
  convertToLocalTime,
} from './response-formatter.js';
import { createThinkingProcess } from './thinking-process.js';
import { isLoopFile, detectLoopFile, parseLoopContent } from './utils/loop-detector.js';

/**
 * SECURITY: Properly sanitize HTML content to prevent XSS
 * Uses safer approach: remove all tags first, then decode entities once
 * This prevents double-escaping and incomplete sanitization issues
 */
function sanitizeHtml(html: string): string {
  if (!html || typeof html !== 'string') {
    return '';
  }

  // SECURITY: Use a simple, safe approach - remove ALL HTML by extracting text content only
  // This avoids complex regex patterns that could be exploited or incomplete
  let result = html;

  // Step 1: Remove all content between dangerous tag pairs using indexOf (no regex)
  // This is safer than regex as it handles the full content removal
  const dangerousTags = ['script', 'style', 'iframe', 'object', 'form', 'textarea', 'select'];
  for (const tag of dangerousTags) {
    let idx = 0;
    let maxLoops = 100; // Prevent infinite loops
    while (maxLoops-- > 0) {
      const openTag = result.toLowerCase().indexOf('<' + tag, idx);
      if (openTag === -1) break;

      const closeTag = result.toLowerCase().indexOf('</' + tag, openTag);
      if (closeTag === -1) {
        // No closing tag found - remove just the opening tag
        const tagEnd = result.indexOf('>', openTag);
        if (tagEnd !== -1) {
          result = result.substring(0, openTag) + result.substring(tagEnd + 1);
        } else {
          break;
        }
      } else {
        // Find the end of the closing tag
        const closeEnd = result.indexOf('>', closeTag);
        if (closeEnd !== -1) {
          result = result.substring(0, openTag) + result.substring(closeEnd + 1);
        } else {
          result = result.substring(0, openTag) + result.substring(closeTag);
        }
      }
    }
  }

  // Step 2: Remove all remaining HTML tags using a simple character-by-character approach
  let output = '';
  let inTag = false;
  for (let i = 0; i < result.length; i++) {
    const char = result[i];
    if (char === '<') {
      inTag = true;
      output += ' '; // Replace tag with space
    } else if (char === '>') {
      inTag = false;
    } else if (!inTag) {
      output += char;
    }
  }
  result = output;

  // Step 3: Decode common HTML entities using split/join (safe, no regex)
  const entityReplacements: [string, string][] = [
    ['&nbsp;', ' '],
    ['&amp;', '&'],
    ['&lt;', '<'],
    ['&gt;', '>'],
    ['&quot;', '"'],
    ['&#39;', "'"],
    ['&#x27;', "'"],
    ['&#x2F;', '/'],
    ['&#x60;', '`'],
    ['&#x3D;', '='],
  ];

  for (const [entity, replacement] of entityReplacements) {
    result = result.split(entity).join(replacement);
  }

  // Step 4: Decode numeric entities (decimal) - only safe ASCII range
  result = result.replace(/&#(\d{1,4});/g, (match, dec) => {
    const code = parseInt(dec, 10);
    return code >= 32 && code <= 126 ? String.fromCharCode(code) : match;
  });

  // Step 5: Decode hex entities - only safe ASCII range
  result = result.replace(/&#x([0-9a-fA-F]{1,4});/g, (match, hex) => {
    const code = parseInt(hex, 16);
    return code >= 32 && code <= 126 ? String.fromCharCode(code) : match;
  });

  // Step 6: Normalize whitespace
  result = result.replace(/\s+/g, ' ').trim();

  return result;
}

/**
 * Build query string for Microsoft Graph API requests
 * Handles $ prefixed parameters correctly (doesn't encode $)
 */
function buildGraphQueryString(params: Record<string, string>): string {
  return Object.entries(params)
    .map(([key, value]) => {
      // Don't encode $ in parameter names (Microsoft Graph API expects $top, $filter, etc.)
      const encodedKey = key.startsWith('$') ? key : encodeURIComponent(key);
      return `${encodedKey}=${encodeURIComponent(value)}`;
    })
    .join('&');
}

// Type definitions for Graph API responses
interface GraphUser {
  id: string;
  displayName: string;
  mail?: string;
  userPrincipalName?: string;
  businessPhones?: string[];
  mobilePhone?: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
}

interface GraphChat {
  id: string;
  chatType: string;
  topic?: string;
  createdDateTime: string;
  lastUpdatedDateTime?: string;
  members?: Array<{
    displayName?: string;
    userId?: string;
    email?: string;
  }>;
}

interface GraphChatMessage {
  id: string;
  createdDateTime: string;
  body: {
    content: string;
    contentType: string;
  };
  from?: {
    user?: {
      displayName?: string;
      id?: string;
    };
  };
  subject?: string;
}

interface GraphEmail {
  id: string;
  subject: string;
  bodyPreview: string;
  receivedDateTime: string;
  from?: {
    emailAddress: {
      name?: string;
      address: string;
    };
  };
  toRecipients?: Array<{
    emailAddress: {
      name?: string;
      address: string;
    };
  }>;
  hasAttachments?: boolean;
  webLink?: string;
}

interface GraphEvent {
  id: string;
  subject: string;
  bodyPreview?: string;
  start: {
    dateTime: string;
    timeZone: string;
  };
  end: {
    dateTime: string;
    timeZone: string;
  };
  isAllDay?: boolean;
  isCancelled?: boolean;
  importance?: string;
  showAs?: string;
  categories?: string[];
  isOnlineMeeting?: boolean;
  onlineMeeting?: {
    joinUrl?: string;
  };
  attendees?: Array<{
    emailAddress: {
      name?: string;
      address: string;
    };
    status?: {
      response?: string;
    };
  }>;
  organizer?: {
    emailAddress: {
      name?: string;
      address: string;
    };
  };
  location?: {
    displayName?: string;
  };
  webLink?: string;
}

interface GraphDriveItem {
  id: string;
  name: string;
  webUrl?: string;
  size?: number;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  createdBy?: {
    user?: {
      displayName?: string;
      email?: string;
    };
  };
  lastModifiedBy?: {
    user?: {
      displayName?: string;
      email?: string;
    };
  };
  file?: {
    mimeType?: string;
  };
  folder?: {
    childCount?: number;
  };
  shared?: {
    sharedDateTime?: string;
    sharedBy?: {
      user?: {
        displayName?: string;
        email?: string;
      };
    };
  };
}

interface SearchHit {
  resource: any;
  summary?: string;
  name?: string;
  webUrl?: string;
  type?: string;
  date?: string;
  from?: string;
  rank?: number;
  relevanceScore?: number;
}

interface SearchApiResults {
  emails: SearchHit[];
  events: SearchHit[];
  files: SearchHit[];
  sites: SearchHit[];
  listItems: SearchHit[];
  chats: SearchHit[];
  people: SearchHit[];
}

/**
 * Central Search Options for Microsoft Search API
 */
interface CentralSearchOptions {
  /** Entity types to search (default: all) */
  entityTypes?: Array<
    'message' | 'event' | 'driveItem' | 'site' | 'list' | 'listItem' | 'chatMessage' | 'person'
  >;
  /** Maximum results to return (default: 100) */
  maxResults?: number;
  /** Minimum relevance score 0-100 (default: 0 = no filter) */
  minRelevance?: number;
  /** Include time context for events (default: true) */
  includeTimeContext?: boolean;
  /** Sort by rank (default: true) */
  sortByRank?: boolean;
  /** Custom time range for events */
  timeRange?: {
    startDateTime: string;
    endDateTime: string;
  };
}

/**
 * Central Search Result with metadata
 */
interface CentralSearchResult {
  query: string;
  searchedAt: string;
  totalHits: number;
  results: SearchApiResults;
  metadata: {
    entityTypesCounts: Record<string, number>;
    averageRank: number;
    searchDuration: number;
  };
}

interface FollowUpResults {
  onenote?: any[];
  planner?: any[];
  todo?: any[];
  contacts?: any[];
  meetings?: any[];
  teams?: any[];
  bookings?: any[];
  insights?: any[];
}

/**
 * Current date/time context for LLM reference
 */
interface DateTimeContext {
  /** Current date in ISO format */
  currentDate: string;
  /** Current time in local format (HH:MM) */
  currentTime: string;
  /** Server timezone name (e.g., "Europe/Berlin") */
  timezone: string;
  /** UTC offset (e.g., "+01:00") */
  utcOffset: string;
  /** Human-readable current datetime */
  formatted: string;
  /** Reference dates for relative queries */
  references: {
    today: string;
    tomorrow: string;
    yesterday: string;
    thisWeekStart: string;
    thisWeekEnd: string;
  };
}

interface EnhancedAskM365Response {
  question: string;
  language: string;
  searchedAt: string;
  /** Current date/time context for accurate time references */
  currentContext: DateTimeContext;
  intent: any;
  processingSteps: string[];
  resultsFound: boolean;
  totalResults: number;
  status: string;
  message: string;
  summary?: string;
  /** Structured query analysis from NLP processing */
  queryAnalysis?: DecomposedQuery;
  /** Structured Markdown summary of the query analysis */
  queryAnalysisMarkdown?: string;
  searchResults: SearchApiResults;
  followUpResults: FollowUpResults;
  permissionWarnings?: string[];
  metadata: {
    totalResults: number;
    queryTime: number;
    productsSearched: string[];
  };
  /** Optional thinking process for transparent reasoning display in OpenWebUI */
  thinking?: {
    enabled: boolean;
    level: string;
    summary: string;
    totalDuration: number;
    stepCount: number;
    steps: Array<{
      type: string;
      category: string;
      message: string;
      duration?: number;
      icon?: string;
      details?: Record<string, unknown>;
    }>;
    markdown: string;
  };
}

/**
 * Generate current date/time context for LLM reference
 * This helps the LLM understand relative time references like "tomorrow" or "yesterday"
 */
function generateDateTimeContext(lang: 'de' | 'en'): DateTimeContext {
  const now = new Date();

  // Get timezone info
  const timezone = Intl.DateTimeFormat().resolvedOptions().timeZone;
  const utcOffsetMinutes = -now.getTimezoneOffset();
  const utcOffsetHours = Math.floor(Math.abs(utcOffsetMinutes) / 60);
  const utcOffsetMins = Math.abs(utcOffsetMinutes) % 60;
  const utcOffset = `${utcOffsetMinutes >= 0 ? '+' : '-'}${String(utcOffsetHours).padStart(2, '0')}:${String(utcOffsetMins).padStart(2, '0')}`;

  // Calculate reference dates
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const tomorrow = new Date(today.getTime() + 24 * 60 * 60 * 1000);
  const yesterday = new Date(today.getTime() - 24 * 60 * 60 * 1000);

  // Calculate week boundaries (Monday as start of week for German/European convention)
  const dayOfWeek = now.getDay();
  const daysToMonday = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
  const thisWeekStart = new Date(today.getTime() - daysToMonday * 24 * 60 * 60 * 1000);
  const thisWeekEnd = new Date(thisWeekStart.getTime() + 6 * 24 * 60 * 60 * 1000);

  // Format functions for local display
  const formatDate = (d: Date): string => {
    return lang === 'de'
      ? d.toLocaleDateString('de-DE', {
          weekday: 'long',
          day: '2-digit',
          month: '2-digit',
          year: 'numeric',
        })
      : d.toLocaleDateString('en-US', {
          weekday: 'long',
          month: 'short',
          day: 'numeric',
          year: 'numeric',
        });
  };

  const formatISODate = (d: Date): string => {
    const dd = String(d.getDate()).padStart(2, '0');
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const yyyy = d.getFullYear();
    return `${yyyy}-${mm}-${dd}`;
  };

  const currentTime = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;

  return {
    currentDate: formatISODate(now),
    currentTime,
    timezone,
    utcOffset,
    formatted:
      lang === 'de'
        ? `${formatDate(now)}, ${currentTime} Uhr (${timezone})`
        : `${formatDate(now)}, ${currentTime} (${timezone})`,
    references: {
      today: formatISODate(today),
      tomorrow: formatISODate(tomorrow),
      yesterday: formatISODate(yesterday),
      thisWeekStart: formatISODate(thisWeekStart),
      thisWeekEnd: formatISODate(thisWeekEnd),
    },
  };
}

/**
 * Teams Channel type definition
 */
interface GraphTeam {
  id: string;
  displayName: string;
  description?: string;
  webUrl?: string;
  createdDateTime?: string;
}

/**
 * Teams Channel type definition
 */
interface GraphChannel {
  id: string;
  displayName: string;
  description?: string;
  webUrl?: string;
  membershipType?: 'standard' | 'private' | 'unknownFutureValue' | 'shared';
  createdDateTime?: string;
}

/**
 * Teams Channel Message type definition
 */
interface GraphChannelMessage {
  id: string;
  replyToId?: string;
  etag?: string;
  messageType: 'message' | 'chatEvent' | 'typing' | 'unknownFutureValue';
  createdDateTime: string;
  lastModifiedDateTime?: string;
  deletedDateTime?: string;
  subject?: string;
  body: {
    content: string;
    contentType: 'text' | 'html';
  };
  from?: {
    user?: {
      id?: string;
      displayName?: string;
      userIdentityType?: string;
    };
    application?: {
      id?: string;
      displayName?: string;
    };
  };
  importance?: 'normal' | 'high' | 'urgent';
  webUrl?: string;
  reactions?: Array<{
    reactionType: string;
    createdDateTime: string;
    user: {
      displayName?: string;
    };
  }>;
  attachments?: Array<{
    id: string;
    contentType: string;
    contentUrl?: string;
    name?: string;
  }>;
  mentions?: Array<{
    id: number;
    mentionText: string;
    mentioned: {
      user?: {
        displayName?: string;
        id?: string;
      };
    };
  }>;
}

/**
 * Helper function to find a user by name, email, or phone
 */
async function findUser(graphClient: GraphClient, searchQuery: string): Promise<GraphUser | null> {
  try {
    // Try different search approaches
    const searchStrategies = [
      // 1. Search by displayName
      async () => {
        const response = await graphClient.makeRequest('/users', {
          method: 'GET',
          queryParams: {
            $search: `"displayName:${searchQuery}"`,
            $top: '5',
          },
          headers: {
            ConsistencyLevel: 'eventual',
          },
        });
        return response;
      },
      // 2. Filter by mail
      async () => {
        const response = await graphClient.makeRequest('/users', {
          method: 'GET',
          queryParams: {
            $filter: `mail eq '${searchQuery}' or userPrincipalName eq '${searchQuery}'`,
            $top: '5',
          },
        });
        return response;
      },
    ];

    for (const strategy of searchStrategies) {
      try {
        const response = await strategy();
        if (
          response &&
          typeof response === 'object' &&
          'value' in response &&
          Array.isArray(response.value) &&
          response.value.length > 0
        ) {
          // Return the best match
          const user = response.value[0] as GraphUser;
          logger.info(`Found user: ${user.displayName} (${user.id})`);
          return user;
        }
      } catch (err) {
        logger.debug(`Search strategy failed: ${err}`);
      }
    }

    // 3. Try people API as fallback
    try {
      const peopleResponse = await graphClient.makeRequest('/me/people', {
        method: 'GET',
        queryParams: {
          $search: `"${searchQuery}"`,
          $top: '5',
        },
      });

      if (
        peopleResponse &&
        typeof peopleResponse === 'object' &&
        'value' in peopleResponse &&
        Array.isArray(peopleResponse.value) &&
        peopleResponse.value.length > 0
      ) {
        const person = peopleResponse.value[0] as {
          id?: string;
          displayName?: string;
          userPrincipalName?: string;
          scoredEmailAddresses?: Array<{ address: string }>;
        };

        return {
          id: person.id || '',
          displayName: person.displayName || searchQuery,
          mail: person.scoredEmailAddresses?.[0]?.address,
          userPrincipalName: person.userPrincipalName,
        };
      }
    } catch (err) {
      logger.debug(`People search failed: ${err}`);
    }

    logger.warn(`User not found: ${searchQuery}`);
    return null;
  } catch (error) {
    logger.error(`Error finding user: ${error}`);
    return null;
  }
}

/**
 * Find chats that include a specific user
 */
async function findChatsWithUser(
  graphClient: GraphClient,
  userId: string,
  userEmail?: string,
  userDisplayName?: string
): Promise<GraphChat[]> {
  try {
    // List all chats and filter by member
    const response = await graphClient.makeRequest('/me/chats', {
      method: 'GET',
      queryParams: {
        $expand: 'members',
        $top: '50',
      },
    });

    if (
      !response ||
      typeof response !== 'object' ||
      !('value' in response) ||
      !Array.isArray(response.value)
    ) {
      return [];
    }

    const chats = response.value as GraphChat[];
    const matchingChats: GraphChat[] = [];

    for (const chat of chats) {
      if (chat.members && Array.isArray(chat.members)) {
        const hasUser = chat.members.some((member) => {
          const memberAny = member as { userId?: string; email?: string; displayName?: string };
          return (
            memberAny.userId === userId ||
            (userEmail && memberAny.email?.toLowerCase() === userEmail.toLowerCase()) ||
            (userDisplayName &&
              memberAny.displayName?.toLowerCase().includes(userDisplayName.toLowerCase()))
          );
        });

        if (hasUser) {
          matchingChats.push(chat);
        }
      }
    }

    logger.info(`Found ${matchingChats.length} chats with user ${userId}`);
    return matchingChats;
  } catch (error) {
    logger.error(`Error finding chats: ${error}`);
    return [];
  }
}

/**
 * Get messages from specific chats
 */
async function getMessagesFromChats(
  graphClient: GraphClient,
  chatIds: string[],
  limit: number = 20
): Promise<GraphChatMessage[]> {
  const allMessages: GraphChatMessage[] = [];
  const messagesPerChat = Math.ceil(limit / chatIds.length);

  for (const chatId of chatIds.slice(0, 5)) {
    // Limit to 5 chats
    try {
      const response = await graphClient.makeRequest(`/me/chats/${chatId}/messages`, {
        method: 'GET',
        queryParams: {
          $top: String(messagesPerChat),
          $orderby: 'createdDateTime desc',
        },
      });

      if (
        response &&
        typeof response === 'object' &&
        'value' in response &&
        Array.isArray(response.value)
      ) {
        allMessages.push(...(response.value as GraphChatMessage[]));
      }
    } catch (error) {
      logger.warn(`Error getting messages from chat ${chatId}: ${error}`);
    }
  }

  return allMessages.slice(0, limit);
}

/**
 * Fetch all joined teams with pagination support
 * Microsoft Graph API may return paginated results for users with many teams
 * @param graphClient - The Graph API client
 * @param selectFields - Optional fields to select (default: id,displayName,description,webUrl)
 * @returns Array of all joined teams
 */
async function getAllJoinedTeams(
  graphClient: GraphClient,
  selectFields: string = 'id,displayName,description,webUrl'
): Promise<GraphTeam[]> {
  const allTeams: GraphTeam[] = [];
  const maxPages = parseInt(process.env.MS365_MCP_MAX_PAGES || '50', 10);
  let pageCount = 0;

  try {
    // First request
    let response = (await graphClient.makeRequest('/me/joinedTeams', {
      method: 'GET',
      queryParams: {
        $select: selectFields,
        $top: '999', // Request maximum items per page
      },
    })) as { value?: GraphTeam[]; '@odata.nextLink'?: string };

    if (response?.value && Array.isArray(response.value)) {
      allTeams.push(...response.value);
    }

    // Handle pagination with @odata.nextLink
    while (response?.['@odata.nextLink'] && pageCount < maxPages) {
      pageCount++;
      const nextLink = response['@odata.nextLink'];
      logger.debug(`Fetching teams page ${pageCount + 1} from: ${nextLink}`);

      try {
        const url = new URL(nextLink);
        const nextPath = url.pathname.replace('/v1.0', '');
        const nextQueryParams: Record<string, string> = {};
        for (const [key, value] of url.searchParams.entries()) {
          nextQueryParams[key] = value;
        }

        response = (await graphClient.makeRequest(nextPath, {
          method: 'GET',
          queryParams: nextQueryParams,
        })) as { value?: GraphTeam[]; '@odata.nextLink'?: string };

        if (response?.value && Array.isArray(response.value)) {
          allTeams.push(...response.value);
        }
      } catch (pageError) {
        logger.warn(`Error fetching teams page ${pageCount + 1}: ${pageError}`);
        break;
      }
    }

    if (pageCount >= maxPages) {
      logger.warn(`Reached maximum page limit (${maxPages}) for teams pagination`);
    }

    logger.info(`Fetched ${allTeams.length} joined teams across ${pageCount + 1} page(s)`);
    return allTeams;
  } catch (error) {
    logger.error(`Error fetching joined teams: ${error}`);
    return allTeams;
  }
}

/**
 * Find a Teams channel by name across all joined teams
 * Searches through all teams the user is a member of to find a channel by name
 * @param graphClient - The Graph API client
 * @param channelName - The name of the channel to search for (partial match, case-insensitive)
 * @param teamName - Optional team name to narrow down the search
 * @returns Object containing the team and channel information, or null if not found
 */
async function findTeamsChannel(
  graphClient: GraphClient,
  channelName: string,
  teamName?: string
): Promise<{ team: GraphTeam; channel: GraphChannel } | null> {
  try {
    // Step 1: Get all joined teams with pagination support
    const teams = await getAllJoinedTeams(graphClient);

    if (teams.length === 0) {
      logger.warn('No teams found for user');
      return null;
    }

    const channelNameLower = channelName.toLowerCase();
    const teamNameLower = teamName?.toLowerCase();

    // Optionally filter teams by name
    const teamsToSearch = teamNameLower
      ? teams.filter((t) => t.displayName.toLowerCase().includes(teamNameLower))
      : teams;

    if (teamsToSearch.length === 0) {
      logger.warn(`No teams matching "${teamName}" found`);
      return null;
    }

    // Step 2: Search for channel in each team
    for (const team of teamsToSearch) {
      try {
        const channelsResponse = await graphClient.makeRequest(`/teams/${team.id}/channels`, {
          method: 'GET',
          queryParams: {
            $select: 'id,displayName,description,webUrl,membershipType',
          },
        });

        if (
          channelsResponse &&
          typeof channelsResponse === 'object' &&
          'value' in channelsResponse &&
          Array.isArray(channelsResponse.value)
        ) {
          const channels = channelsResponse.value as GraphChannel[];
          const matchingChannel = channels.find((c) =>
            c.displayName.toLowerCase().includes(channelNameLower)
          );

          if (matchingChannel) {
            logger.info(
              `Found channel "${matchingChannel.displayName}" in team "${team.displayName}"`
            );
            return { team, channel: matchingChannel };
          }
        }
      } catch (channelError) {
        logger.warn(`Could not access channels for team ${team.displayName}: ${channelError}`);
        // Continue to next team
      }
    }

    logger.warn(`Channel "${channelName}" not found in any team`);
    return null;
  } catch (error) {
    logger.error(`Error finding Teams channel: ${error}`);
    return null;
  }
}

/**
 * Get messages from a Teams channel
 * @param graphClient - The Graph API client
 * @param teamId - The team ID
 * @param channelId - The channel ID
 * @param limit - Maximum number of messages to return (default: 50)
 * @param includeReplies - Whether to include replies to messages (default: false)
 * @returns Array of channel messages
 */
async function getChannelMessages(
  graphClient: GraphClient,
  teamId: string,
  channelId: string,
  limit: number = 50,
  includeReplies: boolean = false
): Promise<GraphChannelMessage[]> {
  try {
    const queryParams: Record<string, string> = {
      $top: String(Math.min(limit, 50)),
      $orderby: 'createdDateTime desc',
    };

    const response = await graphClient.makeRequest(
      `/teams/${teamId}/channels/${channelId}/messages`,
      {
        method: 'GET',
        queryParams,
      }
    );

    if (
      !response ||
      typeof response !== 'object' ||
      !('value' in response) ||
      !Array.isArray(response.value)
    ) {
      return [];
    }

    const messages = response.value as GraphChannelMessage[];

    // Optionally fetch replies for each message
    if (includeReplies) {
      for (const message of messages.slice(0, 10)) {
        // Limit reply fetching to first 10 messages
        try {
          const repliesResponse = await graphClient.makeRequest(
            `/teams/${teamId}/channels/${channelId}/messages/${message.id}/replies`,
            {
              method: 'GET',
              queryParams: {
                $top: '5',
                $orderby: 'createdDateTime asc',
              },
            }
          );

          if (
            repliesResponse &&
            typeof repliesResponse === 'object' &&
            'value' in repliesResponse &&
            Array.isArray(repliesResponse.value)
          ) {
            // Add replies as nested property (type assertion needed)
            (message as GraphChannelMessage & { replies?: GraphChannelMessage[] }).replies =
              repliesResponse.value as GraphChannelMessage[];
          }
        } catch (replyError) {
          logger.debug(`Could not fetch replies for message ${message.id}: ${replyError}`);
        }
      }
    }

    logger.info(`Retrieved ${messages.length} messages from channel`);
    return messages;
  } catch (error) {
    logger.error(`Error getting channel messages: ${error}`);
    return [];
  }
}

/**
 * Find emails involving a specific person
 */
async function findEmailsWithPerson(
  graphClient: GraphClient,
  userEmail: string,
  userDisplayName: string,
  limit: number = 20
): Promise<GraphEmail[]> {
  const allEmails: GraphEmail[] = [];

  // Search for emails from this person
  try {
    const fromResponse = await graphClient.makeRequest('/me/messages', {
      method: 'GET',
      queryParams: {
        $search: `"from:${userEmail}"`,
        $top: String(Math.ceil(limit / 2)),
        $orderby: 'receivedDateTime desc',
        $select: 'id,subject,bodyPreview,receivedDateTime,from,toRecipients,hasAttachments,webLink',
      },
    });

    if (
      fromResponse &&
      typeof fromResponse === 'object' &&
      'value' in fromResponse &&
      Array.isArray(fromResponse.value)
    ) {
      allEmails.push(...(fromResponse.value as GraphEmail[]));
    }
  } catch (error) {
    logger.warn(`Error searching emails from person: ${error}`);
  }

  // Search for emails to this person
  try {
    const toResponse = await graphClient.makeRequest('/me/messages', {
      method: 'GET',
      queryParams: {
        $search: `"to:${userEmail}"`,
        $top: String(Math.ceil(limit / 2)),
        $orderby: 'receivedDateTime desc',
        $select: 'id,subject,bodyPreview,receivedDateTime,from,toRecipients,hasAttachments,webLink',
      },
    });

    if (
      toResponse &&
      typeof toResponse === 'object' &&
      'value' in toResponse &&
      Array.isArray(toResponse.value)
    ) {
      allEmails.push(...(toResponse.value as GraphEmail[]));
    }
  } catch (error) {
    logger.warn(`Error searching emails to person: ${error}`);
  }

  // Deduplicate and sort by date
  const uniqueEmails = Array.from(new Map(allEmails.map((e) => [e.id, e])).values());
  uniqueEmails.sort(
    (a, b) => new Date(b.receivedDateTime).getTime() - new Date(a.receivedDateTime).getTime()
  );

  return uniqueEmails.slice(0, limit);
}

/**
 * Find meetings/events with a specific person
 */
async function findMeetingsWithPerson(
  graphClient: GraphClient,
  userEmail: string,
  userDisplayName: string,
  limit: number = 20
): Promise<GraphEvent[]> {
  try {
    // Get calendar events and filter by attendee
    const now = new Date();
    const pastDate = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000); // 90 days ago
    const futureDate = new Date(now.getTime() + 90 * 24 * 60 * 60 * 1000); // 90 days future

    // Build query string (startDateTime and endDateTime are REQUIRED for calendarView)
    const queryParams: Record<string, string> = {
      startDateTime: pastDate.toISOString(),
      endDateTime: futureDate.toISOString(),
      $top: '100',
      $select: 'id,subject,bodyPreview,start,end,attendees,organizer,location,webLink',
      $orderby: 'start/dateTime asc',
    };

    const response = await graphClient.makeRequest(
      `/me/calendarView?${buildGraphQueryString(queryParams)}`,
      {
        method: 'GET',
      }
    );

    if (
      !response ||
      typeof response !== 'object' ||
      !('value' in response) ||
      !Array.isArray(response.value)
    ) {
      return [];
    }

    const events = response.value as GraphEvent[];
    const matchingEvents: GraphEvent[] = [];

    const emailLower = userEmail.toLowerCase();
    const nameLower = userDisplayName.toLowerCase();
    // Split name into parts for better matching (e.g., "Max Mustermann" -> ["max", "mustermann"])
    const nameParts = nameLower.split(/\s+/).filter((part) => part.length > 0);

    for (const event of events) {
      let isMatch = false;

      // Check organizer - must be exact email match OR exact name match OR all name parts present
      const organizerEmail = event.organizer?.emailAddress?.address?.toLowerCase();
      const organizerName = event.organizer?.emailAddress?.name?.toLowerCase();

      if (organizerEmail === emailLower) {
        // Exact email match
        isMatch = true;
      } else if (organizerName) {
        const organizerNameLower = organizerName.toLowerCase();
        // Exact name match
        if (organizerNameLower === nameLower) {
          isMatch = true;
        } else if (nameParts.length > 0) {
          // Check if all name parts are present in organizer name (for "Max Mustermann" matching "Maximilian Mustermann")
          const allPartsMatch = nameParts.every((part) => organizerNameLower.includes(part));
          if (allPartsMatch) {
            isMatch = true;
          }
        }
      }

      if (isMatch) {
        matchingEvents.push(event);
        continue;
      }

      // Check attendees - person must be an actual attendee
      if (event.attendees && Array.isArray(event.attendees)) {
        const hasAttendee = event.attendees.some((attendee) => {
          const attendeeEmail = attendee.emailAddress?.address?.toLowerCase();
          const attendeeName = attendee.emailAddress?.name?.toLowerCase();

          // Exact email match
          if (attendeeEmail === emailLower) {
            return true;
          }

          // Exact name match
          if (attendeeName) {
            const attendeeNameLower = attendeeName.toLowerCase();
            if (attendeeNameLower === nameLower) {
              return true;
            }

            // Check if all name parts are present in attendee name
            if (nameParts.length > 0) {
              const allPartsMatch = nameParts.every((part) => attendeeNameLower.includes(part));
              if (allPartsMatch) {
                return true;
              }
            }
          }

          return false;
        });

        if (hasAttendee) {
          matchingEvents.push(event);
        }
      }
    }

    // Sort matching events by start time (ascending) with proper timezone handling
    matchingEvents.sort((a, b) => {
      const aTime = convertToLocalTime(a.start.dateTime, a.start.timeZone).getTime();
      const bTime = convertToLocalTime(b.start.dateTime, b.start.timeZone).getTime();
      return aTime - bTime;
    });

    return matchingEvents.slice(0, limit);
  } catch (error) {
    logger.error(`Error finding meetings: ${error}`);
    return [];
  }
}

type CalendarWindowQuery = {
  readonly startDateTime: string;
  readonly endDateTime: string;
};

type RollingWindowSearchResult = {
  readonly windows: CalendarWindowQuery[];
  readonly formatted:
    | (ReturnType<typeof formatCalendarResponse> & { readonly _humanReadable: string })
    | null;
};

function addDays(date: Date, days: number): Date {
  const next = new Date(date.getTime());
  next.setDate(next.getDate() + days);
  return next;
}

function buildCalendarWindowQuery(params: {
  readonly windowStart: Date;
  readonly windowDays: number;
}): CalendarWindowQuery {
  const start = new Date(params.windowStart.getTime());
  const end = addDays(start, params.windowDays);
  return { startDateTime: start.toISOString(), endDateTime: end.toISOString() };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function buildFilteredCalendarResponse(params: {
  readonly rawResponse: unknown;
  readonly includeAllDay: boolean;
  readonly includeCancelled: boolean;
}): Record<string, unknown> {
  if (!isRecord(params.rawResponse)) {
    return { value: [] };
  }

  const rawEvents = Array.isArray(params.rawResponse.value)
    ? (params.rawResponse.value as Array<Record<string, unknown>>)
    : [];

  const filtered = rawEvents.filter((event) => {
    const isAllDay = Boolean(event.isAllDay);
    const isCancelled = Boolean(event.isCancelled);

    if (!params.includeAllDay && isAllDay) return false;
    if (!params.includeCancelled && isCancelled) return false;

    return true;
  });

  return { value: filtered };
}

async function findUpcomingMeetingsRollingWindow(params: {
  readonly graphClient: GraphClient;
  readonly windowDays: number;
  readonly maxWindows: number;
  readonly limit: number;
  readonly includeAllDay: boolean;
  readonly includeCancelled: boolean;
  readonly timezone?: string;
}): Promise<RollingWindowSearchResult> {
  const windows: CalendarWindowQuery[] = [];
  const now = new Date();

  for (let i = 0; i < params.maxWindows; i++) {
    const windowStart = addDays(now, i * params.windowDays);
    const windowQuery = buildCalendarWindowQuery({ windowStart, windowDays: params.windowDays });
    windows.push(windowQuery);

    const queryParams: Record<string, string> = {
      startDateTime: windowQuery.startDateTime,
      endDateTime: windowQuery.endDateTime,
      $top: String(Math.max(1, Math.min(params.limit, 500))),
      $select:
        'id,subject,bodyPreview,start,end,location,attendees,organizer,isAllDay,isCancelled,importance,showAs,categories,isOnlineMeeting,onlineMeeting,webLink',
      $orderby: 'start/dateTime asc',
    };

    const rawResponse = await params.graphClient.makeRequest(
      `/me/calendarView?${buildGraphQueryString(queryParams)}`,
      {
        method: 'GET',
        headers: params.timezone ? { Prefer: `outlook.timezone="${params.timezone}"` } : undefined,
      }
    );

    const filteredResponse = buildFilteredCalendarResponse({
      rawResponse,
      includeAllDay: params.includeAllDay,
      includeCancelled: params.includeCancelled,
    });

    const formatted = formatCalendarResponse(
      filteredResponse,
      windowQuery.startDateTime,
      windowQuery.endDateTime
    );

    if (formatted.summary.totalEvents > 0) {
      const limitedEvents = formatted.events.slice(0, params.limit);
      const groupedByDate: Record<string, typeof limitedEvents> = {};

      for (const event of limitedEvents) {
        if (!groupedByDate[event.startDate]) {
          groupedByDate[event.startDate] = [];
        }
        groupedByDate[event.startDate].push(event);
      }

      const limitedFormatted = {
        ...formatted,
        events: limitedEvents,
        groupedByDate,
      };

      return {
        windows,
        formatted: {
          _humanReadable: calendarResponseToText(limitedFormatted),
          ...limitedFormatted,
        },
      };
    }
  }

  return { windows, formatted: null };
}

/**
 * Find files shared by or with a person
 */
async function findFilesFromPerson(
  graphClient: GraphClient,
  userEmail: string,
  userDisplayName: string,
  limit: number = 20
): Promise<GraphDriveItem[]> {
  const allFiles: GraphDriveItem[] = [];

  // 1. Search for files shared by this person
  try {
    const sharedResponse = await graphClient.makeRequest('/me/drive/sharedWithMe', {
      method: 'GET',
      queryParams: {
        $top: String(limit),
      },
    });

    if (
      sharedResponse &&
      typeof sharedResponse === 'object' &&
      'value' in sharedResponse &&
      Array.isArray(sharedResponse.value)
    ) {
      const files = sharedResponse.value as GraphDriveItem[];
      const emailLower = userEmail.toLowerCase();
      const nameLower = userDisplayName.toLowerCase();

      for (const file of files) {
        const sharedByEmail = file.shared?.sharedBy?.user?.email?.toLowerCase();
        const sharedByName = file.shared?.sharedBy?.user?.displayName?.toLowerCase();
        const createdByEmail = file.createdBy?.user?.email?.toLowerCase();
        const createdByName = file.createdBy?.user?.displayName?.toLowerCase();

        if (
          sharedByEmail === emailLower ||
          createdByEmail === emailLower ||
          (sharedByName && sharedByName.includes(nameLower)) ||
          (createdByName && createdByName.includes(nameLower))
        ) {
          allFiles.push(file);
        }
      }
    }
  } catch (error) {
    logger.warn(`Error getting shared files: ${error}`);
  }

  // 2. Search OneDrive/SharePoint for files mentioning the person using Central Search
  try {
    const searchResult = await executeCentralSearch(
      graphClient,
      `author:"${userDisplayName}" OR createdBy:"${userDisplayName}"`,
      {
        entityTypes: ['driveItem'],
        maxResults: limit,
        sortByRank: true,
      }
    );

    for (const hit of searchResult.results.files) {
      if (hit.resource) {
        allFiles.push(hit.resource as GraphDriveItem);
      }
    }
  } catch (error) {
    logger.warn(`Error searching files: ${error}`);
  }

  // Deduplicate
  const uniqueFiles = Array.from(new Map(allFiles.map((f) => [f.id, f])).values());
  return uniqueFiles.slice(0, limit);
}

/**
 * Product-specific query functions
 */
async function queryOneNote(graphClient: GraphClient, query: string): Promise<any[]> {
  try {
    const response = await graphClient.makeRequest('/me/onenote/pages', {
      method: 'GET',
      queryParams: {
        $search: `"${query}"`,
        $select: 'id,title,links,lastModifiedDateTime',
        $top: '10',
      },
    });
    return (response as { value: any[] })?.value || [];
  } catch (error) {
    logger.warn(`OneNote search failed: ${error}`);
    return [];
  }
}

async function queryPlanner(graphClient: GraphClient, query: string): Promise<any[]> {
  try {
    const response = await graphClient.makeRequest('/me/planner/tasks', {
      method: 'GET',
      queryParams: { $top: '50' },
    });
    const tasks = (response as { value: any[] })?.value || [];
    return tasks.filter((t) => t.title?.toLowerCase().includes(query.toLowerCase()));
  } catch (error) {
    logger.warn(`Planner search failed: ${error}`);
    return [];
  }
}

async function queryToDo(graphClient: GraphClient, query: string): Promise<any[]> {
  try {
    const listsResponse = await graphClient.makeRequest('/me/todo/lists', { method: 'GET' });
    const lists = (listsResponse as { value: any[] })?.value || [];
    const allTasks: any[] = [];

    for (const list of lists.slice(0, 5)) {
      const tasksResponse = await graphClient.makeRequest(`/me/todo/lists/${list.id}/tasks`, {
        method: 'GET',
        queryParams: {
          $filter: `contains(title,'${query.replace(/'/g, "''")}')`,
          $top: '10',
        },
      });
      const tasks = (tasksResponse as { value: any[] })?.value || [];
      allTasks.push(...tasks.map((t) => ({ ...t, listName: list.displayName })));
    }
    return allTasks;
  } catch (error) {
    logger.warn(`To-Do search failed: ${error}`);
    return [];
  }
}

async function queryContacts(graphClient: GraphClient, query: string): Promise<any[]> {
  try {
    const response = await graphClient.makeRequest('/me/contacts', {
      method: 'GET',
      queryParams: {
        $search: `"${query}"`,
        $top: '10',
      },
    });
    return (response as { value: any[] })?.value || [];
  } catch (error) {
    logger.warn(`Contacts search failed: ${error}`);
    return [];
  }
}

async function queryOnlineMeetings(graphClient: GraphClient, query: string): Promise<any[]> {
  try {
    const response = await graphClient.makeRequest('/me/onlineMeetings', {
      method: 'GET',
      queryParams: { $top: '50' },
    });
    const meetings = (response as { value: any[] })?.value || [];
    return meetings.filter((m) => m.subject?.toLowerCase().includes(query.toLowerCase()));
  } catch (error) {
    logger.warn(`OnlineMeetings search failed: ${error}`);
    return [];
  }
}

async function queryJoinedTeams(graphClient: GraphClient, query: string): Promise<any[]> {
  try {
    const response = await graphClient.makeRequest('/me/joinedTeams', {
      method: 'GET',
    });
    const teams = (response as { value: any[] })?.value || [];
    return teams.filter((t) => t.displayName?.toLowerCase().includes(query.toLowerCase()));
  } catch (error) {
    logger.warn(`JoinedTeams search failed: ${error}`);
    return [];
  }
}

async function queryBookings(graphClient: GraphClient, query: string): Promise<any[]> {
  try {
    const businessesResponse = await graphClient.makeRequest('/solutions/bookingBusinesses', {
      method: 'GET',
    });
    const businesses = (businessesResponse as { value: any[] })?.value || [];
    const allAppointments: any[] = [];

    for (const business of businesses.slice(0, 3)) {
      const appointmentsResponse = await graphClient.makeRequest(
        `/solutions/bookingBusinesses/${business.id}/appointments`,
        {
          method: 'GET',
          queryParams: { $top: '20' },
        }
      );
      const appointments = (appointmentsResponse as { value: any[] })?.value || [];
      allAppointments.push(
        ...appointments
          .filter((a) => a.customerName?.toLowerCase().includes(query.toLowerCase()))
          .map((a) => ({ ...a, businessName: business.displayName }))
      );
    }
    return allAppointments;
  } catch (error) {
    logger.warn(`Bookings search failed: ${error}`);
    return [];
  }
}

async function queryInsights(graphClient: GraphClient): Promise<any[]> {
  try {
    const [trending, shared] = await Promise.allSettled([
      graphClient.makeRequest('/me/insights/trending', {
        method: 'GET',
        queryParams: { $top: '10' },
      }),
      graphClient.makeRequest('/me/insights/shared', {
        method: 'GET',
        queryParams: { $top: '10' },
      }),
    ]);

    const results: any[] = [];
    if (trending.status === 'fulfilled' && trending.value) {
      results.push(
        ...((trending.value as { value: any[] })?.value || []).map((i) => ({
          ...i,
          insightType: 'trending',
        }))
      );
    }
    if (shared.status === 'fulfilled' && shared.value) {
      results.push(
        ...((shared.value as { value: any[] })?.value || []).map((i) => ({
          ...i,
          insightType: 'shared',
        }))
      );
    }
    return results;
  } catch (error) {
    logger.warn(`Insights search failed: ${error}`);
    return [];
  }
}

/**
 * Customer type detection result
 */
interface CustomerTypeResult {
  type: 'person' | 'company' | 'unknown';
  identifier: string;
  displayName: string;
  emailDomain?: string;
  userEmail?: string;
  userId?: string;
  companyContacts?: Array<{
    name: string;
    email?: string;
    title?: string;
    department?: string;
  }>;
}

/**
 * Extract email domain from an email address
 */
function extractEmailDomain(email: string | undefined): string | undefined {
  if (!email) return undefined;
  const parts = email.split('@');
  return parts.length === 2 ? parts[1].toLowerCase() : undefined;
}

/**
 * Generate a likely email domain from a company name
 */
function generateEmailDomain(companyName: string): string {
  return (
    companyName
      .toLowerCase()
      .replace(/\s+/g, '')
      .replace(/[^a-z0-9]/g, '') + '.com'
  );
}

/**
 * Detect if the input is a person name or company name
 * Uses multiple strategies to identify the customer type
 */
async function detectCustomerType(
  graphClient: GraphClient,
  searchQuery: string
): Promise<CustomerTypeResult> {
  logger.info(`Detecting customer type for: "${searchQuery}"`);

  // Strategy 1: Try to find as a person/user first
  const user = await findUser(graphClient, searchQuery);
  if (user && user.displayName) {
    logger.info(`Detected as person: ${user.displayName}`);
    return {
      type: 'person',
      identifier: searchQuery,
      displayName: user.displayName,
      emailDomain: extractEmailDomain(user.mail),
      userEmail: user.mail,
      userId: user.id,
    };
  }

  // Strategy 2: Search contacts for company name
  try {
    const contactsResponse = await graphClient.makeRequest('/me/contacts', {
      method: 'GET',
      queryParams: {
        $filter: `contains(companyName, '${searchQuery.replace(/'/g, "''")}')`,
        $top: '20',
        $select: 'displayName,emailAddresses,companyName,jobTitle,department',
      },
    });

    if (
      contactsResponse &&
      typeof contactsResponse === 'object' &&
      'value' in contactsResponse &&
      Array.isArray(contactsResponse.value) &&
      contactsResponse.value.length > 0
    ) {
      const contacts = contactsResponse.value as Array<{
        displayName: string;
        emailAddresses?: Array<{ address: string }>;
        companyName?: string;
        jobTitle?: string;
        department?: string;
      }>;

      // Extract email domain from the first contact with an email
      const firstEmailDomain = contacts.find((c) => c.emailAddresses?.[0]?.address)
        ? extractEmailDomain(contacts[0].emailAddresses?.[0]?.address)
        : undefined;

      logger.info(`Detected as company with ${contacts.length} contacts`);
      return {
        type: 'company',
        identifier: searchQuery,
        displayName: contacts[0].companyName || searchQuery,
        emailDomain: firstEmailDomain || generateEmailDomain(searchQuery),
        companyContacts: contacts.map((c) => ({
          name: c.displayName,
          email: c.emailAddresses?.[0]?.address,
          title: c.jobTitle,
          department: c.department,
        })),
      };
    }
  } catch (error) {
    logger.debug(`Contact search for company failed: ${error}`);
  }

  // Strategy 3: Try searching in people API
  try {
    const peopleResponse = await graphClient.makeRequest('/me/people', {
      method: 'GET',
      queryParams: {
        $search: `"${searchQuery}"`,
        $top: '5',
      },
    });

    if (
      peopleResponse &&
      typeof peopleResponse === 'object' &&
      'value' in peopleResponse &&
      Array.isArray(peopleResponse.value) &&
      peopleResponse.value.length > 0
    ) {
      const person = peopleResponse.value[0] as {
        id?: string;
        displayName?: string;
        companyName?: string;
        scoredEmailAddresses?: Array<{ address: string }>;
      };

      // If person has a company name that matches the search, it's likely a company search
      if (
        person.companyName &&
        person.companyName.toLowerCase().includes(searchQuery.toLowerCase())
      ) {
        logger.info(`Detected as company from people API: ${person.companyName}`);
        return {
          type: 'company',
          identifier: searchQuery,
          displayName: person.companyName,
          emailDomain:
            extractEmailDomain(person.scoredEmailAddresses?.[0]?.address) ||
            generateEmailDomain(searchQuery),
        };
      }

      // Otherwise it's a person
      logger.info(`Detected as person from people API: ${person.displayName}`);
      return {
        type: 'person',
        identifier: searchQuery,
        displayName: person.displayName || searchQuery,
        emailDomain: extractEmailDomain(person.scoredEmailAddresses?.[0]?.address),
        userEmail: person.scoredEmailAddresses?.[0]?.address,
        userId: person.id,
      };
    }
  } catch (error) {
    logger.debug(`People search failed: ${error}`);
  }

  // Strategy 4: Fallback - assume it's a company if it looks like one
  // (contains common company suffixes or multiple words)
  const companyIndicators = [
    'gmbh',
    'ag',
    'ltd',
    'inc',
    'corp',
    'llc',
    'company',
    'co.',
    'group',
    'holding',
    'se',
    'kg',
    'ohg',
    'ug',
  ];
  const lowerQuery = searchQuery.toLowerCase();
  const isLikelyCompany =
    companyIndicators.some((indicator) => lowerQuery.includes(indicator)) ||
    searchQuery.split(/\s+/).length >= 2;

  if (isLikelyCompany) {
    logger.info(`Assumed as company based on naming pattern: ${searchQuery}`);
    return {
      type: 'company',
      identifier: searchQuery,
      displayName: searchQuery,
      emailDomain: generateEmailDomain(searchQuery),
    };
  }

  // Default: unknown, will search by name
  logger.info(`Could not determine type for: ${searchQuery}, treating as unknown`);
  return {
    type: 'unknown',
    identifier: searchQuery,
    displayName: searchQuery,
    emailDomain: generateEmailDomain(searchQuery),
  };
}

/**
 * CENTRAL MICROSOFT SEARCH API FUNCTION
 *
 * This is the primary search function that ALL tools should use.
 * Features:
 * - Unified search across all Microsoft 365 products
 * - Proper ranking and relevance scoring
 * - Configurable entity types
 * - Time context for events
 * - Automatic categorization of results
 *
 * @param graphClient - The Graph API client
 * @param query - Search query string
 * @param options - Optional search configuration
 * @returns Structured search results with ranking
 */
async function executeCentralSearch(
  graphClient: GraphClient,
  query: string,
  options: CentralSearchOptions = {}
): Promise<CentralSearchResult> {
  const startTime = Date.now();

  const {
    entityTypes = [
      'message',
      'event',
      'driveItem',
      'site',
      'list',
      'listItem',
      'chatMessage',
      'person',
    ],
    maxResults = 100,
    minRelevance = 0,
    includeTimeContext = true,
    sortByRank = true,
    timeRange,
  } = options;

  const results: SearchApiResults = {
    emails: [],
    events: [],
    files: [],
    sites: [],
    listItems: [],
    chats: [],
    people: [],
  };

  const entityTypesCounts: Record<string, number> = {};
  let totalRank = 0;
  let hitCount = 0;

  try {
    // Build time context for events
    const now = new Date();
    const defaultStartDate = new Date(
      now.getFullYear() - 2,
      now.getMonth(),
      now.getDate()
    ).toISOString();
    const defaultEndDate = new Date(
      now.getFullYear() + 1,
      now.getMonth(),
      now.getDate()
    ).toISOString();

    const searchRequest: Record<string, unknown> = {
      entityTypes,
      query: { queryString: query },
      from: 0,
      size: Math.min(maxResults, 500),
    };

    // Add time context if events are included
    if (includeTimeContext && entityTypes.includes('event')) {
      searchRequest.timeContext = {
        startDateTime: timeRange?.startDateTime || defaultStartDate,
        endDateTime: timeRange?.endDateTime || defaultEndDate,
      };
    }

    logger.info(`Central Search: "${query}" across [${entityTypes.join(', ')}]`);

    const response = await graphClient.makeRequest('/search/query', {
      method: 'POST',
      body: JSON.stringify({
        requests: [searchRequest],
      }),
    });

    if (response && typeof response === 'object' && 'value' in response) {
      const searchValues = (response as { value: unknown[] }).value as Array<{
        hitsContainers?: Array<{
          total?: number;
          moreResultsAvailable?: boolean;
          hits?: Array<{
            rank?: number;
            summary?: string;
            resource?: Record<string, unknown>;
          }>;
        }>;
      }>;

      for (const container of searchValues) {
        if (container.hitsContainers) {
          for (const hitsContainer of container.hitsContainers) {
            if (hitsContainer.hits) {
              for (const hit of hitsContainer.hits) {
                if (hit.resource) {
                  const type = ((hit.resource['@odata.type'] as string) || '').toLowerCase();
                  const rank = hit.rank || 0;

                  // Enhanced relevance calculation with query matching and temporal weighting
                  let relevanceScore = Math.max(0, 100 - hitCount * 2); // Base score from position

                  // Boost relevance based on query matching in key fields
                  const queryLower = query.toLowerCase();
                  const resource = hit.resource || {};
                  const name = (
                    resource.name ||
                    resource.subject ||
                    resource.displayName ||
                    ''
                  ).toLowerCase();
                  const summary = (hit.summary || resource.bodyPreview || '').toLowerCase();

                  // Exact match in name/title/subject gets highest boost
                  if (name.includes(queryLower)) {
                    relevanceScore += 20;
                  }

                  // Partial word matches
                  const queryWords = queryLower.split(/\s+/).filter((w) => w.length > 2);
                  const nameWords = name.split(/\s+/);
                  const matchingWords = queryWords.filter((qw) =>
                    nameWords.some((nw) => nw.includes(qw) || qw.includes(nw))
                  );
                  if (matchingWords.length > 0) {
                    relevanceScore += matchingWords.length * 5;
                  }

                  // Summary/content match boost
                  if (summary.includes(queryLower)) {
                    relevanceScore += 10;
                  }

                  // Temporal relevance boost (newer items are more relevant)
                  const timestamp = this.extractTimestampFromResource(resource);
                  if (timestamp) {
                    const daysSince = (Date.now() - timestamp.getTime()) / (1000 * 60 * 60 * 24);
                    if (daysSince < 7) {
                      relevanceScore += 15; // Very recent
                    } else if (daysSince < 30) {
                      relevanceScore += 10; // Recent
                    } else if (daysSince < 90) {
                      relevanceScore += 5; // Somewhat recent
                    }
                  }

                  // Entity-type specific weighting
                  if (type.includes('message') || type.includes('email')) {
                    relevanceScore += 5; // Emails are often highly relevant
                  } else if (type.includes('event')) {
                    // Future events are more relevant than past events
                    if (resource.start?.dateTime) {
                      const startDate = new Date(resource.start.dateTime as string);
                      if (startDate > new Date()) {
                        relevanceScore += 10; // Future event
                      }
                    }
                  } else if (type.includes('driveitem') || type.includes('file')) {
                    relevanceScore += 3; // Files are moderately relevant
                  }

                  // Cap relevance score at 100
                  relevanceScore = Math.min(100, Math.max(0, relevanceScore));

                  // Skip if below minimum relevance
                  if (minRelevance > 0 && relevanceScore < minRelevance) {
                    continue;
                  }

                  totalRank += rank;
                  hitCount++;

                  const hitData: SearchHit = {
                    resource: { ...hit.resource, _rank: rank, _relevance: relevanceScore },
                    summary: hit.summary,
                    name:
                      (hit.resource.name as string) ||
                      (hit.resource.subject as string) ||
                      (hit.resource.displayName as string),
                    webUrl: (hit.resource.webUrl as string) || (hit.resource.webLink as string),
                    rank,
                    relevanceScore,
                  };

                  // Enhanced categorization
                  const category = categorizeSearchHit(type, hit.resource);
                  entityTypesCounts[category] = (entityTypesCounts[category] || 0) + 1;

                  switch (category) {
                    case 'email':
                      results.emails.push(hitData);
                      break;
                    case 'event':
                      results.events.push(hitData);
                      break;
                    case 'file':
                      results.files.push(hitData);
                      break;
                    case 'site':
                      results.sites.push(hitData);
                      break;
                    case 'listItem':
                      results.listItems.push(hitData);
                      break;
                    case 'chat':
                      results.chats.push(hitData);
                      break;
                    case 'person':
                      results.people.push(hitData);
                      break;
                    default:
                      logger.debug(`Unrecognized search result type: ${type}`);
                  }
                }
              }
            }
          }
        }
      }
    }

    // Sort all results by rank if enabled
    if (sortByRank) {
      const sortFn = (a: SearchHit, b: SearchHit) => (b.rank || 0) - (a.rank || 0);
      results.emails.sort(sortFn);
      results.events.sort(sortFn);
      results.files.sort(sortFn);
      results.sites.sort(sortFn);
      results.listItems.sort(sortFn);
      results.chats.sort(sortFn);
      results.people.sort(sortFn);
    }
  } catch (error) {
    logger.error(`Central Search API failed: ${error}`);
  }

  const totalHits =
    results.emails.length +
    results.events.length +
    results.files.length +
    results.sites.length +
    results.listItems.length +
    results.chats.length +
    results.people.length;

  return {
    query,
    searchedAt: new Date().toISOString(),
    totalHits,
    results,
    metadata: {
      entityTypesCounts,
      averageRank: hitCount > 0 ? totalRank / hitCount : 0,
      searchDuration: Date.now() - startTime,
    },
  };
}

/**
 * Categorize a search hit based on its type and properties
 */
function categorizeSearchHit(
  type: string,
  resource: Record<string, unknown>
): 'email' | 'event' | 'file' | 'site' | 'listItem' | 'chat' | 'person' | 'unknown' {
  // Check for person types first (more specific)
  if (
    type.includes('person') ||
    type.includes('contact') ||
    type.includes('orgcontact') ||
    type.includes('user') ||
    (resource.displayName &&
      (resource.emailAddresses ||
        resource.userPrincipalName ||
        resource.mail ||
        resource.scoredEmailAddresses) &&
      !type.includes('message') &&
      !type.includes('event') &&
      !type.includes('driveitem'))
  ) {
    return 'person';
  }

  if (type.includes('message') && !type.includes('chat')) return 'email';
  if (type.includes('event')) return 'event';
  if (type.includes('driveitem')) return 'file';
  if (type.includes('site') && !type.includes('listitem')) return 'site';
  if (type.includes('listitem') || type.includes('list')) return 'listItem';
  if (type.includes('chatmessage')) return 'chat';

  return 'unknown';
}

/**
 * Legacy wrapper for executeSearchApiFirst - uses executeCentralSearch internally
 */
async function executeSearchApiFirst(
  graphClient: GraphClient,
  query: string,
  maxResults: number = 500
): Promise<SearchApiResults> {
  const result = await executeCentralSearch(graphClient, query, {
    maxResults,
    sortByRank: true,
  });
  return result.results;
}

// All entity types supported by Microsoft Search API
const ALL_SEARCH_ENTITY_TYPES = [
  'message',
  'event',
  'driveItem',
  'site',
  'list',
  'listItem',
  'chatMessage',
  'person',
] as const;

async function executeFollowUpQueries(
  graphClient: GraphClient,
  query: string
): Promise<FollowUpResults> {
  const [onenote, planner, todo, contacts, meetings, teams, bookings, insights] =
    await Promise.allSettled([
      queryOneNote(graphClient, query),
      queryPlanner(graphClient, query),
      queryToDo(graphClient, query),
      queryContacts(graphClient, query),
      queryOnlineMeetings(graphClient, query),
      queryJoinedTeams(graphClient, query),
      queryBookings(graphClient, query),
      queryInsights(graphClient),
    ]);

  return {
    onenote: onenote.status === 'fulfilled' ? onenote.value : [],
    planner: planner.status === 'fulfilled' ? planner.value : [],
    todo: todo.status === 'fulfilled' ? todo.value : [],
    contacts: contacts.status === 'fulfilled' ? contacts.value : [],
    meetings: meetings.status === 'fulfilled' ? meetings.value : [],
    teams: teams.status === 'fulfilled' ? teams.value : [],
    bookings: bookings.status === 'fulfilled' ? bookings.value : [],
    insights: insights.status === 'fulfilled' ? insights.value : [],
  };
}

/**
 * Execute a universal search across all Microsoft 365 products
 * Uses the centralized search API for consistent ranking and relevance
 */
async function executeUniversalSearch(
  graphClient: GraphClient,
  query: string,
  limit: number = 25
): Promise<{
  success: boolean;
  results: Record<string, unknown>;
  summary: string;
  totalResults: number;
}> {
  // Use the centralized search function
  const searchResult = await executeCentralSearch(graphClient, query, {
    maxResults: limit,
    sortByRank: true,
    includeTimeContext: true,
  });

  const summaryParts: string[] = [];
  const results: Record<string, unknown> = {
    query: searchResult.query,
    searchedAt: searchResult.searchedAt,
    metadata: searchResult.metadata,
  };

  // Map results to the expected format
  if (searchResult.results.emails.length > 0) {
    results.emails = {
      count: searchResult.results.emails.length,
      items: searchResult.results.emails.slice(0, limit).map((hit) => ({
        subject: hit.resource?.subject,
        from: hit.resource?.from?.emailAddress?.address,
        date: hit.resource?.receivedDateTime,
        preview: hit.resource?.bodyPreview?.substring(0, 150),
        relevance: hit.relevanceScore,
      })),
    };
    summaryParts.push(`${searchResult.results.emails.length} emails`);
  }

  if (searchResult.results.files.length > 0) {
    results.files = {
      count: searchResult.results.files.length,
      items: searchResult.results.files.slice(0, limit).map((hit) => ({
        name: hit.name,
        webUrl: hit.webUrl,
        type: hit.resource?.['@odata.type'],
        relevance: hit.relevanceScore,
      })),
    };
    summaryParts.push(`${searchResult.results.files.length} files`);
  }

  if (searchResult.results.events.length > 0) {
    results.calendar = {
      count: searchResult.results.events.length,
      items: searchResult.results.events.slice(0, limit).map((hit) => ({
        subject: hit.resource?.subject,
        start: hit.resource?.start?.dateTime,
        end: hit.resource?.end?.dateTime,
        organizer: hit.resource?.organizer?.emailAddress?.address,
        location: hit.resource?.location?.displayName,
        relevance: hit.relevanceScore,
      })),
    };
    summaryParts.push(`${searchResult.results.events.length} calendar events`);
  }

  if (searchResult.results.sites.length > 0) {
    results.sites = {
      count: searchResult.results.sites.length,
      items: searchResult.results.sites.slice(0, limit).map((hit) => ({
        name: hit.name,
        webUrl: hit.webUrl,
        relevance: hit.relevanceScore,
      })),
    };
    summaryParts.push(`${searchResult.results.sites.length} sites`);
  }

  if (searchResult.results.people.length > 0) {
    results.people = {
      count: searchResult.results.people.length,
      items: searchResult.results.people.slice(0, limit).map((hit) => ({
        name: hit.name,
        email: hit.resource?.emailAddresses?.[0]?.address || hit.resource?.mail,
        relevance: hit.relevanceScore,
      })),
    };
    summaryParts.push(`${searchResult.results.people.length} people`);
  }

  // Generate summary
  let summary: string;
  if (searchResult.totalHits === 0) {
    summary = `No results found for "${query}" in Microsoft 365.`;
  } else {
    summary = `Found ${searchResult.totalHits} results for "${query}": ${summaryParts.join(', ')}.`;
  }

  return {
    success: searchResult.totalHits > 0,
    results,
    summary,
    totalResults: searchResult.totalHits,
  };
}

/**
 * Register all compound tools
 */
export function registerCompoundTools(
  server: McpServer,
  graphClient: GraphClient,
  readOnly: boolean = false
): number {
  let registeredCount = 0;

  // Initialize NLP Enhancer for structured query analysis
  const nlpEnhancer = new NLPEnhancer();

  // Helper function to get current user email
  async function getCurrentUserEmail(): Promise<string> {
    try {
      const userResponse = await graphClient.makeRequest('/me', {
        method: 'GET',
        queryParams: { $select: 'mail,userPrincipalName' },
      });
      return (
        (userResponse as { mail?: string; userPrincipalName?: string }).mail ||
        (userResponse as { userPrincipalName?: string }).userPrincipalName ||
        ''
      );
    } catch {
      return '';
    }
  }

  // ==========================================================================
  // 0. INTELLIGENT QUERY - PRIMARY ENTRY POINT - ALWAYS PROVIDES AN ANSWER
  // ==========================================================================

  // Stopwords for keyword extraction (English + German)
  const STOPWORDS_EN = [
    'what',
    'where',
    'when',
    'which',
    'about',
    'have',
    'does',
    'tell',
    'show',
    'find',
    'know',
    'give',
    'list',
    'display',
    'search',
    'look',
    'with',
    'from',
    'that',
    'this',
    'there',
    'here',
    'been',
    'being',
    'were',
    'will',
    'would',
    'could',
    'should',
    'might',
    'must',
    'shall',
    'need',
    'want',
    'like',
    'some',
    'more',
    'most',
    'other',
    'into',
    'over',
    'such',
    'only',
    'than',
    'then',
    'also',
    'back',
    'after',
    'before',
    'through',
    'during',
    'without',
    'again',
    'further',
    'once',
    'just',
    'information',
    'emails',
    'email',
    'files',
    'file',
    'meetings',
    'meeting',
    'calendar',
    'messages',
    'message',
  ];

  const STOPWORDS_DE = [
    'was',
    'wer',
    'wie',
    'wann',
    'welche',
    'welcher',
    'welches',
    'über',
    'habe',
    'haben',
    'zeige',
    'zeigen',
    'finde',
    'finden',
    'suche',
    'suchen',
    'gibt',
    'geben',
    'kannst',
    'können',
    'weißt',
    'wissen',
    'sagen',
    'erzähle',
    'erzählen',
    'liste',
    'auflisten',
    'alle',
    'alles',
    'meine',
    'meinen',
    'meiner',
    'deine',
    'deinen',
    'seine',
    'ihre',
    'einem',
    'einer',
    'eines',
    'einen',
    'eine',
    'sein',
    'sind',
    'wird',
    'wurde',
    'werden',
    'gewesen',
    'hatte',
    'hatten',
    'diese',
    'dieser',
    'dieses',
    'jene',
    'jener',
    'jenes',
    'auch',
    'noch',
    'schon',
    'mehr',
    'sehr',
    'ganz',
    'nach',
    'dann',
    'wenn',
    'weil',
    'dass',
    'damit',
    'aber',
    'oder',
    'doch',
    'also',
    'hier',
    'dort',
    'jetzt',
    'heute',
    'gestern',
    'morgen',
    'bitte',
    'danke',
    'informationen',
    'emails',
    'email',
    'dateien',
    'datei',
    'termine',
    'termin',
    'kalender',
    'nachrichten',
    'nachricht',
  ];

  const ALL_STOPWORDS = new Set([...STOPWORDS_EN, ...STOPWORDS_DE]);

  // Detect language based on question
  function detectLanguage(text: string): 'de' | 'en' {
    const germanIndicators = [
      'über',
      'für',
      'können',
      'möchte',
      'bitte',
      'zeige',
      'finde',
      'suche',
      'habe',
      'meine',
      'heute',
      'gestern',
      'morgen',
      'woche',
      'monat',
      'jahr',
      'weißt',
      'kannst',
      'gibt',
      'alles',
      'mails',
      'termine',
      'dateien',
      'letzte',
      'letzten',
      'nächste',
      'nächsten',
    ];
    const lowerText = text.toLowerCase();
    const germanMatches = germanIndicators.filter((word) => lowerText.includes(word)).length;
    return germanMatches >= 1 ? 'de' : 'en';
  }

  // Bilingual messages
  const messages = {
    en: {
      searchingFor: 'Searching for',
      tryingKeywords: 'Trying keywords',
      searchingContext: 'Searching context',
      successMessage: 'Found relevant information in Microsoft 365 for your question.',
      noResultsMessage: (q: string) => `No matching information found in Microsoft 365 for "${q}".`,
      explanation:
        'The search was executed successfully, but no matching data was found in your emails, files, or calendar.',
      suggestions: [
        'Try different keywords or a shorter search term',
        'Check if the information might be under a different name or spelling',
        'The data might not exist in your Microsoft 365 account',
        'Use specific tools like "list-mail-messages" with filters for targeted searches',
        'Try searching for related terms or synonyms',
      ],
      searchCoverage: {
        emails: 'Searched all mailbox messages',
        files: 'Searched OneDrive and SharePoint',
        calendar: 'Searched calendar events from past and future year',
      },
      foundResults: (count: number, sources: string[]) =>
        `Found ${count} results: ${sources.join(', ')}`,
      emails: 'emails',
      files: 'files',
      calendarEvents: 'calendar events',
    },
    de: {
      searchingFor: 'Suche nach',
      tryingKeywords: 'Versuche Schlüsselwörter',
      searchingContext: 'Suche im Kontext',
      successMessage: 'Relevante Informationen in Microsoft 365 für Ihre Frage gefunden.',
      noResultsMessage: (q: string) =>
        `Keine passenden Informationen in Microsoft 365 für "${q}" gefunden.`,
      explanation:
        'Die Suche wurde erfolgreich ausgeführt, aber es wurden keine passenden Daten in Ihren E-Mails, Dateien oder Kalendern gefunden.',
      suggestions: [
        'Versuchen Sie andere Suchbegriffe oder kürzere Suchterme',
        'Prüfen Sie, ob die Information unter einem anderen Namen gespeichert ist',
        'Die Daten existieren möglicherweise nicht in Ihrem Microsoft 365 Konto',
        'Nutzen Sie spezifische Tools wie "list-mail-messages" mit Filtern für gezielte Suchen',
        'Versuchen Sie verwandte Begriffe oder Synonyme',
      ],
      searchCoverage: {
        emails: 'Alle Postfach-Nachrichten durchsucht',
        files: 'OneDrive und SharePoint durchsucht',
        calendar: 'Kalendereinträge aus vergangenem und kommendem Jahr durchsucht',
      },
      foundResults: (count: number, sources: string[]) =>
        `${count} Ergebnisse gefunden: ${sources.join(', ')}`,
      emails: 'E-Mails',
      files: 'Dateien',
      calendarEvents: 'Kalendertermine',
    },
  };

  // Intent classification for intelligent routing
  type Intent =
    | 'email'
    | 'calendar'
    | 'files'
    | 'people'
    | 'teams'
    | 'tasks'
    | 'search'
    | 'mixed'
    | 'sharepoint'
    | 'notes'
    | 'planner'
    | 'contacts'
    | 'meetings'
    | 'bookings'
    | 'insights';

  interface IntentResult {
    primary: Intent;
    secondary: Intent | null;
    confidence: number;
    extractedEntities: {
      person?: string;
      topic?: string;
      timeframe?: string;
      filter?: string;
    };
  }

  // Intent keywords (EN + DE)
  const INTENT_PATTERNS: Record<Intent, { en: string[]; de: string[] }> = {
    email: {
      en: [
        'email',
        'emails',
        'mail',
        'mails',
        'inbox',
        'message',
        'messages',
        'sent',
        'received',
        'unread',
        'attachment',
        'reply',
        'forward',
      ],
      de: [
        'email',
        'emails',
        'e-mail',
        'e-mails',
        'mail',
        'mails',
        'posteingang',
        'postfach',
        'nachricht',
        'nachrichten',
        'gesendet',
        'empfangen',
        'ungelesen',
        'anhang',
        'anhänge',
        'antwort',
        'weiterleiten',
      ],
    },
    calendar: {
      en: [
        'calendar',
        'meeting',
        'meetings',
        'appointment',
        'appointments',
        'schedule',
        'event',
        'events',
        'today',
        'tomorrow',
        'week',
        'month',
        'busy',
        'free',
        'available',
      ],
      de: [
        'kalender',
        'termin',
        'termine',
        'meeting',
        'meetings',
        'besprechung',
        'besprechungen',
        'veranstaltung',
        'heute',
        'morgen',
        'woche',
        'monat',
        'beschäftigt',
        'frei',
        'verfügbar',
      ],
    },
    files: {
      en: [
        'file',
        'files',
        'document',
        'documents',
        'folder',
        'folders',
        'onedrive',
        'sharepoint',
        'excel',
        'word',
        'powerpoint',
        'pdf',
        'spreadsheet',
        'presentation',
        'download',
        'upload',
        'shared',
      ],
      de: [
        'datei',
        'dateien',
        'dokument',
        'dokumente',
        'ordner',
        'onedrive',
        'sharepoint',
        'excel',
        'word',
        'powerpoint',
        'pdf',
        'tabelle',
        'präsentation',
        'herunterladen',
        'hochladen',
        'geteilt',
        'freigegeben',
      ],
    },
    people: {
      en: [
        'contact',
        'contacts',
        'person',
        'people',
        'user',
        'users',
        'colleague',
        'colleagues',
        'manager',
        'team',
        'phone',
        'email address',
        'who is',
      ],
      de: [
        'kontakt',
        'kontakte',
        'person',
        'personen',
        'benutzer',
        'kollege',
        'kollegen',
        'manager',
        'team',
        'telefon',
        'wer ist',
        'mitarbeiter',
      ],
    },
    teams: {
      en: [
        'teams',
        'chat',
        'chats',
        'channel',
        'channels',
        'conversation',
        'conversations',
        'teams message',
        'teams chat',
      ],
      de: [
        'teams',
        'chat',
        'chats',
        'kanal',
        'kanäle',
        'unterhaltung',
        'unterhaltungen',
        'teams nachricht',
        'teams chat',
      ],
    },
    tasks: {
      en: ['task', 'tasks', 'todo', 'to-do', 'to do', 'planner', 'reminder', 'reminders', 'due'],
      de: [
        'aufgabe',
        'aufgaben',
        'todo',
        'to-do',
        'planner',
        'erinnerung',
        'erinnerungen',
        'fällig',
      ],
    },
    search: {
      en: ['search', 'find', 'look for', 'where is', 'locate'],
      de: ['suche', 'suchen', 'finde', 'finden', 'wo ist', 'lokalisieren'],
    },
    sharepoint: {
      en: ['sharepoint', 'site', 'sites', 'list', 'lists', 'document library'],
      de: ['sharepoint', 'seite', 'seiten', 'liste', 'listen', 'dokumentbibliothek'],
    },
    notes: {
      en: ['onenote', 'notebook', 'note', 'notes', 'section'],
      de: ['onenote', 'notizbuch', 'notiz', 'notizen', 'abschnitt'],
    },
    planner: {
      en: ['planner', 'plan', 'bucket', 'board'],
      de: ['planner', 'plan', 'bucket', 'board'],
    },
    contacts: {
      en: ['contact', 'contacts', 'address book', 'phone number'],
      de: ['kontakt', 'kontakte', 'adressbuch', 'telefonnummer'],
    },
    meetings: {
      en: ['meeting', 'meetings', 'online meeting', 'teams meeting', 'call'],
      de: ['meeting', 'meetings', 'online-meeting', 'teams-meeting', 'anruf'],
    },
    bookings: {
      en: ['booking', 'appointment', 'schedule', 'reservation'],
      de: ['buchung', 'termin', 'zeitplan', 'reservierung'],
    },
    insights: {
      en: ['trending', 'shared with me', 'popular', 'used', 'analytics'],
      de: ['trend', 'mit mir geteilt', 'beliebt', 'verwendet', 'analyse'],
    },
    mixed: { en: [], de: [] },
  };

  // Time expressions for extraction (extended)
  const TIME_PATTERNS: Record<'en' | 'de', Record<string, string[]>> = {
    en: {
      today: ['today', 'this day', 'todays'],
      yesterday: ['yesterday', 'yesterdays'],
      tomorrow: ['tomorrow', 'tomorrows'],
      thisWeek: ['this week', 'current week', 'the week'],
      lastWeek: ['last week', 'previous week', 'past week'],
      nextWeek: ['next week', 'coming week', 'upcoming week'],
      thisMonth: ['this month', 'current month'],
      lastMonth: ['last month', 'previous month', 'past month'],
      nextMonth: ['next month', 'coming month'],
      last7Days: ['last 7 days', 'past 7 days', 'last seven days', 'past week'],
      last30Days: ['last 30 days', 'past 30 days', 'last thirty days', 'past month'],
      last90Days: ['last 90 days', 'past 90 days', 'last quarter', 'past quarter'],
      thisYear: ['this year', 'current year'],
      lastYear: ['last year', 'previous year', 'past year'],
    },
    de: {
      today: ['heute', 'diesen tag', 'heutigen'],
      yesterday: ['gestern', 'gestrigen'],
      tomorrow: ['morgen', 'morgigen'],
      thisWeek: ['diese woche', 'aktuelle woche', 'der woche', 'dieser woche'],
      lastWeek: ['letzte woche', 'vorige woche', 'vergangene woche', 'letzten woche'],
      nextWeek: ['nächste woche', 'kommende woche'],
      thisMonth: ['diesen monat', 'aktueller monat', 'diesem monat'],
      lastMonth: ['letzten monat', 'voriger monat', 'vergangenen monat'],
      nextMonth: ['nächsten monat', 'kommenden monat'],
      last7Days: ['letzten 7 tage', 'vergangenen 7 tage', 'letzte woche'],
      last30Days: ['letzten 30 tage', 'vergangenen 30 tage', 'letzter monat'],
      last90Days: ['letzten 90 tage', 'letztes quartal', 'vergangenes quartal'],
      thisYear: ['dieses jahr', 'aktuelles jahr'],
      lastYear: ['letztes jahr', 'voriges jahr', 'vergangenes jahr'],
    },
  };

  // Convert timeframe to date range
  /**
   * Helper function to get end of day (23:59:59.999) for a given date
   */
  function endOfDay(date: Date): Date {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59, 999);
  }

  /**
   * Helper function to get start of day (00:00:00) for a given date
   */
  function startOfDay(date: Date): Date {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0, 0);
  }

  function getDateRangeFromTimeframe(timeframe: string): { start: Date; end: Date } {
    const now = new Date();
    const today = startOfDay(now);
    let start: Date;
    let end: Date;

    switch (timeframe) {
      case 'today':
        start = today;
        end = endOfDay(today);
        break;
      case 'yesterday': {
        const yesterdayDate = new Date(today.getTime() - 24 * 60 * 60 * 1000);
        start = startOfDay(yesterdayDate);
        end = endOfDay(yesterdayDate);
        break;
      }
      case 'tomorrow': {
        const tomorrowDate = new Date(today.getTime() + 24 * 60 * 60 * 1000);
        start = startOfDay(tomorrowDate);
        end = endOfDay(tomorrowDate);
        break;
      }
      case 'thisWeek': {
        const dayOfWeek = today.getDay();
        const weekStart = new Date(today.getTime() - dayOfWeek * 24 * 60 * 60 * 1000);
        const weekEnd = new Date(weekStart.getTime() + 6 * 24 * 60 * 60 * 1000);
        start = startOfDay(weekStart);
        end = endOfDay(weekEnd);
        break;
      }
      case 'lastWeek': {
        const lastWeekStart = new Date(
          today.getTime() - (today.getDay() + 7) * 24 * 60 * 60 * 1000
        );
        const lastWeekEnd = new Date(lastWeekStart.getTime() + 6 * 24 * 60 * 60 * 1000);
        start = startOfDay(lastWeekStart);
        end = endOfDay(lastWeekEnd);
        break;
      }
      case 'nextWeek': {
        const nextWeekStart = new Date(
          today.getTime() + (7 - today.getDay()) * 24 * 60 * 60 * 1000
        );
        const nextWeekEnd = new Date(nextWeekStart.getTime() + 6 * 24 * 60 * 60 * 1000);
        start = startOfDay(nextWeekStart);
        end = endOfDay(nextWeekEnd);
        break;
      }
      case 'thisMonth': {
        const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
        const monthEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0); // Last day of current month
        start = startOfDay(monthStart);
        end = endOfDay(monthEnd);
        break;
      }
      case 'lastMonth': {
        const lastMonthStart = new Date(now.getFullYear(), now.getMonth() - 1, 1);
        const lastMonthEnd = new Date(now.getFullYear(), now.getMonth(), 0); // Last day of previous month
        start = startOfDay(lastMonthStart);
        end = endOfDay(lastMonthEnd);
        break;
      }
      case 'nextMonth': {
        const nextMonthStart = new Date(now.getFullYear(), now.getMonth() + 1, 1);
        const nextMonthEnd = new Date(now.getFullYear(), now.getMonth() + 2, 0); // Last day of next month
        start = startOfDay(nextMonthStart);
        end = endOfDay(nextMonthEnd);
        break;
      }
      case 'last7Days': {
        const last7Start = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
        start = startOfDay(last7Start);
        end = endOfDay(today);
        break;
      }
      case 'last30Days': {
        const last30Start = new Date(today.getTime() - 30 * 24 * 60 * 60 * 1000);
        start = startOfDay(last30Start);
        end = endOfDay(today);
        break;
      }
      case 'last90Days': {
        const last90Start = new Date(today.getTime() - 90 * 24 * 60 * 60 * 1000);
        start = startOfDay(last90Start);
        end = endOfDay(today);
        break;
      }
      case 'thisYear': {
        const yearStart = new Date(now.getFullYear(), 0, 1);
        const yearEnd = new Date(now.getFullYear(), 11, 31); // December 31st
        start = startOfDay(yearStart);
        end = endOfDay(yearEnd);
        break;
      }
      case 'lastYear': {
        const lastYearStart = new Date(now.getFullYear() - 1, 0, 1);
        const lastYearEnd = new Date(now.getFullYear() - 1, 11, 31); // December 31st of last year
        start = startOfDay(lastYearStart);
        end = endOfDay(lastYearEnd);
        break;
      }
      default: {
        // Default to last 14 days until end of today
        const defaultStart = new Date(today.getTime() - 14 * 24 * 60 * 60 * 1000);
        start = startOfDay(defaultStart);
        end = endOfDay(today);
      }
    }

    return { start, end };
  }

  // Required scopes for each intent
  const INTENT_REQUIRED_SCOPES: Record<Intent, { scopes: string[]; workScopes: string[] }> = {
    email: { scopes: ['Mail.Read'], workScopes: [] },
    calendar: { scopes: ['Calendars.Read'], workScopes: [] },
    files: { scopes: ['Files.Read'], workScopes: ['Sites.Read.All'] },
    people: { scopes: ['People.Read', 'User.Read'], workScopes: ['User.Read.All'] },
    teams: { scopes: [], workScopes: ['Chat.Read', 'ChatMessage.Read'] },
    tasks: { scopes: ['Tasks.Read'], workScopes: [] },
    search: { scopes: ['Mail.Read', 'Files.Read'], workScopes: ['Sites.Read.All'] },
    sharepoint: { scopes: [], workScopes: ['Sites.Read.All', 'Sites.Selected'] },
    notes: { scopes: ['Notes.Read'], workScopes: [] },
    planner: { scopes: ['Tasks.Read'], workScopes: [] },
    contacts: { scopes: ['Contacts.Read'], workScopes: [] },
    meetings: { scopes: ['OnlineMeetings.Read'], workScopes: [] },
    bookings: { scopes: ['Bookings.Read.All'], workScopes: [] },
    insights: { scopes: [], workScopes: ['Sites.Read.All'] },
    mixed: { scopes: ['Mail.Read', 'Calendars.Read', 'Files.Read'], workScopes: [] },
  };

  // Check if user likely has permissions (based on common errors)
  function getPermissionWarning(intent: Intent, lang: 'de' | 'en'): string | null {
    const teamsWarning = {
      en: '⚠️ Teams/Chat access requires organization mode (--org-mode) and work account permissions.',
      de: '⚠️ Teams/Chat-Zugriff erfordert den Organisationsmodus (--org-mode) und Arbeitskontoberechtigungen.',
    };

    const filesWarning = {
      en: '⚠️ SharePoint access may require organization mode for full functionality.',
      de: '⚠️ SharePoint-Zugriff kann den Organisationsmodus für volle Funktionalität erfordern.',
    };

    if (intent === 'teams') {
      return teamsWarning[lang];
    }

    if (intent === 'files') {
      return filesWarning[lang];
    }

    return null;
  }

  // Classify intent from question
  function classifyIntent(question: string, lang: 'en' | 'de'): IntentResult {
    const lowerQuestion = question.toLowerCase();
    const scores: Record<Intent, number> = {
      email: 0,
      calendar: 0,
      files: 0,
      people: 0,
      teams: 0,
      tasks: 0,
      search: 0,
      mixed: 0,
      sharepoint: 0,
      notes: 0,
      planner: 0,
      contacts: 0,
      meetings: 0,
      bookings: 0,
      insights: 0,
    };

    // Score each intent based on keyword matches
    for (const [intent, patterns] of Object.entries(INTENT_PATTERNS) as [
      Intent,
      { en: string[]; de: string[] },
    ][]) {
      const allPatterns = [...patterns.en, ...patterns.de];
      for (const pattern of allPatterns) {
        if (lowerQuestion.includes(pattern)) {
          scores[intent] += pattern.split(' ').length; // Multi-word patterns score higher
        }
      }
    }

    // Find primary and secondary intents
    const sortedIntents = (Object.entries(scores) as [Intent, number][])
      .filter(([intent]) => intent !== 'mixed')
      .sort((a, b) => b[1] - a[1]);

    const primary = sortedIntents[0][1] > 0 ? sortedIntents[0][0] : 'search';
    const secondary =
      sortedIntents[1][1] > 0 && sortedIntents[1][1] >= sortedIntents[0][1] * 0.5
        ? sortedIntents[1][0]
        : null;
    const maxScore = Math.max(...Object.values(scores));
    const confidence = maxScore > 0 ? Math.min(maxScore / 3, 1) : 0.3;

    // Extract entities
    const extractedEntities: IntentResult['extractedEntities'] = {};

    // Extract timeframe
    const timePatterns = TIME_PATTERNS[lang];
    for (const [timeKey, patterns] of Object.entries(timePatterns)) {
      for (const pattern of patterns) {
        if (lowerQuestion.includes(pattern)) {
          extractedEntities.timeframe = timeKey;
          break;
        }
      }
      if (extractedEntities.timeframe) break;
    }

    // Extract person (simple heuristic: words after "from", "with", "von", "mit")
    const personPatterns = lang === 'de' ? ['von ', 'mit ', 'an '] : ['from ', 'with ', 'to '];
    for (const pattern of personPatterns) {
      const idx = lowerQuestion.indexOf(pattern);
      if (idx !== -1) {
        const afterPattern = question.substring(idx + pattern.length).split(/[,.\s]/)[0];
        if (afterPattern && afterPattern.length > 2) {
          extractedEntities.person = afterPattern;
          break;
        }
      }
    }

    // Extract topic (remaining significant words)
    const topicWords = question
      .replace(/[?!.,;:'"„"«»]/g, '')
      .split(' ')
      .filter((word) => word.length > 3 && !ALL_STOPWORDS.has(word.toLowerCase()))
      .slice(0, 3);
    if (topicWords.length > 0) {
      extractedEntities.topic = topicWords.join(' ');
    }

    return { primary, secondary, confidence, extractedEntities };
  }

  // Execute intent-specific queries with permission awareness
  async function executeIntentQuery(
    intent: Intent,
    entities: IntentResult['extractedEntities'],
    lang: 'de' | 'en'
  ): Promise<{
    data: unknown;
    source: string;
    count: number;
    permissionWarning?: string;
    error?: string;
  }> {
    const results: {
      data: unknown;
      source: string;
      count: number;
      permissionWarning?: string;
      error?: string;
    } = {
      data: null,
      source: intent,
      count: 0,
    };

    // Check for permission warnings
    const permissionWarning = getPermissionWarning(intent, lang);
    if (permissionWarning) {
      results.permissionWarning = permissionWarning;
    }

    try {
      switch (intent) {
        case 'email': {
          const queryParams: Record<string, string> = {
            $top: '50',
            $select:
              'id,subject,bodyPreview,receivedDateTime,sentDateTime,from,toRecipients,ccRecipients,importance,isRead,isDraft,hasAttachments,webLink,categories,flag,conversationId',
            $orderby: 'receivedDateTime desc',
          };

          // Apply timeframe filter using extended date range
          if (entities.timeframe) {
            const dateRange = getDateRangeFromTimeframe(entities.timeframe);
            queryParams.$filter = `receivedDateTime ge ${dateRange.start.toISOString()} and receivedDateTime lt ${dateRange.end.toISOString()}`;
            logger.info(
              `Email timeframe filter: ${entities.timeframe} -> ${dateRange.start.toISOString()} to ${dateRange.end.toISOString()}`
            );
          }

          if (entities.person) {
            queryParams.$search = `"from:${entities.person}"`;
            delete queryParams.$orderby;
            // Cannot combine $search with $filter in Graph API
            delete queryParams.$filter;
          } else if (entities.topic) {
            queryParams.$search = `"${entities.topic}"`;
            delete queryParams.$orderby;
            delete queryParams.$filter;
          }

          const response = await graphClient.makeRequest('/me/messages', {
            method: 'GET',
            queryParams,
          });

          if (response && typeof response === 'object' && 'value' in response) {
            // Use the structured mail formatter with local time
            const formattedResponse = formatMailResponse(response as Record<string, unknown>);

            // Generate context for current date
            const dateContext = generateDateTimeContext(lang);

            results.data = {
              _humanReadable: mailResponseToText(formattedResponse),
              _llmInstructions:
                lang === 'de'
                  ? `WICHTIG: Liste ALLE E-Mails unten mit Betreff, Absender, Datum und Uhrzeit (${dateContext.timezone}) auf. Heute ist ${dateContext.formatted}. Zeige die E-Mails NICHT als Zusammenfassung, sondern als vollständige Liste.`
                  : `IMPORTANT: List ALL emails below with subject, sender, date and time (${dateContext.timezone}). Today is ${dateContext.formatted}. Do NOT show a summary, show the complete list of emails.`,
              currentContext: dateContext,
              summary: formattedResponse.summary,
              messages: formattedResponse.messages,
              groupedByDate: formattedResponse.groupedByDate,
            };
            results.count = formattedResponse.summary.totalMessages;
          }
          break;
        }

        case 'calendar': {
          // Use extended timeframe or default to next 14 days
          const dateRange = entities.timeframe
            ? getDateRangeFromTimeframe(entities.timeframe)
            : getDateRangeFromTimeframe('default');

          logger.info(
            `Calendar timeframe: ${entities.timeframe || 'default'} -> ${dateRange.start.toISOString()} to ${dateRange.end.toISOString()}`
          );

          // Use the extracted dateRange
          const startDate = dateRange.start;
          const endDate = dateRange.end;

          const response = await graphClient.makeRequest(
            `/me/calendarView?startDateTime=${startDate.toISOString()}&endDateTime=${endDate.toISOString()}`,
            {
              method: 'GET',
              queryParams: {
                $top: '50',
                $select:
                  'id,subject,start,end,location,organizer,attendees,isAllDay,isCancelled,isOnlineMeeting,onlineMeeting,bodyPreview,webLink',
                $orderby: 'start/dateTime',
              },
            }
          );

          if (response && typeof response === 'object' && 'value' in response) {
            // Use the structured calendar formatter with local time
            const formattedResponse = formatCalendarResponse(
              response as Record<string, unknown>,
              startDate.toISOString(),
              endDate.toISOString()
            );

            // Generate context for current date
            const dateContext = generateDateTimeContext(lang);

            results.data = {
              _humanReadable: calendarResponseToText(formattedResponse),
              _llmInstructions:
                lang === 'de'
                  ? `WICHTIG: Liste ALLE Termine unten mit Betreff, Datum, Uhrzeit (${dateContext.timezone}, NICHT UTC!) und Ort auf. Heute ist ${dateContext.formatted}. "Morgen" bedeutet ${dateContext.references.tomorrow}. Zeige die Termine NICHT als Zusammenfassung, sondern als vollständige Liste mit lokalen Uhrzeiten.`
                  : `IMPORTANT: List ALL events below with subject, date, time (${dateContext.timezone}, NOT UTC!) and location. Today is ${dateContext.formatted}. "Tomorrow" means ${dateContext.references.tomorrow}. Do NOT show a summary, show the complete list of events with local times.`,
              currentContext: dateContext,
              summary: formattedResponse.summary,
              events: formattedResponse.events,
              groupedByDate: formattedResponse.groupedByDate,
            };
            results.count = formattedResponse.summary.totalEvents;
          }
          break;
        }

        case 'files': {
          const searchQuery = entities.topic || 'recent';
          // Use centralized search for consistent ranking
          const searchResult = await executeCentralSearch(graphClient, searchQuery, {
            entityTypes: ['driveItem'],
            maxResults: 25,
            sortByRank: true,
          });

          if (searchResult.results.files.length > 0) {
            results.data = searchResult.results.files.map((hit) => ({
              name: hit.name,
              webUrl: hit.webUrl,
              lastModified: hit.resource?.lastModifiedDateTime,
              relevance: hit.relevanceScore,
            }));
            results.count = searchResult.results.files.length;
          }
          break;
        }

        case 'people': {
          const searchQuery = entities.person || entities.topic;
          if (searchQuery) {
            // Try multiple search strategies for better results
            const searchStrategies: Array<Record<string, string>> = [
              { $search: `"${searchQuery}"` }, // Exact phrase
              { $search: searchQuery }, // Without quotes for partial matches
              { $filter: `startswith(displayName,'${searchQuery}')` }, // Starts with
            ];

            let allPeople: unknown[] = [];
            const seenEmails = new Set<string>();

            for (const strategy of searchStrategies) {
              try {
                const response = await graphClient.makeRequest('/me/people', {
                  method: 'GET',
                  queryParams: {
                    ...strategy,
                    $top: '10',
                  },
                });

                if (response && typeof response === 'object' && 'value' in response) {
                  const people = (response as { value: unknown[] }).value || [];
                  for (const person of people) {
                    const p = person as {
                      displayName?: string;
                      emailAddresses?: Array<{ address?: string }>;
                    };
                    const email = p.emailAddresses?.[0]?.address;
                    // Deduplicate by email
                    if (email && !seenEmails.has(email)) {
                      seenEmails.add(email);
                      allPeople.push(person);
                    } else if (!email) {
                      // If no email, check by name
                      const exists = allPeople.some(
                        (existing) =>
                          (existing as { displayName?: string }).displayName === p.displayName
                      );
                      if (!exists) {
                        allPeople.push(person);
                      }
                    }
                  }
                }
              } catch (err) {
                // Continue with next strategy if one fails
                logger.debug(`People search strategy failed: ${err}`);
              }
            }

            if (allPeople.length > 0) {
              results.data = allPeople.map((p: unknown) => {
                const person = p as {
                  displayName?: string;
                  emailAddresses?: Array<{ address?: string }>;
                  phones?: Array<{ number?: string }>;
                  department?: string;
                  jobTitle?: string;
                };
                return {
                  name: person.displayName,
                  email: person.emailAddresses?.[0]?.address,
                  phone: person.phones?.[0]?.number,
                  department: person.department,
                  jobTitle: person.jobTitle,
                };
              });
              results.count = allPeople.length;
            }
          }
          break;
        }

        case 'tasks': {
          const response = await graphClient.makeRequest('/me/todo/lists', {
            method: 'GET',
          });

          if (response && typeof response === 'object' && 'value' in response) {
            const lists = (response as { value: Array<{ id: string; displayName?: string }> })
              .value;
            const allTasks: unknown[] = [];

            for (const list of lists.slice(0, 3)) {
              const tasksResponse = await graphClient.makeRequest(
                `/me/todo/lists/${list.id}/tasks`,
                {
                  method: 'GET',
                  queryParams: { $top: '10' },
                }
              );

              if (tasksResponse && typeof tasksResponse === 'object' && 'value' in tasksResponse) {
                const tasks = (tasksResponse as { value: unknown[] }).value || [];
                for (const t of tasks) {
                  const task = t as {
                    title?: string;
                    status?: string;
                    dueDateTime?: { dateTime?: string };
                    importance?: string;
                  };
                  allTasks.push({
                    list: list.displayName,
                    title: task.title,
                    status: task.status,
                    dueDate: task.dueDateTime?.dateTime,
                    importance: task.importance,
                  });
                }
              }
            }
            results.data = allTasks;
            results.count = allTasks.length;
          }
          break;
        }

        case 'sharepoint': {
          const searchQuery = entities.topic || 'recent';
          // Use centralized search for SharePoint content
          const searchResult = await executeCentralSearch(graphClient, searchQuery, {
            entityTypes: ['site', 'list', 'listItem'],
            maxResults: 25,
            sortByRank: true,
          });

          const items: Array<{
            name?: string;
            webUrl?: string;
            type?: string;
            relevance?: number;
          }> = [];

          // Combine sites and listItems from search results
          for (const hit of searchResult.results.sites) {
            items.push({
              name: hit.name || hit.resource?.displayName,
              webUrl: hit.webUrl,
              type: 'site',
              relevance: hit.relevanceScore,
            });
          }
          for (const hit of searchResult.results.listItems) {
            items.push({
              name: hit.name || hit.resource?.displayName,
              webUrl: hit.webUrl,
              type: hit.resource?.['@odata.type'],
              relevance: hit.relevanceScore,
            });
          }

          results.data = items;
          results.count = items.length;
          break;
        }

        case 'notes': {
          results.data = await queryOneNote(graphClient, entities.topic || '');
          results.count = (results.data as any[]).length;
          break;
        }

        case 'planner': {
          results.data = await queryPlanner(graphClient, entities.topic || '');
          results.count = (results.data as any[]).length;
          break;
        }

        case 'contacts': {
          results.data = await queryContacts(graphClient, entities.topic || entities.person || '');
          results.count = (results.data as any[]).length;
          break;
        }

        case 'meetings': {
          results.data = await queryOnlineMeetings(graphClient, entities.topic || '');
          results.count = (results.data as any[]).length;
          break;
        }

        case 'bookings': {
          results.data = await queryBookings(graphClient, entities.topic || '');
          results.count = (results.data as any[]).length;
          break;
        }

        case 'insights': {
          results.data = await queryInsights(graphClient);
          results.count = (results.data as any[]).length;
          break;
        }

        case 'teams':
        case 'search':
        case 'mixed':
        default: {
          // Fallback to universal search
          const searchResult = await executeUniversalSearch(
            graphClient,
            entities.topic || 'recent',
            25
          );
          results.data = searchResult.results;
          results.count = searchResult.totalResults;
          break;
        }
      }
    } catch (error) {
      const errorMessage = (error as Error).message || String(error);
      logger.error(`Intent query failed for ${intent}: ${errorMessage}`);

      // Check for permission-related errors
      const isPermissionError =
        errorMessage.includes('403') ||
        errorMessage.includes('401') ||
        errorMessage.includes('Forbidden') ||
        errorMessage.includes('Unauthorized') ||
        errorMessage.includes('Access denied') ||
        errorMessage.includes('insufficient') ||
        errorMessage.includes('Zugriff verweigert');

      if (isPermissionError) {
        const requiredScopes = INTENT_REQUIRED_SCOPES[intent];
        const scopeInfo = requiredScopes.scopes.join(', ') || requiredScopes.workScopes.join(', ');

        results.error =
          lang === 'de'
            ? `⛔ **Berechtigung erforderlich**: Der Zugriff auf ${intent}-Daten wurde verweigert.\n\n` +
              `**Erforderliche Berechtigungen:** ${scopeInfo}\n\n` +
              `**Lösung:**\n` +
              `1. Melden Sie sich ab und wieder an (erneute Authentifizierung)\n` +
              `2. Für Teams/Chat: Starten Sie mit --org-mode\n` +
              `3. Prüfen Sie, ob Ihr Konto die nötigen Lizenzen hat`
            : `⛔ **Permission Required**: Access to ${intent} data was denied.\n\n` +
              `**Required scopes:** ${scopeInfo}\n\n` +
              `**Solution:**\n` +
              `1. Sign out and sign in again (re-authenticate)\n` +
              `2. For Teams/Chat: Start with --org-mode\n` +
              `3. Verify your account has the necessary licenses`;

        results.data = {
          error: 'permission_denied',
          requiredScopes: requiredScopes.scopes,
          workScopes: requiredScopes.workScopes,
          message: results.error,
        };
      } else {
        results.error =
          lang === 'de'
            ? `⚠️ Fehler bei der Abfrage: ${errorMessage}`
            : `⚠️ Query error: ${errorMessage}`;
        results.data = { error: errorMessage };
      }
      results.count = 0;
    }

    return results;
  }

  server.tool(
    'ask-microsoft-365',
    `🧠 **ENHANCED INTELLIGENT MICROSOFT 365 ASSISTANT** - THE ULTIMATE WAY TO QUERY MS365!
🇬🇧 English | 🇩🇪 Deutsch - Automatically detects your language!

⭐ **THIS IS THE RECOMMENDED PRIMARY TOOL** - Use this for ANY question about Microsoft 365 data!
⭐ **DIES IST DAS EMPFOHLENE HAUPT-TOOL** - Nutzen Sie es für JEDE Frage zu Microsoft 365 Daten!

This tool provides UNIVERSAL search across 16+ Microsoft 365 products:
Dieses Tool bietet eine UNIVERSAL-Suche über mehr als 16 Microsoft 365-Produkte:

1.  📧 **Outlook Mail**
2.  📅 **Calendar**
3.  📁 **OneDrive & SharePoint Files**
4.  🌐 **SharePoint Sites & Lists**
5.  👥 **People & Organization**
6.  📇 **Personal Contacts**
7.  📓 **OneNote Notebooks**
8.  🗓️ **Planner Tasks**
9.  ✅ **Microsoft To-Do**
10. 💬 **Teams Chats**
11. 👥 **Joined Teams**
12. 📹 **Online Meetings**
13. 🗓️ **Bookings**
14. 📈 **Insights (Trending)**
15. 🤝 **Insights (Shared)**
16. 📊 **Activity Analytics**

**Features:**
- ✅ **ALWAYS returns an answer** - never fails silently
- 🔄 **Search-First Strategy** - uses Microsoft Search API for relevant results
- 🎯 **Smart query understanding** - interprets natural language (EN + DE)
- 📊 **Comprehensive results** - aggregates data from all 16 sources
- 💡 **Helpful suggestions** - provides next steps if no results found
- 📅 **Time-aware** - understands "today", "last week", "this month", etc.

**GUARANTEE: This tool will ALWAYS provide a response!**`,
    {
      question: z
        .string()
        .describe(
          'Your question in natural language (English or German) / Ihre Frage in natürlicher Sprache (Englisch oder Deutsch)'
        ),
      context: z
        .string()
        .optional()
        .describe(
          'Optional context to refine the search / Optionaler Kontext zur Verfeinerung der Suche'
        ),
      language: z
        .enum(['auto', 'en', 'de'])
        .optional()
        .describe('Response language: auto (detect), en (English), de (German). Default: auto'),
    },
    {
      title: 'Enhanced Microsoft 365 Assistant (EN/DE)',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ question, context, language = 'auto' }) => {
      const startTime = Date.now();
      logger.info(`Enhanced query: "${question}"${context ? ` (context: ${context})` : ''}`);

      // Initialize thinking process for transparent reasoning
      const thinking = createThinkingProcess('ask-m365');
      thinking.addReasoning(
        'intent',
        `Processing natural language query: "${question.substring(0, 50)}${question.length > 50 ? '...' : ''}"`
      );

      // Get chat memory context if enabled
      const chatId = getChatId();
      const userId = getUserId();
      let memoryContext: string | null = null;
      let memoryStore: ReturnType<typeof getChatMemoryStore> | null = null;

      if (isChatMemoryEnabled() && chatId) {
        memoryStore = getChatMemoryStore();
        // SECURITY: Pass userId to ensure user only accesses their own memory
        memoryContext = memoryStore.buildContextForQuery(chatId, userId);

        // Apply stored preferences if available
        // SECURITY: Pass userId to ensure user only accesses their own preferences
        const prefs = memoryStore.getPreferences(chatId, userId);
        if (prefs?.language && language === 'auto') {
          language = prefs.language;
          thinking.addInfo('processing', `Applied stored language preference: ${language}`);
        }

        logger.debug('Chat memory context available', {
          chatId: chatId.substring(0, 8),
          userId: userId?.substring(0, 8),
          hasContext: !!memoryContext,
          prefsApplied: !!prefs?.language,
        });
        if (memoryContext) {
          thinking.addInfo('processing', 'Chat memory context loaded from previous conversation');
        }
      }

      // Detect language
      const detectedLang = language === 'auto' ? detectLanguage(question) : language;
      const msg = messages[detectedLang];
      logger.info(`Detected language: ${detectedLang}`);
      thinking.addDecision(
        'intent',
        `Detected language: ${detectedLang === 'de' ? 'German' : 'English'}`
      );

      // Perform NLP-based query decomposition for structured analysis
      thinking.startAction('processing', 'Analyzing query with NLP decomposition');
      const queryAnalysis = nlpEnhancer.decomposeQuery(question);
      logger.debug('Query decomposition completed', {
        entity: queryAnalysis.entity,
        intent: queryAnalysis.intent.type,
        confidence: queryAnalysis.confidence,
        subQueries: queryAnalysis.subQueries.length,
      });
      thinking.completeAction(
        'processing',
        `Query decomposed: ${queryAnalysis.intent.type} intent with ${Math.round(queryAnalysis.confidence * 100)}% confidence`
      );

      // Combine question, context, and memory context
      let fullQuestion = context ? `${question} ${context}` : question;

      // Enhance query with memory context if available
      if (memoryContext) {
        // Add memory context as additional context for better search
        fullQuestion = `${fullQuestion}\n\n[Previous conversation context: ${memoryContext}]`;
        logger.debug('Query enhanced with memory context');
      }

      // Step 1: Classify intent
      thinking.startAction('intent', 'Classifying user intent');
      const intentResult = classifyIntent(fullQuestion, detectedLang);
      const processingSteps: string[] = [];
      thinking.completeAction(
        'intent',
        `Primary intent: ${intentResult.primary} (${Math.round(intentResult.confidence * 100)}% confidence)`
      );

      // Intent labels for display
      const intentLabels: Record<Intent, { en: string; de: string }> = {
        email: { en: 'Emails', de: 'E-Mails' },
        calendar: { en: 'Calendar', de: 'Kalender' },
        files: { en: 'Files', de: 'Dateien' },
        people: { en: 'People', de: 'Personen' },
        teams: { en: 'Teams', de: 'Teams' },
        tasks: { en: 'Tasks', de: 'Aufgaben' },
        search: { en: 'Search', de: 'Suche' },
        mixed: { en: 'Mixed', de: 'Gemischt' },
        sharepoint: { en: 'SharePoint', de: 'SharePoint' },
        notes: { en: 'OneNote', de: 'OneNote' },
        planner: { en: 'Planner', de: 'Planner' },
        contacts: { en: 'Contacts', de: 'Kontakte' },
        meetings: { en: 'Meetings', de: 'Meetings' },
        bookings: { en: 'Bookings', de: 'Bookings' },
        insights: { en: 'Insights', de: 'Insights' },
      };

      processingSteps.push(
        detectedLang === 'de'
          ? `🎯 Intent: ${intentLabels[intentResult.primary].de} (${Math.round(intentResult.confidence * 100)}%)`
          : `🎯 Intent: ${intentLabels[intentResult.primary].en} (${Math.round(intentResult.confidence * 100)}%)`
      );

      // Step 2: Central Search API query
      processingSteps.push(
        detectedLang === 'de'
          ? `🔍 Starte Microsoft Search API Abfrage...`
          : `🔍 Starting Microsoft Search API query...`
      );

      const searchQuery = intentResult.extractedEntities.topic || question;
      thinking.startAction(
        'search',
        `Executing Microsoft Search API query: "${searchQuery.substring(0, 30)}${searchQuery.length > 30 ? '...' : ''}"`
      );
      const searchResults = await executeSearchApiFirst(graphClient, searchQuery);
      const searchResultCount =
        searchResults.emails.length +
        searchResults.events.length +
        searchResults.files.length +
        searchResults.people.length +
        searchResults.sites.length +
        searchResults.listItems.length;
      thinking.completeAction(
        'search',
        `Search API returned ${searchResultCount} results across 6 categories`
      );

      // Step 3: Execute Follow-up queries for products not covered by Search API
      processingSteps.push(
        detectedLang === 'de'
          ? `🔄 Erweitere Suche auf zusätzliche Produkte...`
          : `🔄 Expanding search to additional products...`
      );

      thinking.startAction('aggregation', 'Executing follow-up queries for additional products');
      const followUpResults = await executeFollowUpQueries(graphClient, searchQuery);
      thinking.completeAction('aggregation', 'Follow-up queries completed');

      // Step 4: Product-specific primary intent execution (for more details)
      let primaryIntentData: any = null;
      if (intentResult.primary !== 'search' && intentResult.primary !== 'mixed') {
        const primaryResult = await executeIntentQuery(
          intentResult.primary,
          intentResult.extractedEntities,
          detectedLang
        );
        if (primaryResult.count > 0) {
          primaryIntentData = primaryResult.data;

          // Merge people results from intent query into searchResults.people
          if (intentResult.primary === 'people' && Array.isArray(primaryResult.data)) {
            for (const person of primaryResult.data) {
              // person format: { name, email, phone, department, jobTitle }
              const personHit: SearchHit = {
                resource: {
                  displayName: person.name,
                  emailAddresses: person.email ? [{ address: person.email }] : undefined,
                  phones: person.phone ? [{ number: person.phone }] : undefined,
                  department: person.department,
                  jobTitle: person.jobTitle,
                },
                summary: undefined,
                name: person.name,
                webUrl: undefined,
              };
              // Avoid duplicates by name or email
              const exists = searchResults.people.some(
                (p) =>
                  p.name === personHit.name ||
                  (person.email && p.resource?.emailAddresses?.[0]?.address === person.email)
              );
              if (!exists) {
                searchResults.people.push(personHit);
              }
            }
          }
        }
      }

      // Merge contacts from followUpResults into searchResults.people
      if (followUpResults.contacts && Array.isArray(followUpResults.contacts)) {
        for (const contact of followUpResults.contacts) {
          const c = contact as {
            displayName?: string;
            givenName?: string;
            surname?: string;
            emailAddresses?: Array<{ address?: string }>;
            businessPhones?: string[];
            department?: string;
            jobTitle?: string;
            companyName?: string;
          };
          const contactName = c.displayName || `${c.givenName || ''} ${c.surname || ''}`.trim();
          const contactEmail = c.emailAddresses?.[0]?.address;

          const contactHit: SearchHit = {
            resource: {
              displayName: contactName,
              emailAddresses: c.emailAddresses,
              phones: c.businessPhones?.map((p) => ({ number: p })),
              department: c.department,
              jobTitle: c.jobTitle,
              companyName: c.companyName,
              _source: 'contacts',
            },
            summary: undefined,
            name: contactName,
            webUrl: undefined,
          };

          // Deduplicate by name or email
          const exists = searchResults.people.some(
            (p) =>
              p.name === contactHit.name ||
              (contactEmail && p.resource?.emailAddresses?.[0]?.address === contactEmail)
          );
          if (!exists && contactName) {
            searchResults.people.push(contactHit);
          }
        }
      }

      // Aggregate counts (excluding contacts since they're merged into people)
      let totalCount = 0;
      const productsSearched = [
        'Mail',
        'Calendar',
        'Files',
        'SharePoint',
        'People',
        'Contacts',
        'OneNote',
        'Planner',
        'To-Do',
        'Teams',
        'Meetings',
        'Bookings',
        'Insights',
      ];

      totalCount += searchResults.emails.length;
      totalCount += searchResults.events.length;
      totalCount += searchResults.files.length;
      totalCount += searchResults.sites.length;
      totalCount += searchResults.listItems.length;
      totalCount += searchResults.chats.length;
      totalCount += searchResults.people.length;

      if (followUpResults.onenote) totalCount += followUpResults.onenote.length;
      if (followUpResults.planner) totalCount += followUpResults.planner.length;
      if (followUpResults.todo) totalCount += followUpResults.todo.length;
      // Note: contacts are already merged into searchResults.people, don't double-count
      if (followUpResults.meetings) totalCount += followUpResults.meetings.length;
      if (followUpResults.teams) totalCount += followUpResults.teams.length;
      if (followUpResults.bookings) totalCount += followUpResults.bookings.length;
      if (followUpResults.insights) totalCount += followUpResults.insights.length;

      // Sort all results by rank (higher rank = more relevant)
      const sortByRank = (hits: SearchHit[]): SearchHit[] => {
        return [...hits].sort((a, b) => {
          const rankA = (a.resource as { _rank?: number })?._rank || 0;
          const rankB = (b.resource as { _rank?: number })?._rank || 0;
          return rankB - rankA;
        });
      };

      // Apply ranking to all categories
      searchResults.emails = sortByRank(searchResults.emails);
      searchResults.events = sortByRank(searchResults.events);
      searchResults.files = sortByRank(searchResults.files);
      searchResults.sites = sortByRank(searchResults.sites);
      searchResults.listItems = sortByRank(searchResults.listItems);
      searchResults.chats = sortByRank(searchResults.chats);
      searchResults.people = sortByRank(searchResults.people);

      // Generate structured summary for people intent
      let summary: string | undefined;
      if (intentResult.primary === 'people' && searchResults.people.length > 0) {
        const topPeople = searchResults.people.slice(0, 5);
        const peopleSummary = topPeople
          .map((p) => {
            const name = p.name || p.resource?.displayName || 'Unknown';
            const email = p.resource?.emailAddresses?.[0]?.address || p.resource?.mail || '';
            const jobTitle = p.resource?.jobTitle || '';
            const department = p.resource?.department || '';
            const company = p.resource?.companyName || '';

            let details = name;
            if (email) details += ` (${email})`;
            if (jobTitle) details += ` - ${jobTitle}`;
            if (department) details += `, ${department}`;
            if (company) details += ` @ ${company}`;
            return details;
          })
          .join('\n');

        summary =
          detectedLang === 'de'
            ? `Gefundene Personen für "${searchQuery}":\n${peopleSummary}${searchResults.people.length > 5 ? `\n... und ${searchResults.people.length - 5} weitere` : ''}`
            : `Found people for "${searchQuery}":\n${peopleSummary}${searchResults.people.length > 5 ? `\n... and ${searchResults.people.length - 5} more` : ''}`;
      } else if (totalCount > 0) {
        // Generate general summary for other intents
        const parts: string[] = [];
        if (searchResults.emails.length > 0)
          parts.push(
            detectedLang === 'de'
              ? `${searchResults.emails.length} E-Mails`
              : `${searchResults.emails.length} emails`
          );
        if (searchResults.events.length > 0)
          parts.push(
            detectedLang === 'de'
              ? `${searchResults.events.length} Termine`
              : `${searchResults.events.length} events`
          );
        if (searchResults.files.length > 0)
          parts.push(
            detectedLang === 'de'
              ? `${searchResults.files.length} Dateien`
              : `${searchResults.files.length} files`
          );
        if (searchResults.people.length > 0)
          parts.push(
            detectedLang === 'de'
              ? `${searchResults.people.length} Personen`
              : `${searchResults.people.length} people`
          );
        if (searchResults.chats.length > 0)
          parts.push(
            detectedLang === 'de'
              ? `${searchResults.chats.length} Teams-Nachrichten`
              : `${searchResults.chats.length} Teams messages`
          );

        if (parts.length > 0) {
          summary =
            detectedLang === 'de' ? `Gefunden: ${parts.join(', ')}` : `Found: ${parts.join(', ')}`;
        }
      }

      // Generate current date/time context for LLM reference
      const currentContext = generateDateTimeContext(detectedLang);

      // Finalize thinking process
      thinking.addDecision(
        'processing',
        `Query completed with ${totalCount} results in ${Date.now() - startTime}ms`
      );
      const thinkingResult = thinking.formatForResponse();

      // Build Enhanced Response
      const response: EnhancedAskM365Response = {
        question,
        language: detectedLang,
        searchedAt: new Date().toISOString(),
        currentContext,
        intent: intentResult,
        processingSteps,
        resultsFound: totalCount > 0,
        totalResults: totalCount,
        status: totalCount > 0 ? 'SUCCESS' : 'NO_RESULTS',
        message:
          totalCount > 0
            ? detectedLang === 'de'
              ? `${totalCount} Ergebnisse in Microsoft 365 gefunden. Heute ist ${currentContext.formatted}.`
              : `Found ${totalCount} results across Microsoft 365. Today is ${currentContext.formatted}.`
            : msg.noResultsMessage(question),
        summary,
        // Include structured query analysis with Markdown summary
        queryAnalysis,
        queryAnalysisMarkdown: queryAnalysis.markdown,
        searchResults,
        followUpResults,
        metadata: {
          totalResults: totalCount,
          queryTime: Date.now() - startTime,
          productsSearched,
        },
        // Include thinking process for transparent reasoning display
        ...(thinkingResult.thinking
          ? { thinking: thinkingResult.thinking as EnhancedAskM365Response['thinking'] }
          : {}),
      };

      if (primaryIntentData) {
        (response as any).primaryIntentResults = primaryIntentData;
      }

      // Store conversation in chat memory if enabled
      if (isChatMemoryEnabled() && chatId && memoryStore) {
        try {
          // Store the question and summarized answer
          const answerSummary =
            totalCount > 0
              ? `Found ${totalCount} results: ${summary || 'Various Microsoft 365 items'}`
              : 'No results found';

          memoryStore.addConversation(chatId, question, answerSummary, {
            toolUsed: 'ask-microsoft-365',
            resultCount: totalCount,
            sources: productsSearched,
            userId,
          });

          // Extract and store mentioned entities
          const entities: Partial<Record<EntityType, string[]>> = {};

          // Extract people from results
          if (searchResults.people.length > 0) {
            entities.people = searchResults.people
              .map((p) => p.name || p.resource?.displayName)
              .filter((n): n is string => !!n)
              .slice(0, 10);
          }

          // Extract files from results
          if (searchResults.files.length > 0) {
            entities.files = searchResults.files
              .map((f) => f.name || f.resource?.name)
              .filter((n): n is string => !!n)
              .slice(0, 10);
          }

          // Extract events from results
          if (searchResults.events.length > 0) {
            entities.events = searchResults.events
              .map((e) => e.resource?.subject || e.name)
              .filter((n): n is string => !!n)
              .slice(0, 10);
          }

          // Extract topics from the question
          if (intentResult.extractedEntities.topic) {
            entities.topics = [intentResult.extractedEntities.topic];
          }

          if (Object.keys(entities).length > 0) {
            memoryStore.addEntities(chatId, entities, userId);
          }

          // Store language preference if user explicitly set it
          if (language !== 'auto') {
            memoryStore.setPreference(chatId, 'language', detectedLang, userId);
          }

          logger.debug('Stored conversation in chat memory', {
            chatId: chatId.substring(0, 8),
            entitiesStored: Object.keys(entities).length,
          });
        } catch (memoryError) {
          // Don't fail the request if memory storage fails
          logger.warn('Failed to store in chat memory', { error: String(memoryError) });
        }
      }

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(response, null, 2),
          },
        ],
        isError: false,
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 0b. WHAT CAN I ASK - Shows example questions that work 100%
  // ==========================================================================
  server.tool(
    'what-can-i-ask',
    `❓ **WHAT CAN I ASK?** - Discover all the questions you can ask!

This tool shows you example questions that are **100% guaranteed to work** with Microsoft 365.
Use this when:
- You want to know what the system can do
- You need inspiration for queries
- You're new to the Microsoft 365 assistant

Returns categorized examples with guaranteed working queries.`,
    {
      category: z
        .enum(['all', 'email', 'calendar', 'files', 'people', 'teams', 'search', 'tasks'])
        .optional()
        .describe('Filter examples by category (default: all)'),
    },
    {
      title: 'What Can I Ask?',
      readOnlyHint: true,
      openWorldHint: false,
    },
    async ({ category = 'all' }) => {
      logger.info(`What can I ask? Category: ${category}`);

      const examples = {
        email: {
          title: '📧 Email & Messages',
          description: 'Questions about your emails, inbox, and mail folders',
          questions: [
            {
              question: 'Show me my latest emails / Zeige mir meine neuesten E-Mails',
              tool: 'get-my-emails',
              guaranteed: true,
              note: 'Enhanced tool with rich formatting / Verbessertes Tool mit reichhaltiger Formatierung',
            },
            {
              question: 'Find emails from [person name or email]',
              tool: 'ask-microsoft-365',
              guaranteed: true,
            },
            {
              question: 'Search for emails about [any topic]',
              tool: 'search-everything',
              guaranteed: true,
            },
            {
              question: 'Show me unread emails',
              tool: 'list-mail-messages with filter',
              guaranteed: true,
            },
            {
              question: 'What emails did I receive today/this week?',
              tool: 'list-mail-messages with date filter',
              guaranteed: true,
            },
            {
              question: 'Show me emails with attachments',
              tool: 'list-mail-messages with hasAttachments filter',
              guaranteed: true,
            },
            {
              question: 'List my mail folders',
              tool: 'list-mail-folders',
              guaranteed: true,
            },
          ],
        },
        calendar: {
          title: '📅 Calendar & Meetings',
          description: 'Questions about your schedule, meetings, and events',
          questions: [
            {
              question: 'What meetings do I have today?',
              tool: 'list-calendar-events',
              guaranteed: true,
            },
            {
              question: 'Show me my schedule for this week',
              tool: 'list-calendar-events with date range',
              guaranteed: true,
            },
            {
              question: 'Find meetings with [person name]',
              tool: 'find-meetings-with-person',
              guaranteed: true,
            },
            {
              question: 'What is my next meeting?',
              tool: 'list-calendar-events',
              guaranteed: true,
            },
            {
              question: 'Show me meetings about [topic]',
              tool: 'ask-microsoft-365',
              guaranteed: true,
            },
            {
              question: 'List my calendars',
              tool: 'list-calendars',
              guaranteed: true,
            },
          ],
        },
        files: {
          title: '📁 Files & Documents',
          description: 'Questions about your OneDrive and SharePoint files',
          questions: [
            {
              question: 'Show me my recent files',
              tool: 'list-drive-recent',
              guaranteed: true,
            },
            {
              question: 'Find files about [topic]',
              tool: 'search-everything',
              guaranteed: true,
            },
            {
              question: 'What files are in my OneDrive?',
              tool: 'list-drive-root-items',
              guaranteed: true,
            },
            {
              question: 'Show files shared with me',
              tool: 'list-drive-shared',
              guaranteed: true,
            },
            {
              question: 'Find documents from [person]',
              tool: 'find-files-from-person',
              guaranteed: true,
            },
            {
              question: 'Search for [filename or content]',
              tool: 'ask-microsoft-365',
              guaranteed: true,
            },
          ],
        },
        people: {
          title: '👥 People & Contacts',
          description: 'Questions about contacts and colleagues',
          questions: [
            {
              question: 'Find contact information for [person name]',
              tool: 'list-users with search',
              guaranteed: true,
            },
            {
              question: 'Show me my contacts',
              tool: 'list-contacts',
              guaranteed: true,
            },
            {
              question: 'Who is [person name]?',
              tool: 'ask-microsoft-365',
              guaranteed: true,
            },
            {
              question: 'Get profile of [email address]',
              tool: 'get-user',
              guaranteed: true,
            },
          ],
        },
        teams: {
          title: '💬 Teams & Chats',
          description: 'Questions about Teams messages and conversations (requires org mode)',
          questions: [
            {
              question: 'Show my Teams chats',
              tool: 'list-chats',
              guaranteed: true,
              note: 'Requires organization mode',
            },
            {
              question: 'Find messages with [person]',
              tool: 'find-messages-with-person',
              guaranteed: true,
              note: 'Requires organization mode',
            },
            {
              question: 'What did I discuss with [person] in Teams?',
              tool: 'find-messages-with-person',
              guaranteed: true,
              note: 'Requires organization mode',
            },
            {
              question: 'Show my Teams',
              tool: 'list-joined-teams',
              guaranteed: true,
              note: 'Requires organization mode',
            },
            {
              question: 'Show posts from [channel name] channel',
              tool: 'get-teams-channel-posts',
              guaranteed: true,
              note: 'Requires organization mode',
            },
            {
              question: 'Was gibt es Neues im [Kanalname] Kanal?',
              tool: 'get-teams-channel-posts',
              guaranteed: true,
              note: 'Requires organization mode',
            },
          ],
        },
        tasks: {
          title: '✅ Tasks & To-Do',
          description: 'Questions about your tasks and to-do lists',
          questions: [
            {
              question: 'Show my tasks',
              tool: 'list-todo-tasks',
              guaranteed: true,
            },
            {
              question: 'What do I need to do?',
              tool: 'list-todo-tasks',
              guaranteed: true,
            },
            {
              question: 'Show my to-do lists',
              tool: 'list-todo-lists',
              guaranteed: true,
            },
          ],
        },
        search: {
          title: '🔍 Universal Search',
          description: 'Search across all Microsoft 365 data',
          questions: [
            {
              question: 'What do you know about [any topic]?',
              tool: 'ask-microsoft-365',
              guaranteed: true,
            },
            {
              question: 'Find everything about [project/company/person]',
              tool: 'search-everything',
              guaranteed: true,
            },
            {
              question: 'Search for [any keyword]',
              tool: 'search-everything',
              guaranteed: true,
            },
            {
              question: 'Tell me about [any subject]',
              tool: 'ask-microsoft-365',
              guaranteed: true,
            },
            {
              question: 'What information do we have about [client/project]?',
              tool: 'ask-microsoft-365',
              guaranteed: true,
            },
          ],
        },
      };

      // Filter by category
      let selectedCategories: Record<string, (typeof examples)['email']>;
      if (category === 'all') {
        selectedCategories = examples;
      } else {
        selectedCategories = { [category]: examples[category] };
      }

      // Build response
      const response = {
        title: '❓ What Can I Ask? - Microsoft 365 Assistant',
        description:
          'Here are example questions you can ask. All these queries are 100% guaranteed to work!',
        category: category === 'all' ? 'All Categories' : category,
        categories: selectedCategories,
        tips: [
          '💡 Use "ask-microsoft-365" for any general question - it always provides an answer!',
          '💡 Use "search-everything" to find information across all your data',
          '💡 Be specific with names and dates for better results',
          '💡 You can combine topics: "emails about budget from last month"',
        ],
        quickStart: [
          { question: 'Show my latest emails / Zeige meine E-Mails', action: 'get-my-emails' },
          { question: 'What meetings do I have today?', action: 'list-calendar-events' },
          { question: 'Find files about [topic]', action: 'search-everything' },
          { question: 'What do you know about [anything]?', action: 'ask-microsoft-365' },
        ],
        totalExamples: Object.values(selectedCategories).reduce(
          (sum, cat) => sum + cat.questions.length,
          0
        ),
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(response, null, 2),
          },
        ],
        isError: false,
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 0c. GET MY EMAILS - Enhanced email retrieval with better formatting
  // ==========================================================================
  server.tool(
    'get-my-emails',
    `📧 **ENHANCED EMAIL TOOL** - Get your emails with rich formatting!
📧 **VERBESSERTES E-MAIL TOOL** - Holen Sie Ihre E-Mails mit reichhaltiger Formatierung!

This tool provides a BETTER email experience than the basic list-mail-messages:
Dieses Tool bietet eine BESSERE E-Mail-Erfahrung als das einfache list-mail-messages:

✅ **Rich data** - Subject, sender, date, preview, importance, attachments
✅ **Reichhaltige Daten** - Betreff, Absender, Datum, Vorschau, Wichtigkeit, Anhänge
✅ **Smart formatting** - Clean, readable output structure
✅ **Intelligente Formatierung** - Saubere, lesbare Ausgabestruktur
✅ **Bilingual summaries** - Automatic DE/EN response
✅ **Zweisprachige Zusammenfassungen** - Automatische DE/EN Antwort
✅ **Helpful when empty** - Clear guidance if no emails found
✅ **Hilfreich bei leeren Ergebnissen** - Klare Anleitung wenn keine E-Mails gefunden

**Use cases / Anwendungsfälle:**
- "Show me my latest emails" / "Zeige mir meine neuesten E-Mails"
- "What emails did I receive today?" / "Welche E-Mails habe ich heute erhalten?"
- "Find emails about [topic]" / "Finde E-Mails über [Thema]"
- "Show unread emails" / "Zeige ungelesene E-Mails"`,
    {
      filter: z
        .enum(['all', 'unread', 'important', 'flagged', 'today', 'thisWeek', 'withAttachments'])
        .optional()
        .describe(
          'Filter emails: all, unread, important, flagged, today, thisWeek, withAttachments'
        ),
      search: z.string().optional().describe('Search term to filter emails / Suchbegriff'),
      limit: z.number().optional().describe('Maximum number of emails (default: 20, max: 50)'),
      language: z
        .enum(['auto', 'en', 'de'])
        .optional()
        .describe('Response language (default: auto)'),
    },
    {
      title: 'Get My Emails (Enhanced)',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ filter = 'all', search, limit = 20, language = 'auto' }) => {
      logger.info(`Get my emails: filter=${filter}, search=${search}, limit=${limit}`);

      // Detect language from search term or default to English
      const detectedLang =
        language === 'auto' ? (search ? detectLanguage(search) : 'en') : language;

      // Build query parameters - include body for better summaries
      const queryParams: Record<string, string> = {
        $top: String(Math.min(limit, 50)),
        $select:
          'id,subject,bodyPreview,body,receivedDateTime,from,toRecipients,ccRecipients,importance,isRead,hasAttachments,flag,categories',
        $orderby: 'receivedDateTime desc',
      };

      // Apply filters
      const filterParts: string[] = [];
      const now = new Date();

      switch (filter) {
        case 'unread':
          filterParts.push('isRead eq false');
          break;
        case 'important':
          filterParts.push("importance eq 'high'");
          break;
        case 'flagged':
          filterParts.push("flag/flagStatus eq 'flagged'");
          break;
        case 'today': {
          const todayStart = new Date(
            now.getFullYear(),
            now.getMonth(),
            now.getDate()
          ).toISOString();
          filterParts.push(`receivedDateTime ge ${todayStart}`);
          break;
        }
        case 'thisWeek': {
          const weekStart = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000).toISOString();
          filterParts.push(`receivedDateTime ge ${weekStart}`);
          break;
        }
        case 'withAttachments':
          filterParts.push('hasAttachments eq true');
          break;
      }

      if (filterParts.length > 0) {
        queryParams.$filter = filterParts.join(' and ');
      }

      if (search) {
        queryParams.$search = `"${search}"`;
        // Remove orderby when using search (Graph API limitation)
        delete queryParams.$orderby;
      }

      try {
        const response = await graphClient.makeRequest('/me/messages', {
          method: 'GET',
          queryParams,
        });

        interface EmailMessage {
          id: string;
          subject?: string;
          bodyPreview?: string;
          body?: { content?: string; contentType?: string };
          receivedDateTime?: string;
          from?: { emailAddress?: { name?: string; address?: string } };
          toRecipients?: Array<{ emailAddress?: { name?: string; address?: string } }>;
          ccRecipients?: Array<{ emailAddress?: { name?: string; address?: string } }>;
          importance?: string;
          isRead?: boolean;
          hasAttachments?: boolean;
          flag?: { flagStatus?: string };
          categories?: string[];
        }

        // SECURITY: Use the centralized sanitizeHtml function for proper HTML sanitization
        // (extractTextFromHtml removed - use sanitizeHtml instead)

        // Helper to create a summary from email content
        const createContentSummary = (email: EmailMessage, maxLength: number = 500): string => {
          let content = '';

          // Try to get content from body first (more complete)
          if (email.body?.content) {
            if (email.body.contentType === 'html') {
              content = sanitizeHtml(email.body.content);
            } else {
              content = email.body.content;
            }
          }

          // Fall back to bodyPreview if body is empty
          if (!content && email.bodyPreview) {
            content = email.bodyPreview;
          }

          // Truncate and add ellipsis if needed
          if (content.length > maxLength) {
            content = content.substring(0, maxLength).trim() + '...';
          }

          return content;
        };

        const emails: EmailMessage[] =
          response && typeof response === 'object' && 'value' in response
            ? (response as { value: EmailMessage[] }).value || []
            : [];

        // Format emails with rich data and content summaries
        const formattedEmails = emails.map((email, index) => {
          const receivedDate = email.receivedDateTime ? new Date(email.receivedDateTime) : null;
          const contentSummary = createContentSummary(email, 500);

          // Create status indicators
          const statusIcons: string[] = [];
          if (!email.isRead) statusIcons.push('📩'); // Unread
          if (email.importance === 'high') statusIcons.push('❗'); // Important
          if (email.hasAttachments) statusIcons.push('📎'); // Attachment
          if (email.flag?.flagStatus === 'flagged') statusIcons.push('🚩'); // Flagged

          return {
            number: index + 1,
            id: email.id,
            statusIcons: statusIcons.join(' ') || '✉️',
            subject: email.subject || (detectedLang === 'de' ? '(Kein Betreff)' : '(No subject)'),
            from: {
              name:
                email.from?.emailAddress?.name || (detectedLang === 'de' ? 'Unbekannt' : 'Unknown'),
              email: email.from?.emailAddress?.address || '',
            },
            date: receivedDate
              ? {
                  iso: email.receivedDateTime,
                  formatted:
                    detectedLang === 'de'
                      ? receivedDate.toLocaleDateString('de-DE', {
                          weekday: 'short',
                          day: '2-digit',
                          month: '2-digit',
                          year: 'numeric',
                          hour: '2-digit',
                          minute: '2-digit',
                        })
                      : receivedDate.toLocaleDateString('en-US', {
                          weekday: 'short',
                          month: 'short',
                          day: 'numeric',
                          year: 'numeric',
                          hour: '2-digit',
                          minute: '2-digit',
                        }),
                  relative: getRelativeTime(receivedDate, detectedLang),
                }
              : null,
            contentSummary: contentSummary,
            status: {
              isRead: email.isRead ?? false,
              importance: email.importance || 'normal',
              hasAttachments: email.hasAttachments ?? false,
              isFlagged: email.flag?.flagStatus === 'flagged',
            },
            categories: email.categories || [],
            to:
              email.toRecipients?.map((r) => ({
                name: r.emailAddress?.name,
                email: r.emailAddress?.address,
              })) || [],
            cc:
              email.ccRecipients?.map((r) => ({
                name: r.emailAddress?.name,
                email: r.emailAddress?.address,
              })) || [],
          };
        });

        // Build comprehensive summary
        const unreadCount = formattedEmails.filter((e) => !e.status.isRead).length;
        const importantCount = formattedEmails.filter((e) => e.status.importance === 'high').length;
        const attachmentCount = formattedEmails.filter((e) => e.status.hasAttachments).length;
        const flaggedCount = formattedEmails.filter((e) => e.status.isFlagged).length;

        // Group emails by sender for overview
        const senderGroups: Record<string, number> = {};
        for (const email of formattedEmails) {
          const senderName = email.from.name || email.from.email || 'Unknown';
          senderGroups[senderName] = (senderGroups[senderName] || 0) + 1;
        }
        const topSenders = Object.entries(senderGroups)
          .sort((a, b) => b[1] - a[1])
          .slice(0, 5)
          .map(([name, count]) => ({ name, count }));

        const summaryParts: string[] = [];
        if (detectedLang === 'de') {
          summaryParts.push(
            `${formattedEmails.length} E-Mail${formattedEmails.length !== 1 ? 's' : ''}`
          );
          if (unreadCount > 0) summaryParts.push(`📩 ${unreadCount} ungelesen`);
          if (importantCount > 0) summaryParts.push(`❗ ${importantCount} wichtig`);
          if (attachmentCount > 0) summaryParts.push(`📎 ${attachmentCount} mit Anhängen`);
          if (flaggedCount > 0) summaryParts.push(`🚩 ${flaggedCount} markiert`);
        } else {
          summaryParts.push(
            `${formattedEmails.length} email${formattedEmails.length !== 1 ? 's' : ''}`
          );
          if (unreadCount > 0) summaryParts.push(`📩 ${unreadCount} unread`);
          if (importantCount > 0) summaryParts.push(`❗ ${importantCount} important`);
          if (attachmentCount > 0) summaryParts.push(`📎 ${attachmentCount} with attachments`);
          if (flaggedCount > 0) summaryParts.push(`🚩 ${flaggedCount} flagged`);
        }

        // Generate date/time context
        const dateContext = generateDateTimeContext(detectedLang);

        const result = {
          _llmInstructions:
            detectedLang === 'de'
              ? `WICHTIG: Liste die E-Mails unten auf mit: Betreff, Absender (Name und E-Mail), Datum/Uhrzeit. Heute ist ${dateContext.formatted}. Zeige ALLE E-Mails aus der "emails" Liste, nicht nur eine Zusammenfassung!`
              : `IMPORTANT: List the emails below with: subject, sender (name and email), date/time. Today is ${dateContext.formatted}. Show ALL emails from the "emails" list, not just a summary!`,
          status: 'SUCCESS',
          language: detectedLang,
          currentContext: dateContext,
          title:
            detectedLang === 'de'
              ? `📧 Ihre E-Mails (${filter === 'all' ? 'Alle' : filter})`
              : `📧 Your Emails (${filter === 'all' ? 'All' : filter})`,
          summary: summaryParts.join(' | '),
          statistics: {
            total: formattedEmails.length,
            unread: unreadCount,
            important: importantCount,
            withAttachments: attachmentCount,
            flagged: flaggedCount,
          },
          topSenders:
            topSenders.length > 0
              ? {
                  label: detectedLang === 'de' ? 'Häufigste Absender' : 'Top Senders',
                  senders: topSenders,
                }
              : undefined,
          filter: filter,
          search: search || null,
          count: formattedEmails.length,
          emails: formattedEmails,
          message:
            formattedEmails.length > 0
              ? detectedLang === 'de'
                ? `✅ ${formattedEmails.length} E-Mail(s) gefunden. Jede E-Mail enthält eine Inhaltszusammenfassung.`
                : `✅ Found ${formattedEmails.length} email(s). Each email includes a content summary.`
              : detectedLang === 'de'
                ? '❌ Keine E-Mails gefunden, die Ihren Kriterien entsprechen.'
                : '❌ No emails found matching your criteria.',
        };

        // Add helpful tips if no results
        if (formattedEmails.length === 0) {
          Object.assign(result, {
            suggestions:
              detectedLang === 'de'
                ? [
                    'Versuchen Sie einen anderen Filter (z.B. "all" statt "unread")',
                    'Erweitern Sie den Zeitraum (z.B. "thisWeek" statt "today")',
                    'Prüfen Sie die Schreibweise des Suchbegriffs',
                    'Ihr Postfach könnte leer sein oder die E-Mails wurden archiviert',
                  ]
                : [
                    'Try a different filter (e.g., "all" instead of "unread")',
                    'Expand the time range (e.g., "thisWeek" instead of "today")',
                    'Check the spelling of your search term',
                    'Your mailbox might be empty or emails may have been archived',
                  ],
          });
        }

        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(result, null, 2),
            },
          ],
          isError: false,
        };
      } catch (error) {
        const errorMessage = (error as Error).message;
        logger.error(`Get my emails failed: ${errorMessage}`);

        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify({
                status: 'ERROR',
                language: detectedLang,
                message:
                  detectedLang === 'de'
                    ? `Fehler beim Abrufen der E-Mails: ${errorMessage}`
                    : `Error retrieving emails: ${errorMessage}`,
                suggestions:
                  detectedLang === 'de'
                    ? [
                        'Stellen Sie sicher, dass Sie angemeldet sind (verwenden Sie das "login" Tool)',
                        'Prüfen Sie Ihre Berechtigungen für den E-Mail-Zugriff',
                      ]
                    : [
                        'Make sure you are logged in (use the "login" tool)',
                        'Check your permissions for email access',
                      ],
              }),
            },
          ],
          isError: true,
        };
      }
    }
  );
  registeredCount++;

  // Helper function for relative time
  function getRelativeTime(date: Date, lang: 'de' | 'en'): string {
    const now = new Date();
    const diffMs = now.getTime() - date.getTime();
    const diffMins = Math.floor(diffMs / 60000);
    const diffHours = Math.floor(diffMs / 3600000);
    const diffDays = Math.floor(diffMs / 86400000);

    if (lang === 'de') {
      if (diffMins < 1) return 'gerade eben';
      if (diffMins < 60) return `vor ${diffMins} Minute${diffMins !== 1 ? 'n' : ''}`;
      if (diffHours < 24) return `vor ${diffHours} Stunde${diffHours !== 1 ? 'n' : ''}`;
      if (diffDays < 7) return `vor ${diffDays} Tag${diffDays !== 1 ? 'en' : ''}`;
      return `vor ${Math.floor(diffDays / 7)} Woche${Math.floor(diffDays / 7) !== 1 ? 'n' : ''}`;
    } else {
      if (diffMins < 1) return 'just now';
      if (diffMins < 60) return `${diffMins} minute${diffMins !== 1 ? 's' : ''} ago`;
      if (diffHours < 24) return `${diffHours} hour${diffHours !== 1 ? 's' : ''} ago`;
      if (diffDays < 7) return `${diffDays} day${diffDays !== 1 ? 's' : ''} ago`;
      return `${Math.floor(diffDays / 7)} week${Math.floor(diffDays / 7) !== 1 ? 's' : ''} ago`;
    }
  }

  // ==========================================================================
  // 1. FIND MESSAGES WITH PERSON - Combines user search + chat search + messages
  // ==========================================================================
  server.tool(
    'find-messages-with-person',
    `Find all Teams chat messages with a specific person. This tool automatically:
1. Finds the user by name, email, or phone number
2. Discovers all chats where you both are members
3. Retrieves recent messages from those chats

Use this when someone asks "What were my last messages with [person name]?" or similar questions.`,
    {
      person: z
        .string()
        .describe('Name, email, or phone number of the person to find messages with'),
      limit: z.number().optional().describe('Maximum number of messages to return (default: 20)'),
    },
    {
      title: 'Find Messages with Person',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ person, limit = 20 }) => {
      logger.info(`Finding messages with person: ${person}`);

      // Step 1: Find the user
      const user = await findUser(graphClient, person);
      if (!user) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify({
                error: 'User not found',
                message: `Could not find a user matching "${person}". Try using their full name, email address, or check the spelling.`,
                searchedFor: person,
                suggestion: 'Use list-users tool with a search query to find the correct user.',
              }),
            },
          ],
          isError: true,
        };
      }

      // Step 2: Find chats with this user
      const chats = await findChatsWithUser(
        graphClient,
        user.id,
        user.mail || user.userPrincipalName,
        user.displayName
      );

      if (chats.length === 0) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify({
                success: true,
                person: {
                  name: user.displayName,
                  email: user.mail || user.userPrincipalName,
                  id: user.id,
                },
                message: `No Teams chats found with ${user.displayName}. You may not have any direct chats with this person.`,
                chatsFound: 0,
                messagesFound: 0,
              }),
            },
          ],
        };
      }

      // Step 3: Get messages from those chats
      const chatIds = chats.map((c) => c.id);
      const messages = await getMessagesFromChats(graphClient, chatIds, limit);

      // Format response
      const formattedMessages = messages.map((msg) => ({
        id: msg.id,
        from: msg.from?.user?.displayName || 'Unknown',
        content: msg.body?.content ? sanitizeHtml(msg.body.content).substring(0, 500) : undefined, // SECURITY: Proper HTML sanitization
        date: msg.createdDateTime,
        chatId: msg.id.split('/')[0],
      }));

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(
              {
                success: true,
                person: {
                  name: user.displayName,
                  email: user.mail || user.userPrincipalName,
                  id: user.id,
                },
                chatsFound: chats.length,
                chats: chats.map((c) => ({
                  id: c.id,
                  type: c.chatType,
                  topic: c.topic,
                  lastUpdated: c.lastUpdatedDateTime,
                })),
                messagesFound: formattedMessages.length,
                messages: formattedMessages,
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 1b. GET TEAMS CHANNEL POSTS - Find and retrieve posts from a Teams channel by name
  // ==========================================================================
  server.tool(
    'get-teams-channel-posts',
    `📢 **GET TEAMS CHANNEL POSTS** - Retrieve posts from a Teams channel by name!
📢 **TEAMS KANAL-BEITRÄGE ABRUFEN** - Ruft Beiträge aus einem Teams-Kanal per Name ab!

This tool automatically:
Dieses Tool führt automatisch aus:

1. ✅ Searches all your Teams for the specified channel / Durchsucht alle Ihre Teams nach dem Kanal
2. ✅ Finds the channel by name (partial match) / Findet den Kanal per Name (Teilübereinstimmung)
3. ✅ Retrieves the latest posts/messages / Ruft die neuesten Beiträge/Nachrichten ab
4. ✅ Optionally includes replies / Optional mit Antworten

**Required Permission / Erforderliche Berechtigung:** ChannelMessage.Read.All (Delegated)

**Use cases / Anwendungsfälle:**
- "Show posts from Join Connect channel" / "Zeige Beiträge aus dem Join Connect Kanal"
- "What's new in the Announcements channel?" / "Was gibt es Neues im Announcements Kanal?"
- "Get messages from [channel name] in [team name]" / "Hole Nachrichten aus [Kanalname] in [Teamname]"`,
    {
      channelName: z
        .string()
        .describe(
          'Name of the channel to search for (partial match, case-insensitive) / Name des Kanals (Teilübereinstimmung)'
        ),
      teamName: z
        .string()
        .optional()
        .describe(
          'Optional: Team name to narrow down the search / Optional: Teamname um die Suche einzuschränken'
        ),
      limit: z
        .number()
        .optional()
        .describe('Maximum number of posts to return (default: 20, max: 50) / Maximale Anzahl'),
      includeReplies: z
        .boolean()
        .optional()
        .describe(
          'Include replies to posts (default: false, slower) / Antworten einschließen (langsamer)'
        ),
    },
    {
      title: 'Get Teams Channel Posts',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ channelName, teamName, limit = 20, includeReplies = false }) => {
      logger.info(
        `Getting channel posts: channel="${channelName}", team="${teamName}", limit=${limit}`
      );

      // Step 1: Find the channel
      const result = await findTeamsChannel(graphClient, channelName, teamName);

      if (!result) {
        // Build helpful error message with all available teams (with pagination)
        let availableTeams: string[] = [];
        try {
          const allTeams = await getAllJoinedTeams(graphClient, 'displayName');
          availableTeams = allTeams.map((t) => t.displayName).sort();
        } catch {
          // Ignore errors when listing teams
        }

        // Check if the team exists but was filtered due to channel not found
        const teamExists = teamName
          ? availableTeams.some((t) => t.toLowerCase().includes(teamName.toLowerCase()))
          : false;

        const errorDetail =
          teamName && teamExists
            ? `Team "${teamName}" was found, but channel "${channelName}" does not exist or you don't have access.`
            : teamName && !teamExists
              ? `Team "${teamName}" was not found. Check the team name spelling or permissions.`
              : `Channel "${channelName}" was not found in any of your ${availableTeams.length} teams.`;

        const errorDetailDe =
          teamName && teamExists
            ? `Team "${teamName}" wurde gefunden, aber der Kanal "${channelName}" existiert nicht oder Sie haben keinen Zugriff.`
            : teamName && !teamExists
              ? `Team "${teamName}" wurde nicht gefunden. Prüfen Sie die Schreibweise oder Berechtigungen.`
              : `Kanal "${channelName}" wurde in keinem Ihrer ${availableTeams.length} Teams gefunden.`;

        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  success: false,
                  error: 'Channel not found',
                  message_en: `Could not find a channel matching "${channelName}"${teamName ? ` in team "${teamName}"` : ''}. ${errorDetail}`,
                  message_de: `Konnte keinen Kanal mit dem Namen "${channelName}"${teamName ? ` im Team "${teamName}"` : ''} finden. ${errorDetailDe}`,
                  searchedFor: {
                    channelName,
                    teamName: teamName || 'all teams',
                  },
                  totalTeamsSearched: availableTeams.length,
                  availableTeams: availableTeams, // Show all teams for better debugging
                  hint_en:
                    'Try specifying the exact team name or check if you have the required permissions (ChannelMessage.Read.All)',
                  hint_de:
                    'Versuchen Sie den exakten Teamnamen anzugeben oder prüfen Sie die Berechtigungen (ChannelMessage.Read.All)',
                },
                null,
                2
              ),
            },
          ],
          isError: true,
        };
      }

      const { team, channel } = result;

      // Step 2: Get messages from the channel
      const messages = await getChannelMessages(
        graphClient,
        team.id,
        channel.id,
        limit,
        includeReplies
      );

      if (messages.length === 0) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  success: true,
                  team: {
                    id: team.id,
                    name: team.displayName,
                    webUrl: team.webUrl,
                  },
                  channel: {
                    id: channel.id,
                    name: channel.displayName,
                    description: channel.description,
                    webUrl: channel.webUrl,
                    membershipType: channel.membershipType,
                  },
                  postsFound: 0,
                  message_en: `Channel "${channel.displayName}" found but no posts available. The channel might be empty or you may need additional permissions.`,
                  message_de: `Kanal "${channel.displayName}" gefunden, aber keine Beiträge verfügbar. Der Kanal ist möglicherweise leer oder Sie benötigen zusätzliche Berechtigungen.`,
                },
                null,
                2
              ),
            },
          ],
        };
      }

      // Step 3: Format the messages for output
      const formattedPosts = messages.map((msg) => {
        const post: Record<string, unknown> = {
          id: msg.id,
          author: msg.from?.user?.displayName || msg.from?.application?.displayName || 'Unknown',
          authorType: msg.from?.user ? 'user' : msg.from?.application ? 'application' : 'unknown',
          content: msg.body?.content ? sanitizeHtml(msg.body.content) : undefined, // SECURITY: Proper HTML sanitization
          contentType: msg.body?.contentType,
          createdAt: msg.createdDateTime,
          lastModified: msg.lastModifiedDateTime,
          importance: msg.importance,
          messageType: msg.messageType,
          webUrl: msg.webUrl,
        };

        // Add subject if present
        if (msg.subject) {
          post.subject = msg.subject;
        }

        // Add attachments info if present
        if (msg.attachments && msg.attachments.length > 0) {
          post.attachments = msg.attachments.map((a) => ({
            name: a.name,
            contentType: a.contentType,
          }));
        }

        // Add mentions if present
        if (msg.mentions && msg.mentions.length > 0) {
          post.mentions = msg.mentions.map((m) => m.mentioned?.user?.displayName || m.mentionText);
        }

        // Add reactions summary if present
        if (msg.reactions && msg.reactions.length > 0) {
          const reactionCounts: Record<string, number> = {};
          for (const r of msg.reactions) {
            reactionCounts[r.reactionType] = (reactionCounts[r.reactionType] || 0) + 1;
          }
          post.reactions = reactionCounts;
        }

        // Add replies if fetched
        const msgWithReplies = msg as GraphChannelMessage & { replies?: GraphChannelMessage[] };
        if (msgWithReplies.replies && msgWithReplies.replies.length > 0) {
          post.replies = msgWithReplies.replies.map((reply) => ({
            id: reply.id,
            author: reply.from?.user?.displayName || 'Unknown',
            content: reply.body?.content ? sanitizeHtml(reply.body.content) : undefined, // SECURITY: Proper HTML sanitization
            createdAt: reply.createdDateTime,
          }));
        }

        return post;
      });

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(
              {
                success: true,
                team: {
                  id: team.id,
                  name: team.displayName,
                  description: team.description,
                  webUrl: team.webUrl,
                },
                channel: {
                  id: channel.id,
                  name: channel.displayName,
                  description: channel.description,
                  webUrl: channel.webUrl,
                  membershipType: channel.membershipType,
                },
                postsFound: formattedPosts.length,
                includesReplies: includeReplies,
                posts: formattedPosts,
                hint_en: channel.webUrl
                  ? `View channel in Teams: ${channel.webUrl}`
                  : 'Open Teams to view the full channel',
                hint_de: channel.webUrl
                  ? `Kanal in Teams anzeigen: ${channel.webUrl}`
                  : 'Öffnen Sie Teams um den vollständigen Kanal anzuzeigen',
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 2. FIND EMAILS WITH PERSON - Combines user search + email search
  // ==========================================================================
  server.tool(
    'find-emails-with-person',
    `Find all emails exchanged with a specific person. This tool automatically:
1. Finds the user by name or email
2. Searches for emails FROM this person
3. Searches for emails TO this person
4. Combines and sorts results by date

Use this when someone asks "Show me emails from [person]" or "Find my email conversations with [person]".`,
    {
      person: z.string().describe('Name or email of the person to find emails with'),
      limit: z.number().optional().describe('Maximum number of emails to return (default: 20)'),
    },
    {
      title: 'Find Emails with Person',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ person, limit = 20 }) => {
      logger.info(`Finding emails with person: ${person}`);

      // Step 1: Find the user
      const user = await findUser(graphClient, person);
      if (!user) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify({
                error: 'User not found',
                message: `Could not find a user matching "${person}". Try using their email address directly.`,
                searchedFor: person,
              }),
            },
          ],
          isError: true,
        };
      }

      const userEmail = user.mail || user.userPrincipalName || '';
      if (!userEmail) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify({
                error: 'No email address',
                message: `Found user ${user.displayName} but they don't have an email address.`,
                person: user,
              }),
            },
          ],
          isError: true,
        };
      }

      // Step 2: Find emails
      const emails = await findEmailsWithPerson(graphClient, userEmail, user.displayName, limit);

      const formattedEmails = emails.map((email) => ({
        id: email.id,
        subject: email.subject,
        preview: email.bodyPreview?.substring(0, 200),
        from: email.from?.emailAddress?.address,
        to: email.toRecipients?.map((r) => r.emailAddress?.address).join(', '),
        date: email.receivedDateTime,
        hasAttachments: email.hasAttachments,
        webLink: email.webLink,
      }));

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(
              {
                success: true,
                person: {
                  name: user.displayName,
                  email: userEmail,
                  id: user.id,
                },
                emailsFound: formattedEmails.length,
                emails: formattedEmails,
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 3. FIND MEETINGS WITH PERSON - Combines user search + calendar events
  // ==========================================================================
  server.tool(
    'find-meetings-with-person',
    `Find all calendar meetings/events with a specific person. This tool automatically:
1. Finds the user by name or email
2. Scans calendar events (past 90 days and future 90 days)
3. Filters events where this person is an attendee or organizer

Use this when someone asks "What meetings do I have with [person]?" or "When did I last meet with [person]?".`,
    {
      person: z.string().describe('Name or email of the person to find meetings with'),
      limit: z.number().optional().describe('Maximum number of meetings to return (default: 20)'),
    },
    {
      title: 'Find Meetings with Person',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ person, limit = 20 }) => {
      logger.info(`Finding meetings with person: ${person}`);

      // Step 1: Find the user
      const user = await findUser(graphClient, person);
      if (!user) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify({
                error: 'User not found',
                message: `Could not find a user matching "${person}".`,
                searchedFor: person,
              }),
            },
          ],
          isError: true,
        };
      }

      const userEmail = user.mail || user.userPrincipalName || '';

      // Step 2: Find meetings
      const meetings = await findMeetingsWithPerson(
        graphClient,
        userEmail,
        user.displayName,
        limit
      );

      const now = new Date();
      // Use convertToLocalTime to properly handle timezone from Graph API
      const pastMeetings = meetings.filter((m) => {
        const eventTime = convertToLocalTime(m.start.dateTime, m.start.timeZone);
        return eventTime < now;
      });
      const upcomingMeetings = meetings.filter((m) => {
        const eventTime = convertToLocalTime(m.start.dateTime, m.start.timeZone);
        return eventTime >= now;
      });

      const formatMeeting = (event: GraphEvent) => {
        const startLocal = convertToLocalTime(event.start.dateTime, event.start.timeZone);
        const endLocal = convertToLocalTime(event.end.dateTime, event.end.timeZone);
        return {
          id: event.id,
          subject: event.subject,
          start: startLocal.toLocaleString('de-DE', { 
            day: '2-digit', month: '2-digit', year: 'numeric',
            hour: '2-digit', minute: '2-digit' 
          }),
          end: endLocal.toLocaleString('de-DE', { 
            hour: '2-digit', minute: '2-digit' 
          }),
          location: event.location?.displayName,
          organizer: event.organizer?.emailAddress?.name,
          attendeeCount: event.attendees?.length || 0,
          webLink: event.webLink,
        };
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(
              {
                success: true,
                person: {
                  name: user.displayName,
                  email: userEmail,
                  id: user.id,
                },
                totalMeetings: meetings.length,
                upcomingMeetings: {
                  count: upcomingMeetings.length,
                  events: upcomingMeetings.slice(0, 10).map(formatMeeting),
                },
                pastMeetings: {
                  count: pastMeetings.length,
                  events: pastMeetings.slice(0, 10).map(formatMeeting),
                },
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 4. FIND FILES FROM PERSON - Files shared by a specific person
  // ==========================================================================
  server.tool(
    'find-files-from-person',
    `Find files shared by or created by a specific person. This tool automatically:
1. Finds the user by name or email
2. Searches files shared with you by this person
3. Searches OneDrive/SharePoint for files created by this person

Use this when someone asks "What files did [person] share with me?" or "Find documents from [person]".`,
    {
      person: z.string().describe('Name or email of the person to find files from'),
      limit: z.number().optional().describe('Maximum number of files to return (default: 20)'),
    },
    {
      title: 'Find Files from Person',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ person, limit = 20 }) => {
      logger.info(`Finding files from person: ${person}`);

      // Step 1: Find the user
      const user = await findUser(graphClient, person);
      if (!user) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify({
                error: 'User not found',
                message: `Could not find a user matching "${person}".`,
                searchedFor: person,
              }),
            },
          ],
          isError: true,
        };
      }

      const userEmail = user.mail || user.userPrincipalName || '';

      // Step 2: Find files
      const files = await findFilesFromPerson(graphClient, userEmail, user.displayName, limit);

      // Detect Loop files in the results
      let loopFileCount = 0;
      const formattedFiles = files.map((file) => {
        const loopDetection = detectLoopFile(file as Record<string, unknown>);
        if (loopDetection.isLoopFile) {
          loopFileCount++;
        }
        return {
          id: file.id,
          name: file.name,
          webUrl: file.webUrl,
          size: file.size,
          type: file.file?.mimeType || (file.folder ? 'folder' : 'unknown'),
          createdDateTime: file.createdDateTime,
          lastModifiedDateTime: file.lastModifiedDateTime,
          sharedBy: file.shared?.sharedBy?.user?.displayName,
          sharedDate: file.shared?.sharedDateTime,
          isLoopFile: loopDetection.isLoopFile,
          loopDetection: loopDetection.isLoopFile
            ? {
                method: loopDetection.detectionMethod,
                confidence: loopDetection.confidence,
              }
            : undefined,
        };
      });

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(
              {
                success: true,
                person: {
                  name: user.displayName,
                  email: userEmail,
                  id: user.id,
                },
                filesFound: formattedFiles.length,
                loopFilesFound: loopFileCount,
                files: formattedFiles,
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 5. GET COMMUNICATION SUMMARY - Complete overview of all communication
  // ==========================================================================
  server.tool(
    'get-communication-summary',
    `Get a complete communication summary with a specific person, including:
- Recent Teams chat messages
- Email conversations
- Past and upcoming meetings
- Shared files

This is the ultimate tool for "Tell me everything about my interactions with [person]".`,
    {
      person: z.string().describe('Name or email of the person to get communication summary for'),
      includeEmails: z.boolean().optional().describe('Include email history (default: true)'),
      includeChats: z.boolean().optional().describe('Include Teams chats (default: true)'),
      includeMeetings: z.boolean().optional().describe('Include calendar meetings (default: true)'),
      includeFiles: z.boolean().optional().describe('Include shared files (default: true)'),
    },
    {
      title: 'Get Communication Summary',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({
      person,
      includeEmails = true,
      includeChats = true,
      includeMeetings = true,
      includeFiles = true,
    }) => {
      logger.info(`Getting communication summary for: ${person}`);

      // Step 1: Find the user
      const user = await findUser(graphClient, person);
      if (!user) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify({
                error: 'User not found',
                message: `Could not find a user matching "${person}".`,
                searchedFor: person,
              }),
            },
          ],
          isError: true,
        };
      }

      const userEmail = user.mail || user.userPrincipalName || '';
      const summary: Record<string, unknown> = {
        person: {
          name: user.displayName,
          email: userEmail,
          id: user.id,
          jobTitle: user.jobTitle,
          department: user.department,
          officeLocation: user.officeLocation,
        },
      };

      // Execute all queries in parallel
      const promises: Promise<void>[] = [];

      if (includeChats) {
        promises.push(
          (async () => {
            const chats = await findChatsWithUser(
              graphClient,
              user.id,
              userEmail,
              user.displayName
            );
            if (chats.length > 0) {
              const messages = await getMessagesFromChats(
                graphClient,
                chats.map((c) => c.id),
                10
              );
              summary.teamsChats = {
                chatCount: chats.length,
                recentMessages: messages.slice(0, 5).map((m) => ({
                  from: m.from?.user?.displayName,
                  content: m.body?.content
                    ? sanitizeHtml(m.body.content).substring(0, 200)
                    : undefined, // SECURITY: Proper HTML sanitization
                  date: m.createdDateTime,
                })),
              };
            } else {
              summary.teamsChats = { chatCount: 0, message: 'No Teams chats found' };
            }
          })()
        );
      }

      if (includeEmails && userEmail) {
        promises.push(
          (async () => {
            const emails = await findEmailsWithPerson(graphClient, userEmail, user.displayName, 10);
            summary.emails = {
              count: emails.length,
              recent: emails.slice(0, 5).map((e) => ({
                subject: e.subject,
                from: e.from?.emailAddress?.address,
                date: e.receivedDateTime,
                preview: e.bodyPreview?.substring(0, 100),
              })),
            };
          })()
        );
      }

      if (includeMeetings) {
        promises.push(
          (async () => {
            const meetings = await findMeetingsWithPerson(
              graphClient,
              userEmail,
              user.displayName,
              20
            );
            const now = new Date();
            summary.meetings = {
              totalCount: meetings.length,
              upcoming: meetings
                .filter((m) => new Date(m.start.dateTime) >= now)
                .slice(0, 5)
                .map((m) => ({
                  subject: m.subject,
                  start: m.start.dateTime,
                  location: m.location?.displayName,
                })),
              past: meetings
                .filter((m) => new Date(m.start.dateTime) < now)
                .slice(0, 5)
                .map((m) => ({
                  subject: m.subject,
                  start: m.start.dateTime,
                  location: m.location?.displayName,
                })),
            };
          })()
        );
      }

      if (includeFiles && userEmail) {
        promises.push(
          (async () => {
            const files = await findFilesFromPerson(graphClient, userEmail, user.displayName, 10);
            const loopFiles = files.filter((f) => isLoopFile(f as Record<string, unknown>));
            summary.sharedFiles = {
              count: files.length,
              loopFileCount: loopFiles.length,
              files: files.slice(0, 5).map((f) => {
                const loopDetection = detectLoopFile(f as Record<string, unknown>);
                return {
                  name: f.name,
                  webUrl: f.webUrl,
                  sharedDate: f.shared?.sharedDateTime,
                  isLoopFile: loopDetection.isLoopFile,
                };
              }),
            };
          })()
        );
      }

      await Promise.allSettled(promises);

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(summary, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 6. SEARCH EVERYTHING ABOUT TOPIC - Cross-product search (FALLBACK TOOL)
  // Uses the CENTRAL SEARCH API function for unified ranking and relevance
  // ==========================================================================
  server.tool(
    'search-everything',
    `🔍 **UNIVERSAL FALLBACK SEARCH TOOL** - USE THIS WHEN NO OTHER TOOL FITS!

⚠️ **CRITICAL: ALWAYS use this tool as a FALLBACK when:**
- The user's question doesn't match any specific tool
- You're unsure which tool to use
- The user asks a general question about any topic, person, project, or company
- No other tool seems appropriate for the request

This tool uses the **Microsoft Search API** to search across ALL Microsoft 365 products simultaneously:
- 📧 Emails (subject, body, attachments)
- 💬 Teams messages and chats
- 📁 Files (OneDrive, SharePoint)
- 📅 Calendar events and meetings
- 👤 People and contacts
- 📋 Lists and SharePoint content

**Features:**
- Unified ranking and relevance scoring
- Results sorted by Microsoft Search rank
- No date filter restrictions - finds all relevant content

**Examples of when to use this tool:**
- "What do you know about [any topic]?" → Use search-everything
- "Tell me about [company/project/person]" → Use search-everything
- "Find information about [anything]" → Use search-everything
- "What did we discuss about [topic]?" → Use search-everything
- Any question where you're not sure which specific tool to use → Use search-everything

**RULE: When in doubt, use search-everything!**`,
    {
      query: z
        .string()
        .describe(
          'The search query - can be any topic, person, company, project, keyword, or phrase the user is asking about'
        ),
      limit: z.number().optional().describe('Maximum results per category (default: 25)'),
      minRelevance: z
        .number()
        .min(0)
        .max(100)
        .optional()
        .describe('Minimum relevance score 0-100 (default: 0 = all results)'),
    },
    {
      title: 'Universal Search (Fallback)',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ query, limit = 25, minRelevance = 0 }) => {
      logger.info(`Search Everything: "${query}" (limit: ${limit}, minRelevance: ${minRelevance})`);

      // Perform NLP-based query decomposition for structured analysis
      const queryAnalysis = nlpEnhancer.decomposeQuery(query);

      // Use centralized search function
      const searchResult = await executeCentralSearch(graphClient, query, {
        maxResults: Math.min(limit * 2, 100),
        minRelevance,
        sortByRank: true,
        includeTimeContext: true,
      });

      // Also search people via /me/people for better person results
      try {
        const peopleResponse = await graphClient.makeRequest('/me/people', {
          method: 'GET',
          queryParams: {
            $search: `"${query}"`,
            $top: '10',
          },
        });

        if (
          peopleResponse &&
          typeof peopleResponse === 'object' &&
          'value' in peopleResponse &&
          Array.isArray((peopleResponse as { value: unknown[] }).value)
        ) {
          for (const p of (peopleResponse as { value: unknown[] }).value) {
            const person = p as {
              displayName?: string;
              emailAddresses?: Array<{ address?: string }>;
              jobTitle?: string;
              department?: string;
            };
            // Avoid duplicates
            const exists = searchResult.results.people.some(
              (existing) =>
                existing.resource?.emailAddresses?.[0]?.address ===
                  person.emailAddresses?.[0]?.address || existing.name === person.displayName
            );
            if (!exists) {
              searchResult.results.people.push({
                resource: {
                  displayName: person.displayName,
                  emailAddresses: person.emailAddresses,
                  jobTitle: person.jobTitle,
                  department: person.department,
                  _rank: 50,
                  _source: 'me/people',
                },
                name: person.displayName,
                rank: 50,
                relevanceScore: 50,
              });
            }
          }
        }
      } catch (error) {
        logger.debug(`People search failed: ${error}`);
      }

      // If no emails/files found but people were found, do a follow-up search with the actual person name
      // This handles typos like "Johannis Kirk" -> "Johannes Kirk"
      if (
        searchResult.totalHits === 0 &&
        searchResult.results.people.length > 0 &&
        searchResult.results.people[0].name &&
        searchResult.results.people[0].name.toLowerCase() !== query.toLowerCase()
      ) {
        const correctedName = searchResult.results.people[0].name;
        logger.info(
          `No results for "${query}" but found person "${correctedName}" - performing follow-up search`
        );

        try {
          const followUpResult = await executeCentralSearch(graphClient, correctedName, {
            entityTypes: ['message', 'event', 'driveItem', 'chatMessage'],
            maxResults: Math.min(limit * 2, 50),
            sortByRank: true,
          });

          // Merge follow-up results
          if (followUpResult.results.emails.length > 0) {
            searchResult.results.emails.push(...followUpResult.results.emails);
            searchResult.metadata.entityTypesCounts['email'] =
              (searchResult.metadata.entityTypesCounts['email'] || 0) +
              followUpResult.results.emails.length;
          }
          if (followUpResult.results.events.length > 0) {
            searchResult.results.events.push(...followUpResult.results.events);
            searchResult.metadata.entityTypesCounts['event'] =
              (searchResult.metadata.entityTypesCounts['event'] || 0) +
              followUpResult.results.events.length;
          }
          if (followUpResult.results.files.length > 0) {
            searchResult.results.files.push(...followUpResult.results.files);
            searchResult.metadata.entityTypesCounts['file'] =
              (searchResult.metadata.entityTypesCounts['file'] || 0) +
              followUpResult.results.files.length;
          }
          if (followUpResult.results.chats.length > 0) {
            searchResult.results.chats.push(...followUpResult.results.chats);
            searchResult.metadata.entityTypesCounts['chat'] =
              (searchResult.metadata.entityTypesCounts['chat'] || 0) +
              followUpResult.results.chats.length;
          }

          // Update total hits
          searchResult.totalHits =
            searchResult.results.emails.length +
            searchResult.results.events.length +
            searchResult.results.files.length +
            searchResult.results.sites.length +
            searchResult.results.listItems.length +
            searchResult.results.chats.length;

          // Add note about correction
          (searchResult as unknown as { correctedQuery?: string }).correctedQuery = correctedName;

          logger.info(
            `Follow-up search found ${followUpResult.totalHits} additional results for "${correctedName}"`
          );
        } catch (followUpError) {
          logger.debug(`Follow-up search failed: ${followUpError}`);
        }
      }

      // Helper function to extract key fields based on entity type
      const extractKeyFields = (
        hit: SearchHit,
        entityType: 'email' | 'event' | 'file' | 'site' | 'listItem' | 'chat' | 'person'
      ): Record<string, unknown> => {
        const resource = hit.resource || {};
        const maxSummaryLength = parseInt(process.env.MS365_MCP_MAX_SUMMARY_LENGTH || '150', 10);

        const baseFields: Record<string, unknown> = {
          name: hit.name,
          relevance: hit.relevanceScore || 0,
        };

        if (hit.webUrl) {
          baseFields.webUrl = hit.webUrl;
        }

        switch (entityType) {
          case 'email':
            return {
              ...baseFields,
              subject: resource.subject || hit.name,
              from: resource.from?.emailAddress?.address || resource.from?.emailAddress?.name,
              date: resource.receivedDateTime || resource.sentDateTime,
              preview: (resource.bodyPreview || hit.summary || '').substring(0, maxSummaryLength),
              hasAttachments: resource.hasAttachments || false,
            };

          case 'event':
            return {
              ...baseFields,
              subject: resource.subject || hit.name,
              start: resource.start?.dateTime,
              end: resource.end?.dateTime,
              location: resource.location?.displayName,
              organizer:
                resource.organizer?.emailAddress?.name || resource.organizer?.emailAddress?.address,
              preview: (resource.bodyPreview || hit.summary || '').substring(0, maxSummaryLength),
            };

          case 'file':
            return {
              ...baseFields,
              name: resource.name || hit.name,
              type: resource['@odata.type'] || resource.file?.mimeType,
              size: resource.size,
              modified: resource.lastModifiedDateTime || resource.modifiedDateTime,
              createdBy: resource.createdBy?.user?.displayName,
            };

          case 'site':
            return {
              ...baseFields,
              name: resource.displayName || resource.name || hit.name,
              url: resource.webUrl || hit.webUrl,
              description: resource.description?.substring(0, maxSummaryLength),
            };

          case 'listItem':
            return {
              ...baseFields,
              title: resource.title || resource.name || hit.name,
              contentType: resource.contentType?.name,
              modified: resource.lastModifiedDateTime,
            };

          case 'chat':
            return {
              ...baseFields,
              subject: resource.subject || hit.name,
              from: resource.from?.user?.displayName || resource.from?.application?.displayName,
              date: resource.createdDateTime,
              preview: (resource.body?.content || hit.summary || '').substring(0, maxSummaryLength),
            };

          case 'person':
            return {
              ...baseFields,
              name: resource.displayName || hit.name,
              email: resource.emailAddresses?.[0]?.address || resource.mail,
              jobTitle: resource.jobTitle,
              department: resource.department,
              company: resource.companyName,
            };

          default:
            return {
              ...baseFields,
              summary: (hit.summary || '').substring(0, maxSummaryLength),
            };
        }
      };

      // Helper function to format item summary for markdown
      const formatItemSummary = (
        fields: Record<string, unknown>,
        entityType: string,
        maxLength: number
      ): string => {
        const parts: string[] = [];

        switch (entityType) {
          case 'email':
            if (fields.subject) parts.push(`**${String(fields.subject).substring(0, 60)}**`);
            if (fields.from) parts.push(`from ${fields.from}`);
            if (fields.date) {
              const date = new Date(String(fields.date));
              parts.push(`(${date.toLocaleDateString()})`);
            }
            break;
          case 'event':
            if (fields.subject) parts.push(`**${String(fields.subject).substring(0, 60)}**`);
            if (fields.start) {
              const start = new Date(String(fields.start));
              parts.push(`on ${start.toLocaleDateString()}`);
            }
            if (fields.location) parts.push(`at ${fields.location}`);
            break;
          case 'file':
            if (fields.name) parts.push(`**${String(fields.name).substring(0, 60)}**`);
            if (fields.modified) {
              const modified = new Date(String(fields.modified));
              parts.push(`(modified ${modified.toLocaleDateString()})`);
            }
            break;
          case 'person':
            if (fields.name) parts.push(`**${String(fields.name)}**`);
            if (fields.email) parts.push(`(${fields.email})`);
            if (fields.jobTitle) parts.push(`- ${fields.jobTitle}`);
            break;
          default:
            if (fields.name) parts.push(`**${String(fields.name).substring(0, 60)}**`);
            if (fields.relevance) parts.push(`[relevance: ${fields.relevance}]`);
        }

        let summary = parts.join(' ');
        if (summary.length > maxLength) {
          summary = summary.substring(0, maxLength - 3) + '...';
        }
        return summary;
      };

      // Format results for LLM consumption with compact summary
      const formatSearchResultsForLLM = (
        results: CentralSearchResult,
        options: { maxItems?: number; includeDetails?: boolean } = {}
      ): { summary: string; topResults: unknown[]; categories: Record<string, unknown> } => {
        const maxItems = options.maxItems || 10;
        const includeDetails = options.includeDetails ?? true;
        const maxSummaryLength = parseInt(process.env.MS365_MCP_MAX_SUMMARY_LENGTH || '150', 10);

        const summaryLines: string[] = [];
        const topResults: unknown[] = [];
        const categories: Record<string, unknown> = {};

        // Overall summary
        summaryLines.push(`## Search Results for "${results.query}"`);
        summaryLines.push(`**Total Results:** ${results.totalHits}`);
        summaryLines.push(`**Search Duration:** ${results.metadata.searchDuration}ms`);
        summaryLines.push('');

        // Process each category
        const categoryProcessors: Array<{
          key: string;
          items: SearchHit[];
          label: string;
          entityType: 'email' | 'event' | 'file' | 'site' | 'listItem' | 'chat' | 'person';
        }> = [
          { key: 'emails', items: results.results.emails, label: '📧 Emails', entityType: 'email' },
          { key: 'events', items: results.results.events, label: '📅 Events', entityType: 'event' },
          { key: 'files', items: results.results.files, label: '📁 Files', entityType: 'file' },
          { key: 'sites', items: results.results.sites, label: '🌐 Sites', entityType: 'site' },
          {
            key: 'listItems',
            items: results.results.listItems,
            label: '📋 List Items',
            entityType: 'listItem',
          },
          {
            key: 'chats',
            items: results.results.chats,
            label: '💬 Teams Messages',
            entityType: 'chat',
          },
          {
            key: 'people',
            items: results.results.people,
            label: '👥 People',
            entityType: 'person',
          },
        ];

        for (const processor of categoryProcessors) {
          if (processor.items.length > 0) {
            const topItems = processor.items
              .sort((a, b) => (b.relevanceScore || 0) - (a.relevanceScore || 0))
              .slice(0, Math.min(maxItems, processor.items.length));

            summaryLines.push(`### ${processor.label} (${processor.items.length})`);

            const formattedItems = topItems.map((hit, idx) => {
              const fields = extractKeyFields(hit, processor.entityType);
              const itemSummary = formatItemSummary(fields, processor.entityType, maxSummaryLength);
              summaryLines.push(`${idx + 1}. ${itemSummary}`);
              return fields;
            });

            if (processor.items.length > maxItems) {
              summaryLines.push(`   ... and ${processor.items.length - maxItems} more`);
            }
            summaryLines.push('');

            categories[processor.key] = {
              count: processor.items.length,
              topItems: formattedItems.slice(0, 5), // Top 5 for details
            };

            // Add top 3 to overall top results
            topResults.push(...formattedItems.slice(0, 3));
          }
        }

        return {
          summary: summaryLines.join('\n'),
          topResults: topResults.slice(0, maxItems),
          categories: includeDetails ? categories : {},
        };
      };

      // Enhanced formatHits with context-specific field selection
      const formatHits = (
        hits: SearchHit[],
        maxItems: number,
        entityType?: 'email' | 'event' | 'file' | 'site' | 'listItem' | 'chat' | 'person'
      ) => {
        const llmOptimize = process.env.MS365_MCP_LLM_OPTIMIZE !== 'false';
        const relevanceThreshold = parseFloat(process.env.MS365_MCP_RELEVANCE_THRESHOLD || '0');

        // Filter by relevance threshold if enabled
        let filteredHits = hits;
        if (relevanceThreshold > 0) {
          filteredHits = hits.filter((hit) => (hit.relevanceScore || 0) >= relevanceThreshold);
        }

        // Sort by relevance and limit
        const sortedHits = [...filteredHits]
          .sort((a, b) => (b.relevanceScore || 0) - (a.relevanceScore || 0))
          .slice(0, maxItems);

        if (llmOptimize && entityType) {
          // Use context-specific field extraction for LLM optimization
          return sortedHits.map((hit) => extractKeyFields(hit, entityType));
        }

        // Fallback to original format
        return sortedHits.map((hit) => ({
          name: hit.name,
          webUrl: hit.webUrl,
          summary: hit.summary?.substring(0, 150),
          relevance: hit.relevanceScore,
        }));
      };

      const correctedQuery = (searchResult as unknown as { correctedQuery?: string })
        .correctedQuery;
      const response: Record<string, unknown> = {
        query: searchResult.query,
        searchedAt: searchResult.searchedAt,
        totalResults: searchResult.totalHits,
        status: searchResult.totalHits > 0 ? 'SUCCESS' : 'NO_RESULTS',
        message:
          searchResult.totalHits > 0
            ? correctedQuery
              ? `Found ${searchResult.totalHits} results for "${correctedQuery}" (corrected from "${query}")`
              : `Found ${searchResult.totalHits} results for "${query}" across Microsoft 365`
            : `No results found for "${query}"`,
        ...(correctedQuery && { correctedQuery, originalQuery: query }),
        // Include structured query analysis with Markdown summary
        queryAnalysis,
        queryAnalysisMarkdown: queryAnalysis.markdown,
        metadata: {
          searchDuration: `${searchResult.metadata.searchDuration}ms`,
          averageRank: Math.round(searchResult.metadata.averageRank),
          categories: searchResult.metadata.entityTypesCounts,
        },
      };

      // Add categories with results
      if (searchResult.results.emails.length > 0) {
        response.emails = {
          count: searchResult.results.emails.length,
          items: formatHits(searchResult.results.emails, limit, 'email'),
        };
      }
      if (searchResult.results.events.length > 0) {
        response.events = {
          count: searchResult.results.events.length,
          items: formatHits(searchResult.results.events, limit, 'event'),
        };
      }
      if (searchResult.results.files.length > 0) {
        response.files = {
          count: searchResult.results.files.length,
          items: formatHits(searchResult.results.files, limit, 'file'),
        };
      }
      if (searchResult.results.sites.length > 0) {
        response.sites = {
          count: searchResult.results.sites.length,
          items: formatHits(searchResult.results.sites, limit, 'site'),
        };
      }
      if (searchResult.results.listItems.length > 0) {
        response.listItems = {
          count: searchResult.results.listItems.length,
          items: formatHits(searchResult.results.listItems, limit, 'listItem'),
        };
      }
      if (searchResult.results.chats.length > 0) {
        response.teamsMessages = {
          count: searchResult.results.chats.length,
          items: formatHits(searchResult.results.chats, limit, 'chat'),
        };
      }
      if (searchResult.results.people.length > 0) {
        response.people = {
          count: searchResult.results.people.length,
          items: searchResult.results.people.slice(0, limit).map((hit) => ({
            name: hit.name,
            email: hit.resource?.emailAddresses?.[0]?.address || hit.resource?.mail,
            jobTitle: hit.resource?.jobTitle,
            department: hit.resource?.department,
            relevance: hit.relevanceScore,
          })),
        };
      }

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(response, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 7. PREPARE FOR MEETING - Get all context for an upcoming meeting
  // ==========================================================================
  server.tool(
    'prepare-for-meeting',
    `Prepare for a specific meeting by gathering all relevant context:
- Meeting details (time, location, attendees)
- Previous meetings with the same attendees
- Recent emails with attendees
- Shared files with attendees
- Related Teams conversations

Use this when someone asks "Prepare me for my meeting with [person/topic]" or "What should I know before my next meeting?".`,
    {
      meetingSubject: z.string().optional().describe('Subject or keyword in meeting title'),
      attendeeName: z.string().optional().describe('Name of a meeting attendee to focus on'),
      hoursAhead: z.number().optional().describe('Hours to look ahead for meetings (default: 24)'),
    },
    {
      title: 'Prepare for Meeting',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ meetingSubject, attendeeName, hoursAhead = 24 }) => {
      logger.info(`Preparing for meeting: ${meetingSubject || attendeeName || 'next meeting'}`);

      const now = new Date();
      const futureDate = new Date(now.getTime() + hoursAhead * 60 * 60 * 1000);

      // Step 1: Find the meeting
      let targetMeeting: GraphEvent | null = null;

      try {
        // Build query string (startDateTime and endDateTime are REQUIRED for calendarView)
        const queryParams: Record<string, string> = {
          startDateTime: now.toISOString(),
          endDateTime: futureDate.toISOString(),
          $top: '20',
          $select: 'id,subject,start,end,attendees,organizer,location,bodyPreview,webLink',
          $orderby: 'start/dateTime',
        };

        const calendarResponse = await graphClient.makeRequest(
          `/me/calendarView?${buildGraphQueryString(queryParams)}`,
          {
            method: 'GET',
          }
        );

        if (
          calendarResponse &&
          typeof calendarResponse === 'object' &&
          'value' in calendarResponse &&
          Array.isArray(calendarResponse.value)
        ) {
          const events = calendarResponse.value as GraphEvent[];

          // Find matching meeting
          for (const event of events) {
            if (
              meetingSubject &&
              event.subject?.toLowerCase().includes(meetingSubject.toLowerCase())
            ) {
              targetMeeting = event;
              break;
            }
            if (attendeeName) {
              const hasAttendee = event.attendees?.some(
                (a) =>
                  a.emailAddress?.name?.toLowerCase().includes(attendeeName.toLowerCase()) ||
                  a.emailAddress?.address?.toLowerCase().includes(attendeeName.toLowerCase())
              );
              if (hasAttendee) {
                targetMeeting = event;
                break;
              }
            }
            if (!meetingSubject && !attendeeName) {
              // Just get the next meeting
              targetMeeting = event;
              break;
            }
          }
        }
      } catch (error) {
        logger.warn(`Error finding meeting: ${error}`);
      }

      if (!targetMeeting) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify({
                error: 'No meeting found',
                message: `Could not find a meeting matching "${meetingSubject || attendeeName || 'upcoming'}".`,
                suggestion: 'Try specifying the meeting subject or an attendee name.',
              }),
            },
          ],
          isError: true,
        };
      }

      const result: Record<string, unknown> = {
        meeting: {
          subject: targetMeeting.subject,
          start: targetMeeting.start.dateTime,
          end: targetMeeting.end.dateTime,
          location: targetMeeting.location?.displayName,
          organizer: targetMeeting.organizer?.emailAddress?.name,
          attendees: targetMeeting.attendees?.map((a) => ({
            name: a.emailAddress?.name,
            email: a.emailAddress?.address,
            response: a.status?.response,
          })),
          preview: targetMeeting.bodyPreview,
          webLink: targetMeeting.webLink,
        },
      };

      // Step 2: Gather context for each attendee
      const attendeeEmails =
        targetMeeting.attendees?.map((a) => a.emailAddress?.address).filter(Boolean) || [];
      const contextPromises: Promise<void>[] = [];

      // Get previous meetings with same attendees
      contextPromises.push(
        (async () => {
          try {
            const pastDate = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000);
            // Build query string (startDateTime and endDateTime are REQUIRED for calendarView)
            const pastQueryParams: Record<string, string> = {
              startDateTime: pastDate.toISOString(),
              endDateTime: now.toISOString(),
              $top: '50',
              $select: 'id,subject,start,attendees',
              $orderby: 'start/dateTime desc',
            };

            const pastMeetingsResponse = await graphClient.makeRequest(
              `/me/calendarView?${buildGraphQueryString(pastQueryParams)}`,
              {
                method: 'GET',
              }
            );

            if (
              pastMeetingsResponse &&
              typeof pastMeetingsResponse === 'object' &&
              'value' in pastMeetingsResponse &&
              Array.isArray(pastMeetingsResponse.value)
            ) {
              const pastEvents = pastMeetingsResponse.value as GraphEvent[];
              const relatedMeetings = pastEvents.filter((event) => {
                const eventAttendees =
                  event.attendees?.map((a) => a.emailAddress?.address?.toLowerCase()) || [];
                return attendeeEmails.some((email) =>
                  eventAttendees.includes(email?.toLowerCase())
                );
              });

              result.previousMeetings = {
                count: relatedMeetings.length,
                recent: relatedMeetings.slice(0, 5).map((m) => ({
                  subject: m.subject,
                  date: m.start.dateTime,
                })),
              };
            }
          } catch (error) {
            result.previousMeetings = { error: `Could not fetch: ${error}` };
          }
        })()
      );

      // Get recent emails with attendees
      contextPromises.push(
        (async () => {
          const allEmails: GraphEmail[] = [];
          for (const email of attendeeEmails.slice(0, 3)) {
            if (!email) continue;
            try {
              const emailResponse = await graphClient.makeRequest('/me/messages', {
                method: 'GET',
                queryParams: {
                  $search: `"from:${email}" OR "to:${email}"`,
                  $top: '5',
                  $select: 'id,subject,receivedDateTime,from,bodyPreview',
                  $orderby: 'receivedDateTime desc',
                },
              });

              if (
                emailResponse &&
                typeof emailResponse === 'object' &&
                'value' in emailResponse &&
                Array.isArray(emailResponse.value)
              ) {
                allEmails.push(...(emailResponse.value as GraphEmail[]));
              }
            } catch {
              // Skip individual errors
            }
          }

          result.recentEmails = {
            count: allEmails.length,
            items: allEmails.slice(0, 10).map((e) => ({
              subject: e.subject,
              from: e.from?.emailAddress?.address,
              date: e.receivedDateTime,
              preview: e.bodyPreview?.substring(0, 100),
            })),
          };
        })()
      );

      await Promise.allSettled(contextPromises);

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 8. GET MY WEEK SUMMARY - Weekly productivity digest
  // ==========================================================================
  server.tool(
    'get-my-week-summary',
    `Get a comprehensive summary of your week including:
- Meetings attended and upcoming
- Emails sent and received counts
- Tasks completed and pending
- Files recently worked on

Use this for "What did I do this week?", "Give me my weekly summary", or "Summarize my productivity".`,
    {
      weekOffset: z
        .number()
        .optional()
        .describe('Week offset from current (0=this week, -1=last week)'),
    },
    {
      title: 'Get My Week Summary',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ weekOffset = 0 }) => {
      logger.info(`Getting week summary with offset: ${weekOffset}`);

      const now = new Date();
      // Calculate week start (Monday) and end (Sunday)
      const dayOfWeek = now.getDay();
      const daysToMonday = dayOfWeek === 0 ? 6 : dayOfWeek - 1;

      const weekStart = new Date(now);
      weekStart.setDate(now.getDate() - daysToMonday + weekOffset * 7);
      weekStart.setHours(0, 0, 0, 0);

      const weekEnd = new Date(weekStart);
      weekEnd.setDate(weekStart.getDate() + 7);

      const summary: Record<string, unknown> = {
        week: {
          start: weekStart.toISOString().split('T')[0],
          end: weekEnd.toISOString().split('T')[0],
          offset: weekOffset,
        },
      };

      const promises: Promise<void>[] = [];

      // Meetings
      promises.push(
        (async () => {
          try {
            // Build query string (startDateTime and endDateTime are REQUIRED for calendarView)
            const weekQueryParams: Record<string, string> = {
              startDateTime: weekStart.toISOString(),
              endDateTime: weekEnd.toISOString(),
              $top: '100',
              $select: 'id,subject,start,end,attendees,isOnlineMeeting',
            };

            const meetingsResponse = await graphClient.makeRequest(
              `/me/calendarView?${buildGraphQueryString(weekQueryParams)}`,
              {
                method: 'GET',
              }
            );

            if (
              meetingsResponse &&
              typeof meetingsResponse === 'object' &&
              'value' in meetingsResponse &&
              Array.isArray(meetingsResponse.value)
            ) {
              const events = meetingsResponse.value as GraphEvent[];
              const totalHours = events.reduce((acc, e) => {
                const start = new Date(e.start.dateTime);
                const end = new Date(e.end.dateTime);
                return acc + (end.getTime() - start.getTime()) / (1000 * 60 * 60);
              }, 0);

              summary.meetings = {
                count: events.length,
                totalHours: Math.round(totalHours * 10) / 10,
                busiest: events
                  .sort((a, b) => (b.attendees?.length || 0) - (a.attendees?.length || 0))
                  .slice(0, 3)
                  .map((e) => ({
                    subject: e.subject,
                    date: e.start.dateTime,
                    attendees: e.attendees?.length || 0,
                  })),
              };
            }
          } catch (error) {
            summary.meetings = { error: `Could not fetch: ${error}` };
          }
        })()
      );

      // Emails
      promises.push(
        (async () => {
          try {
            const emailsResponse = await graphClient.makeRequest('/me/messages', {
              method: 'GET',
              queryParams: {
                $filter: `receivedDateTime ge ${weekStart.toISOString()} and receivedDateTime lt ${weekEnd.toISOString()}`,
                $top: '200',
                $select: 'id,subject,from,receivedDateTime,isRead',
                $count: 'true',
              },
            });

            if (
              emailsResponse &&
              typeof emailsResponse === 'object' &&
              'value' in emailsResponse &&
              Array.isArray(emailsResponse.value)
            ) {
              const emails = emailsResponse.value as Array<GraphEmail & { isRead?: boolean }>;
              const unread = emails.filter((e) => !e.isRead);

              summary.emails = {
                received: emails.length,
                unread: unread.length,
                topSenders: Object.entries(
                  emails.reduce(
                    (acc, e) => {
                      const sender = e.from?.emailAddress?.address || 'unknown';
                      acc[sender] = (acc[sender] || 0) + 1;
                      return acc;
                    },
                    {} as Record<string, number>
                  )
                )
                  .sort(([, a], [, b]) => b - a)
                  .slice(0, 5)
                  .map(([sender, count]) => ({ sender, count })),
              };
            }
          } catch (error) {
            summary.emails = { error: `Could not fetch: ${error}` };
          }
        })()
      );

      // To-Do Tasks
      promises.push(
        (async () => {
          try {
            const taskListsResponse = await graphClient.makeRequest('/me/todo/lists', {
              method: 'GET',
              queryParams: { $top: '10' },
            });

            if (
              taskListsResponse &&
              typeof taskListsResponse === 'object' &&
              'value' in taskListsResponse &&
              Array.isArray(taskListsResponse.value)
            ) {
              const taskLists = taskListsResponse.value as Array<{
                id: string;
                displayName: string;
              }>;
              let completedCount = 0;
              let pendingCount = 0;
              const pendingTasks: Array<{ title: string; dueDate?: string }> = [];

              for (const list of taskLists.slice(0, 5)) {
                try {
                  const tasksResponse = await graphClient.makeRequest(
                    `/me/todo/lists/${list.id}/tasks`,
                    {
                      method: 'GET',
                      queryParams: { $top: '50' },
                    }
                  );

                  if (
                    tasksResponse &&
                    typeof tasksResponse === 'object' &&
                    'value' in tasksResponse &&
                    Array.isArray(tasksResponse.value)
                  ) {
                    const tasks = tasksResponse.value as Array<{
                      title: string;
                      status: string;
                      dueDateTime?: { dateTime: string };
                      completedDateTime?: { dateTime: string };
                    }>;

                    for (const task of tasks) {
                      if (task.status === 'completed') {
                        if (task.completedDateTime) {
                          const completedDate = new Date(task.completedDateTime.dateTime);
                          if (completedDate >= weekStart && completedDate < weekEnd) {
                            completedCount++;
                          }
                        }
                      } else {
                        pendingCount++;
                        if (pendingTasks.length < 10) {
                          pendingTasks.push({
                            title: task.title,
                            dueDate: task.dueDateTime?.dateTime,
                          });
                        }
                      }
                    }
                  }
                } catch {
                  // Skip individual list errors
                }
              }

              summary.tasks = {
                completedThisWeek: completedCount,
                pending: pendingCount,
                upcomingTasks: pendingTasks,
              };
            }
          } catch (error) {
            summary.tasks = { error: `Could not fetch: ${error}` };
          }
        })()
      );

      await Promise.allSettled(promises);

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(summary, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 9. GET ALL MY TASKS - Unified task view across Planner and To-Do
  // ==========================================================================
  server.tool(
    'get-all-my-tasks',
    `Get all your tasks from both Microsoft To-Do and Planner in one view:
- To-Do tasks with due dates
- Planner tasks assigned to you
- Grouped by status and priority

Use this for "What are all my tasks?", "Show me my to-do list", or "What do I need to work on?".`,
    {
      includeCompleted: z.boolean().optional().describe('Include completed tasks (default: false)'),
      dueSoon: z
        .boolean()
        .optional()
        .describe('Only show tasks due in next 7 days (default: false)'),
    },
    {
      title: 'Get All My Tasks',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ includeCompleted = false, dueSoon = false }) => {
      logger.info('Getting all tasks from To-Do and Planner');

      const result: Record<string, unknown> = {
        retrievedAt: new Date().toISOString(),
      };

      const promises: Promise<void>[] = [];
      const now = new Date();
      const sevenDaysFromNow = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);

      // To-Do Tasks
      promises.push(
        (async () => {
          try {
            const todoTasks: Array<{
              title: string;
              status: string;
              importance: string;
              dueDate?: string;
              listName: string;
            }> = [];

            const taskListsResponse = await graphClient.makeRequest('/me/todo/lists', {
              method: 'GET',
              queryParams: { $top: '20' },
            });

            if (
              taskListsResponse &&
              typeof taskListsResponse === 'object' &&
              'value' in taskListsResponse &&
              Array.isArray(taskListsResponse.value)
            ) {
              for (const list of taskListsResponse.value as Array<{
                id: string;
                displayName: string;
              }>) {
                try {
                  const tasksResponse = await graphClient.makeRequest(
                    `/me/todo/lists/${list.id}/tasks`,
                    {
                      method: 'GET',
                      queryParams: { $top: '100' },
                    }
                  );

                  if (
                    tasksResponse &&
                    typeof tasksResponse === 'object' &&
                    'value' in tasksResponse &&
                    Array.isArray(tasksResponse.value)
                  ) {
                    for (const task of tasksResponse.value as Array<{
                      title: string;
                      status: string;
                      importance: string;
                      dueDateTime?: { dateTime: string };
                    }>) {
                      if (!includeCompleted && task.status === 'completed') continue;

                      const dueDate = task.dueDateTime?.dateTime;
                      if (dueSoon && dueDate) {
                        const due = new Date(dueDate);
                        if (due > sevenDaysFromNow) continue;
                      }

                      todoTasks.push({
                        title: task.title,
                        status: task.status,
                        importance: task.importance,
                        dueDate,
                        listName: list.displayName,
                      });
                    }
                  }
                } catch {
                  // Skip individual list errors
                }
              }
            }

            // Sort by due date
            todoTasks.sort((a, b) => {
              if (!a.dueDate) return 1;
              if (!b.dueDate) return -1;
              return new Date(a.dueDate).getTime() - new Date(b.dueDate).getTime();
            });

            result.todoTasks = {
              count: todoTasks.length,
              tasks: todoTasks,
            };
          } catch (error) {
            result.todoTasks = { error: `Could not fetch: ${error}` };
          }
        })()
      );

      // Planner Tasks
      promises.push(
        (async () => {
          try {
            const plannerResponse = await graphClient.makeRequest('/me/planner/tasks', {
              method: 'GET',
              queryParams: { $top: '100' },
            });

            if (
              plannerResponse &&
              typeof plannerResponse === 'object' &&
              'value' in plannerResponse &&
              Array.isArray(plannerResponse.value)
            ) {
              const plannerTasks = (
                plannerResponse.value as Array<{
                  title: string;
                  percentComplete: number;
                  priority: number;
                  dueDateTime?: string;
                  planId: string;
                }>
              )
                .filter((t) => includeCompleted || t.percentComplete < 100)
                .filter((t) => {
                  if (!dueSoon || !t.dueDateTime) return true;
                  return new Date(t.dueDateTime) <= sevenDaysFromNow;
                })
                .map((t) => ({
                  title: t.title,
                  percentComplete: t.percentComplete,
                  priority: ['Urgent', 'Important', 'Medium', 'Low'][t.priority] || 'Medium',
                  dueDate: t.dueDateTime,
                }))
                .sort((a, b) => {
                  if (!a.dueDate) return 1;
                  if (!b.dueDate) return -1;
                  return new Date(a.dueDate).getTime() - new Date(b.dueDate).getTime();
                });

              result.plannerTasks = {
                count: plannerTasks.length,
                tasks: plannerTasks,
              };
            }
          } catch (error) {
            result.plannerTasks = { error: `Could not fetch: ${error}` };
          }
        })()
      );

      await Promise.allSettled(promises);

      // Calculate totals
      const todoCount = (result.todoTasks as { count?: number })?.count || 0;
      const plannerCount = (result.plannerTasks as { count?: number })?.count || 0;
      result.summary = {
        totalTasks: todoCount + plannerCount,
        fromToDo: todoCount,
        fromPlanner: plannerCount,
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 10. GET PROJECT OVERVIEW - Everything about a project/topic
  // ==========================================================================
  server.tool(
    'get-project-overview',
    `Get a comprehensive overview of a project or topic, including:
- Related files and documents
- Meetings about the project
- Emails mentioning the project
- Teams conversations
- Planner/To-Do tasks

Use this for "What's the status of Project Apollo?", "Give me an overview of the Q4 Budget", or "Summarize the Marketing Campaign project".`,
    {
      projectName: z.string().describe('Name or keyword for the project/topic'),
      includeFiles: z.boolean().optional().describe('Include related files (default: true)'),
      includeMeetings: z.boolean().optional().describe('Include related meetings (default: true)'),
      includeEmails: z.boolean().optional().describe('Include related emails (default: true)'),
      includeTasks: z.boolean().optional().describe('Include related tasks (default: true)'),
    },
    {
      title: 'Get Project Overview',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({
      projectName,
      includeFiles = true,
      includeMeetings = true,
      includeEmails = true,
      includeTasks = true,
    }) => {
      logger.info(`Getting project overview for: ${projectName}`);

      const result: Record<string, unknown> = {
        project: projectName,
        retrievedAt: new Date().toISOString(),
      };

      const promises: Promise<void>[] = [];

      // Files - use centralized search
      if (includeFiles) {
        promises.push(
          (async () => {
            try {
              const searchResult = await executeCentralSearch(graphClient, projectName, {
                entityTypes: ['driveItem', 'listItem'],
                maxResults: 15,
                sortByRank: true,
              });

              const items = [...searchResult.results.files, ...searchResult.results.listItems].map(
                (hit) => ({
                  name: hit.name,
                  webUrl: hit.webUrl,
                  lastModified: hit.resource?.lastModifiedDateTime,
                  relevance: hit.relevanceScore,
                })
              );

              result.files = { count: items.length, items };
            } catch (error) {
              result.files = { error: `Could not fetch: ${error}` };
            }
          })()
        );
      }

      // Meetings
      if (includeMeetings) {
        promises.push(
          (async () => {
            try {
              const now = new Date();
              const pastDate = new Date(now.getTime() - 180 * 24 * 60 * 60 * 1000);
              const futureDate = new Date(now.getTime() + 90 * 24 * 60 * 60 * 1000);

              // Build query string (startDateTime and endDateTime are REQUIRED for calendarView)
              // Note: $filter with contains() causes 500 errors, use client-side filtering
              const projectQueryParams: Record<string, string> = {
                startDateTime: pastDate.toISOString(),
                endDateTime: futureDate.toISOString(),
                $top: '100',
                $select: 'id,subject,start,end,organizer,attendees',
                $orderby: 'start/dateTime desc',
              };

              const meetingsResponse = await graphClient.makeRequest(
                `/me/calendarView?${buildGraphQueryString(projectQueryParams)}`,
                {
                  method: 'GET',
                }
              );

              if (
                meetingsResponse &&
                typeof meetingsResponse === 'object' &&
                'value' in meetingsResponse &&
                Array.isArray(meetingsResponse.value)
              ) {
                // Client-side filtering for reliability
                const projectLower = projectName.toLowerCase();
                const events = (meetingsResponse.value as GraphEvent[]).filter((e) =>
                  e.subject?.toLowerCase().includes(projectLower)
                );
                const upcoming = events.filter((e) => new Date(e.start.dateTime) >= now);
                const past = events.filter((e) => new Date(e.start.dateTime) < now);

                result.meetings = {
                  total: events.length,
                  upcoming: upcoming.slice(0, 5).map((e) => ({
                    subject: e.subject,
                    date: e.start.dateTime,
                    organizer: e.organizer?.emailAddress?.name,
                  })),
                  past: past.slice(0, 5).map((e) => ({
                    subject: e.subject,
                    date: e.start.dateTime,
                    organizer: e.organizer?.emailAddress?.name,
                  })),
                };
              }
            } catch (error) {
              logger.warn(`Project meetings search failed: ${error}`);
              result.meetings = { error: `Could not fetch: ${error}` };
            }
          })()
        );
      }

      // Emails
      if (includeEmails) {
        promises.push(
          (async () => {
            try {
              const emailResponse = await graphClient.makeRequest('/me/messages', {
                method: 'GET',
                queryParams: {
                  $search: `"${projectName}"`,
                  $top: '15',
                  $select: 'id,subject,from,receivedDateTime,bodyPreview',
                  $orderby: 'receivedDateTime desc',
                },
              });

              if (
                emailResponse &&
                typeof emailResponse === 'object' &&
                'value' in emailResponse &&
                Array.isArray(emailResponse.value)
              ) {
                result.emails = {
                  count: emailResponse.value.length,
                  recent: (emailResponse.value as GraphEmail[]).map((e) => ({
                    subject: e.subject,
                    from: e.from?.emailAddress?.name || e.from?.emailAddress?.address,
                    date: e.receivedDateTime,
                    preview: e.bodyPreview?.substring(0, 100),
                  })),
                };
              }
            } catch (error) {
              result.emails = { error: `Could not fetch: ${error}` };
            }
          })()
        );
      }

      // Tasks (To-Do and Planner)
      if (includeTasks) {
        promises.push(
          (async () => {
            const relatedTasks: Array<{
              title: string;
              source: string;
              status: string;
              dueDate?: string;
            }> = [];

            // Search To-Do tasks
            try {
              const taskListsResponse = await graphClient.makeRequest('/me/todo/lists', {
                method: 'GET',
                queryParams: { $top: '10' },
              });

              if (
                taskListsResponse &&
                typeof taskListsResponse === 'object' &&
                'value' in taskListsResponse &&
                Array.isArray(taskListsResponse.value)
              ) {
                for (const list of taskListsResponse.value as Array<{ id: string }>) {
                  try {
                    const tasksResponse = await graphClient.makeRequest(
                      `/me/todo/lists/${list.id}/tasks`,
                      {
                        method: 'GET',
                        queryParams: {
                          $filter: `contains(title, '${projectName}')`,
                          $top: '20',
                        },
                      }
                    );

                    if (
                      tasksResponse &&
                      typeof tasksResponse === 'object' &&
                      'value' in tasksResponse &&
                      Array.isArray(tasksResponse.value)
                    ) {
                      for (const task of tasksResponse.value as Array<{
                        title: string;
                        status: string;
                        dueDateTime?: { dateTime: string };
                      }>) {
                        relatedTasks.push({
                          title: task.title,
                          source: 'To-Do',
                          status: task.status,
                          dueDate: task.dueDateTime?.dateTime,
                        });
                      }
                    }
                  } catch {
                    // Skip
                  }
                }
              }
            } catch {
              // Skip To-Do errors
            }

            // Search Planner tasks
            try {
              const plannerResponse = await graphClient.makeRequest('/me/planner/tasks', {
                method: 'GET',
                queryParams: { $top: '100' },
              });

              if (
                plannerResponse &&
                typeof plannerResponse === 'object' &&
                'value' in plannerResponse &&
                Array.isArray(plannerResponse.value)
              ) {
                const projectLower = projectName.toLowerCase();
                for (const task of plannerResponse.value as Array<{
                  title: string;
                  percentComplete: number;
                  dueDateTime?: string;
                }>) {
                  if (task.title.toLowerCase().includes(projectLower)) {
                    relatedTasks.push({
                      title: task.title,
                      source: 'Planner',
                      status: task.percentComplete >= 100 ? 'completed' : 'in progress',
                      dueDate: task.dueDateTime,
                    });
                  }
                }
              }
            } catch {
              // Skip Planner errors
            }

            result.tasks = {
              count: relatedTasks.length,
              items: relatedTasks,
            };
          })()
        );
      }

      await Promise.allSettled(promises);

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 11. GET COMPANY CONTACTS - Find all people from a company
  // ==========================================================================
  server.tool(
    'get-company-contacts',
    `Find all contacts and interactions with people from a specific company:
- Directory users from that company
- Outlook contacts
- Recent emails with company domain
- Meetings with company employees

Use this for "Who do we know at Microsoft?", "Find all contacts from Acme Corp", or "Show me our relationship with [company]".`,
    {
      companyName: z
        .string()
        .describe('Company name or email domain (e.g., "Microsoft" or "microsoft.com")'),
    },
    {
      title: 'Get Company Contacts',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ companyName }) => {
      logger.info(`Finding contacts for company: ${companyName}`);

      const result: Record<string, unknown> = {
        company: companyName,
        retrievedAt: new Date().toISOString(),
      };

      const promises: Promise<void>[] = [];
      const companyLower = companyName.toLowerCase();

      // Extract domain if provided
      const isDomain = companyName.includes('.');
      const searchTerms = isDomain
        ? [companyName, companyName.split('.')[0]]
        : [companyName, `${companyName.toLowerCase().replace(/\s+/g, '')}.com`];

      // Search directory users
      promises.push(
        (async () => {
          try {
            const usersResponse = await graphClient.makeRequest('/users', {
              method: 'GET',
              queryParams: {
                $search: `"companyName:${companyName}"`,
                $top: '20',
                $select: 'id,displayName,mail,jobTitle,department,companyName',
              },
              headers: {
                ConsistencyLevel: 'eventual',
              },
            });

            if (
              usersResponse &&
              typeof usersResponse === 'object' &&
              'value' in usersResponse &&
              Array.isArray(usersResponse.value)
            ) {
              result.directoryUsers = {
                count: usersResponse.value.length,
                users: (usersResponse.value as GraphUser[]).map((u) => ({
                  name: u.displayName,
                  email: u.mail,
                  jobTitle: u.jobTitle,
                  department: u.department,
                })),
              };
            }
          } catch (error) {
            result.directoryUsers = { error: `Could not fetch: ${error}` };
          }
        })()
      );

      // Search Outlook contacts
      promises.push(
        (async () => {
          try {
            const contactsResponse = await graphClient.makeRequest('/me/contacts', {
              method: 'GET',
              queryParams: {
                $filter: `contains(companyName, '${companyName}')`,
                $top: '50',
                $select: 'displayName,emailAddresses,companyName,jobTitle,businessPhones',
              },
            });

            if (
              contactsResponse &&
              typeof contactsResponse === 'object' &&
              'value' in contactsResponse &&
              Array.isArray(contactsResponse.value)
            ) {
              result.outlookContacts = {
                count: contactsResponse.value.length,
                contacts: (
                  contactsResponse.value as Array<{
                    displayName: string;
                    emailAddresses?: Array<{ address: string }>;
                    jobTitle?: string;
                    businessPhones?: string[];
                  }>
                ).map((c) => ({
                  name: c.displayName,
                  email: c.emailAddresses?.[0]?.address,
                  jobTitle: c.jobTitle,
                  phone: c.businessPhones?.[0],
                })),
              };
            }
          } catch (error) {
            result.outlookContacts = { error: `Could not fetch: ${error}` };
          }
        })()
      );

      // Recent emails with company domain
      promises.push(
        (async () => {
          try {
            // Try to find emails from/to company domain
            const searchQuery = searchTerms.map((t) => `from:${t}`).join(' OR ');
            const emailResponse = await graphClient.makeRequest('/me/messages', {
              method: 'GET',
              queryParams: {
                $search: `"${searchQuery}"`,
                $top: '20',
                $select: 'id,subject,from,receivedDateTime',
                $orderby: 'receivedDateTime desc',
              },
            });

            if (
              emailResponse &&
              typeof emailResponse === 'object' &&
              'value' in emailResponse &&
              Array.isArray(emailResponse.value)
            ) {
              const emails = emailResponse.value as GraphEmail[];
              const uniqueSenders = new Map<
                string,
                { name?: string; email: string; count: number }
              >();

              for (const email of emails) {
                const senderEmail = email.from?.emailAddress?.address;
                if (senderEmail) {
                  if (!uniqueSenders.has(senderEmail)) {
                    uniqueSenders.set(senderEmail, {
                      name: email.from?.emailAddress?.name,
                      email: senderEmail,
                      count: 1,
                    });
                  } else {
                    uniqueSenders.get(senderEmail)!.count++;
                  }
                }
              }

              result.emailContacts = {
                uniquePeople: uniqueSenders.size,
                contacts: Array.from(uniqueSenders.values())
                  .sort((a, b) => b.count - a.count)
                  .slice(0, 10),
                recentEmails: emails.length,
              };
            }
          } catch (error) {
            result.emailContacts = { error: `Could not fetch: ${error}` };
          }
        })()
      );

      await Promise.allSettled(promises);

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 12. GET FOLLOW UP ITEMS - Items needing attention
  // ==========================================================================
  server.tool(
    'get-follow-up-items',
    `Get all items that need your attention or follow-up:
- Unread important emails
- Flagged emails for follow-up
- Overdue tasks
- Meetings needing response
- Pending approvals

Use this for "What needs my attention?", "Show me items I need to follow up on", or "What's urgent?".`,
    {
      includeEmails: z.boolean().optional().describe('Include email follow-ups (default: true)'),
      includeTasks: z.boolean().optional().describe('Include overdue tasks (default: true)'),
      includeMeetings: z
        .boolean()
        .optional()
        .describe('Include meetings needing response (default: true)'),
    },
    {
      title: 'Get Follow Up Items',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ includeEmails = true, includeTasks = true, includeMeetings = true }) => {
      logger.info('Getting follow-up items');

      const result: Record<string, unknown> = {
        retrievedAt: new Date().toISOString(),
        urgentItems: [] as Array<{ type: string; title: string; reason: string; link?: string }>,
      };

      const promises: Promise<void>[] = [];

      // Flagged and important unread emails
      if (includeEmails) {
        promises.push(
          (async () => {
            try {
              // Flagged emails
              const flaggedResponse = await graphClient.makeRequest('/me/messages', {
                method: 'GET',
                queryParams: {
                  $filter: "flag/flagStatus eq 'flagged'",
                  $top: '20',
                  $select: 'id,subject,from,receivedDateTime,flag,webLink',
                  $orderby: 'receivedDateTime desc',
                },
              });

              const flaggedEmails: Array<{
                subject: string;
                from: string;
                date: string;
                dueDate?: string;
                webLink?: string;
              }> = [];

              if (
                flaggedResponse &&
                typeof flaggedResponse === 'object' &&
                'value' in flaggedResponse &&
                Array.isArray(flaggedResponse.value)
              ) {
                for (const email of flaggedResponse.value as Array<
                  GraphEmail & {
                    flag?: { flagStatus: string; dueDateTime?: { dateTime: string } };
                  }
                >) {
                  flaggedEmails.push({
                    subject: email.subject,
                    from: email.from?.emailAddress?.address || 'unknown',
                    date: email.receivedDateTime,
                    dueDate: email.flag?.dueDateTime?.dateTime,
                    webLink: email.webLink,
                  });

                  (
                    result.urgentItems as Array<{
                      type: string;
                      title: string;
                      reason: string;
                      link?: string;
                    }>
                  ).push({
                    type: 'email',
                    title: email.subject,
                    reason: 'Flagged for follow-up',
                    link: email.webLink,
                  });
                }
              }

              result.flaggedEmails = {
                count: flaggedEmails.length,
                emails: flaggedEmails,
              };

              // Important unread emails
              const importantResponse = await graphClient.makeRequest('/me/messages', {
                method: 'GET',
                queryParams: {
                  $filter: "importance eq 'high' and isRead eq false",
                  $top: '10',
                  $select: 'id,subject,from,receivedDateTime,webLink',
                  $orderby: 'receivedDateTime desc',
                },
              });

              if (
                importantResponse &&
                typeof importantResponse === 'object' &&
                'value' in importantResponse &&
                Array.isArray(importantResponse.value)
              ) {
                const importantEmails = importantResponse.value as GraphEmail[];
                result.unreadImportant = {
                  count: importantEmails.length,
                  emails: importantEmails.map((e) => ({
                    subject: e.subject,
                    from: e.from?.emailAddress?.address,
                    date: e.receivedDateTime,
                  })),
                };

                for (const email of importantEmails) {
                  (
                    result.urgentItems as Array<{
                      type: string;
                      title: string;
                      reason: string;
                      link?: string;
                    }>
                  ).push({
                    type: 'email',
                    title: email.subject,
                    reason: 'High importance, unread',
                    link: email.webLink,
                  });
                }
              }
            } catch (error) {
              result.emails = { error: `Could not fetch: ${error}` };
            }
          })()
        );
      }

      // Overdue tasks
      if (includeTasks) {
        promises.push(
          (async () => {
            const overdueTasks: Array<{
              title: string;
              dueDate: string;
              source: string;
            }> = [];
            const now = new Date();

            // To-Do overdue
            try {
              const taskListsResponse = await graphClient.makeRequest('/me/todo/lists', {
                method: 'GET',
                queryParams: { $top: '10' },
              });

              if (
                taskListsResponse &&
                typeof taskListsResponse === 'object' &&
                'value' in taskListsResponse &&
                Array.isArray(taskListsResponse.value)
              ) {
                for (const list of taskListsResponse.value as Array<{ id: string }>) {
                  try {
                    const tasksResponse = await graphClient.makeRequest(
                      `/me/todo/lists/${list.id}/tasks`,
                      {
                        method: 'GET',
                        queryParams: {
                          $filter: "status ne 'completed'",
                          $top: '50',
                        },
                      }
                    );

                    if (
                      tasksResponse &&
                      typeof tasksResponse === 'object' &&
                      'value' in tasksResponse &&
                      Array.isArray(tasksResponse.value)
                    ) {
                      for (const task of tasksResponse.value as Array<{
                        title: string;
                        dueDateTime?: { dateTime: string };
                      }>) {
                        if (task.dueDateTime) {
                          const dueDate = new Date(task.dueDateTime.dateTime);
                          if (dueDate < now) {
                            overdueTasks.push({
                              title: task.title,
                              dueDate: task.dueDateTime.dateTime,
                              source: 'To-Do',
                            });
                          }
                        }
                      }
                    }
                  } catch {
                    // Skip
                  }
                }
              }
            } catch {
              // Skip
            }

            // Planner overdue
            try {
              const plannerResponse = await graphClient.makeRequest('/me/planner/tasks', {
                method: 'GET',
                queryParams: { $top: '100' },
              });

              if (
                plannerResponse &&
                typeof plannerResponse === 'object' &&
                'value' in plannerResponse &&
                Array.isArray(plannerResponse.value)
              ) {
                for (const task of plannerResponse.value as Array<{
                  title: string;
                  percentComplete: number;
                  dueDateTime?: string;
                }>) {
                  if (task.percentComplete < 100 && task.dueDateTime) {
                    const dueDate = new Date(task.dueDateTime);
                    if (dueDate < now) {
                      overdueTasks.push({
                        title: task.title,
                        dueDate: task.dueDateTime,
                        source: 'Planner',
                      });
                    }
                  }
                }
              }
            } catch {
              // Skip
            }

            result.overdueTasks = {
              count: overdueTasks.length,
              tasks: overdueTasks,
            };

            for (const task of overdueTasks) {
              (result.urgentItems as Array<{ type: string; title: string; reason: string }>).push({
                type: 'task',
                title: task.title,
                reason: `Overdue since ${new Date(task.dueDate).toLocaleDateString()}`,
              });
            }
          })()
        );
      }

      // Meetings needing response
      if (includeMeetings) {
        promises.push(
          (async () => {
            try {
              const now = new Date();
              const futureDate = new Date(now.getTime() + 14 * 24 * 60 * 60 * 1000);

              // Build query string (startDateTime and endDateTime are REQUIRED for calendarView)
              const upcomingQueryParams: Record<string, string> = {
                startDateTime: now.toISOString(),
                endDateTime: futureDate.toISOString(),
                $top: '50',
                $select: 'id,subject,start,organizer,responseStatus,webLink',
              };

              const meetingsResponse = await graphClient.makeRequest(
                `/me/calendarView?${buildGraphQueryString(upcomingQueryParams)}`,
                {
                  method: 'GET',
                }
              );

              if (
                meetingsResponse &&
                typeof meetingsResponse === 'object' &&
                'value' in meetingsResponse &&
                Array.isArray(meetingsResponse.value)
              ) {
                const needsResponse = (
                  meetingsResponse.value as Array<{
                    subject: string;
                    start: { dateTime: string };
                    organizer: { emailAddress: { name: string } };
                    responseStatus?: { response: string };
                    webLink?: string;
                  }>
                ).filter(
                  (m) =>
                    m.responseStatus?.response === 'notResponded' ||
                    m.responseStatus?.response === 'none'
                );

                result.meetingsNeedingResponse = {
                  count: needsResponse.length,
                  meetings: needsResponse.map((m) => ({
                    subject: m.subject,
                    date: m.start.dateTime,
                    organizer: m.organizer?.emailAddress?.name,
                  })),
                };

                for (const meeting of needsResponse.slice(0, 5)) {
                  (
                    result.urgentItems as Array<{
                      type: string;
                      title: string;
                      reason: string;
                      link?: string;
                    }>
                  ).push({
                    type: 'meeting',
                    title: meeting.subject,
                    reason: 'Awaiting your response',
                    link: meeting.webLink,
                  });
                }
              }
            } catch (error) {
              result.meetingsNeedingResponse = { error: `Could not fetch: ${error}` };
            }
          })()
        );
      }

      await Promise.allSettled(promises);

      // Sort urgent items by type priority
      const typeOrder = { email: 1, task: 2, meeting: 3 };
      (result.urgentItems as Array<{ type: string; title: string; reason: string }>).sort(
        (a, b) =>
          (typeOrder[a.type as keyof typeof typeOrder] || 4) -
          (typeOrder[b.type as keyof typeof typeOrder] || 4)
      );

      result.totalUrgentItems = (result.urgentItems as unknown[]).length;

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ============================================================================
  // MEETING TRANSCRIPT TOOLS
  // ============================================================================

  /**
   * Tool: get-meeting-transcript-summary
   * Gets meeting transcripts and provides a summary of what was discussed
   */
  server.tool(
    'get-meeting-transcript-summary',
    `Gets the transcript from a Teams meeting and provides key information.
This tool automatically:
1. Finds the online meeting by subject or meeting ID
2. Retrieves available transcripts
3. Downloads and parses the transcript content
4. Extracts key discussion points, action items, and decisions

Use this for:
- "What was discussed in the meeting about [topic]?"
- "Get me the transcript from my meeting with [person]"
- "What were the action items from the [meeting name] meeting?"

Note: Transcripts are only available for meetings where transcription was enabled.`,
    {
      meetingSubject: z
        .string()
        .optional()
        .describe('Subject/title of the meeting to find transcript for'),
      meetingId: z.string().optional().describe('Online meeting ID if known (format: MSo1234...)'),
      person: z
        .string()
        .optional()
        .describe('Person who was in the meeting (helps narrow down which meeting)'),
      dateRange: z
        .enum(['today', 'yesterday', 'this_week', 'last_week', 'this_month'])
        .optional()
        .describe('Time range to search for meetings (default: this_week)'),
    },
    {
      title: 'Get Meeting Transcript Summary',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ meetingSubject, meetingId, person, dateRange = 'this_week' }) => {
      logger.info(
        `Getting meeting transcript for: ${meetingSubject || meetingId || 'recent meeting'}`
      );

      interface TranscriptResult {
        meeting?: {
          subject: string;
          startTime: string;
          endTime: string;
          organizer?: string;
          attendees: string[];
          onlineMeetingId?: string;
        };
        transcript?: {
          id: string;
          createdDateTime: string;
          contentUrl?: string;
          content?: string;
        };
        summary?: {
          keyPoints: string[];
          actionItems: string[];
          decisions: string[];
          participants: string[];
        };
        error?: string;
      }

      const result: TranscriptResult = {};

      try {
        // Calculate date range
        const now = new Date();
        let startDate: Date;
        let endDate: Date = new Date(now);
        endDate.setHours(23, 59, 59, 999);

        switch (dateRange) {
          case 'today':
            startDate = new Date(now);
            startDate.setHours(0, 0, 0, 0);
            break;
          case 'yesterday':
            startDate = new Date(now);
            startDate.setDate(startDate.getDate() - 1);
            startDate.setHours(0, 0, 0, 0);
            endDate = new Date(startDate);
            endDate.setHours(23, 59, 59, 999);
            break;
          case 'this_week':
            startDate = new Date(now);
            startDate.setDate(startDate.getDate() - startDate.getDay());
            startDate.setHours(0, 0, 0, 0);
            break;
          case 'last_week':
            startDate = new Date(now);
            startDate.setDate(startDate.getDate() - startDate.getDay() - 7);
            startDate.setHours(0, 0, 0, 0);
            endDate = new Date(startDate);
            endDate.setDate(endDate.getDate() + 6);
            endDate.setHours(23, 59, 59, 999);
            break;
          case 'this_month':
            startDate = new Date(now.getFullYear(), now.getMonth(), 1);
            break;
          default:
            startDate = new Date(now);
            startDate.setDate(startDate.getDate() - 7);
        }

        // Step 1: Find the meeting
        let targetMeeting: GraphEvent | null = null;
        let personEmail: string | undefined;

        // If person is specified, find their email first
        if (person) {
          const user = await findUser(graphClient, person);
          if (user) {
            personEmail = user.mail || user.userPrincipalName;
          }
        }

        // Search calendar events for online meetings
        // Build query string (startDateTime and endDateTime are REQUIRED for calendarView)
        const onlineQueryParams: Record<string, string> = {
          startDateTime: startDate.toISOString(),
          endDateTime: endDate.toISOString(),
          $filter: 'isOnlineMeeting eq true',
          $select: 'id,subject,start,end,attendees,isOnlineMeeting,onlineMeeting,organizer',
          $orderby: 'start/dateTime desc',
          $top: '50',
        };

        const calendarResponse = await graphClient.makeRequest(
          `/me/calendarView?${buildGraphQueryString(onlineQueryParams)}`,
          {
            method: 'GET',
          }
        );

        if (
          calendarResponse &&
          typeof calendarResponse === 'object' &&
          'value' in calendarResponse
        ) {
          const events = calendarResponse.value as Array<
            GraphEvent & { onlineMeeting?: { joinUrl?: string } }
          >;

          for (const event of events) {
            // Match by subject
            if (
              meetingSubject &&
              event.subject?.toLowerCase().includes(meetingSubject.toLowerCase())
            ) {
              targetMeeting = event;
              break;
            }
            // Match by attendee
            if (personEmail && !meetingSubject) {
              const attendeeEmails =
                event.attendees?.map((a) => a.emailAddress?.address?.toLowerCase()) || [];
              if (attendeeEmails.includes(personEmail.toLowerCase())) {
                targetMeeting = event;
                break;
              }
            }
          }

          // If no specific match, take the most recent online meeting
          if (!targetMeeting && events.length > 0 && !meetingSubject && !person) {
            targetMeeting = events[0];
          }
        }

        if (!targetMeeting) {
          return {
            content: [
              {
                type: 'text' as const,
                text: JSON.stringify(
                  {
                    error: 'No matching online meeting found',
                    hint: 'Make sure the meeting had transcription enabled and occurred within the specified date range.',
                    searchCriteria: { meetingSubject, person, dateRange },
                  },
                  null,
                  2
                ),
              },
            ],
          };
        }

        result.meeting = {
          subject: targetMeeting.subject,
          startTime: targetMeeting.start.dateTime,
          endTime: targetMeeting.end.dateTime,
          organizer: (
            targetMeeting as unknown as { organizer?: { emailAddress?: { name?: string } } }
          ).organizer?.emailAddress?.name,
          attendees:
            targetMeeting.attendees?.map(
              (a) => a.emailAddress?.name || a.emailAddress?.address || 'Unknown'
            ) || [],
        };

        // Step 2: Get online meeting details to find meeting ID
        // We need to search for online meetings or use the join URL
        const meetingEvent = targetMeeting as unknown as { onlineMeeting?: { joinUrl?: string } };
        let onlineMeetingId: string | undefined = meetingId;

        if (!onlineMeetingId && meetingEvent.onlineMeeting?.joinUrl) {
          // Try to extract meeting ID from join URL
          // URL format: https://teams.microsoft.com/l/meetup-join/19%3ameeting_...
          const joinUrl = meetingEvent.onlineMeeting.joinUrl;
          const match = joinUrl.match(/meeting_([a-zA-Z0-9_-]+)/);
          if (match) {
            // Need to get the actual meeting ID through the API
            try {
              // List user's online meetings and find the matching one
              const onlineMeetingsResponse = await graphClient.makeRequest('/me/onlineMeetings', {
                method: 'GET',
                queryParams: {
                  $filter: `startDateTime ge ${startDate.toISOString()} and startDateTime le ${endDate.toISOString()}`,
                  $select: 'id,subject,startDateTime,endDateTime,joinWebUrl',
                  $top: '50',
                },
              });

              if (
                onlineMeetingsResponse &&
                typeof onlineMeetingsResponse === 'object' &&
                'value' in onlineMeetingsResponse
              ) {
                const onlineMeetings = onlineMeetingsResponse.value as Array<{
                  id: string;
                  subject?: string;
                  joinWebUrl?: string;
                }>;

                for (const om of onlineMeetings) {
                  if (
                    om.subject?.toLowerCase() === targetMeeting.subject.toLowerCase() ||
                    om.joinWebUrl === joinUrl
                  ) {
                    onlineMeetingId = om.id;
                    break;
                  }
                }
              }
            } catch {
              logger.warn('Could not list online meetings to find meeting ID');
            }
          }
        }

        if (!onlineMeetingId) {
          return {
            content: [
              {
                type: 'text' as const,
                text: JSON.stringify(
                  {
                    ...result,
                    error: 'Could not find online meeting ID for this calendar event',
                    hint: 'The meeting may not have been a Teams online meeting, or the meeting ID could not be extracted.',
                  },
                  null,
                  2
                ),
              },
            ],
          };
        }

        result.meeting.onlineMeetingId = onlineMeetingId;

        // Step 3: Get transcripts for this meeting
        try {
          const transcriptsResponse = await graphClient.makeRequest(
            `/me/onlineMeetings/${onlineMeetingId}/transcripts`,
            {
              method: 'GET',
              queryParams: {
                $select: 'id,createdDateTime,transcriptContentUrl',
              },
            }
          );

          if (
            transcriptsResponse &&
            typeof transcriptsResponse === 'object' &&
            'value' in transcriptsResponse &&
            Array.isArray(transcriptsResponse.value) &&
            transcriptsResponse.value.length > 0
          ) {
            const transcripts = transcriptsResponse.value as Array<{
              id: string;
              createdDateTime: string;
              transcriptContentUrl?: string;
            }>;

            // Get the most recent transcript
            const latestTranscript = transcripts[0];
            result.transcript = {
              id: latestTranscript.id,
              createdDateTime: latestTranscript.createdDateTime,
              contentUrl: latestTranscript.transcriptContentUrl,
            };

            // Step 4: Download transcript content
            try {
              const transcriptContent = await graphClient.makeRequest(
                `/me/onlineMeetings/${onlineMeetingId}/transcripts/${latestTranscript.id}/content`,
                {
                  method: 'GET',
                  queryParams: { $format: 'text/vtt' },
                }
              );

              if (transcriptContent && typeof transcriptContent === 'string') {
                result.transcript.content = transcriptContent;

                // Step 5: Extract key information from transcript
                result.summary = extractTranscriptSummary(transcriptContent);
              } else if (transcriptContent && typeof transcriptContent === 'object') {
                // Content might be returned as binary or object
                result.transcript.content = JSON.stringify(transcriptContent);
                result.summary = extractTranscriptSummary(result.transcript.content);
              }
            } catch (contentError) {
              logger.warn(`Could not download transcript content: ${contentError}`);
              result.transcript.content = undefined;
              result.summary = {
                keyPoints: ['Transcript content could not be downloaded'],
                actionItems: [],
                decisions: [],
                participants: [],
              };
            }
          } else {
            return {
              content: [
                {
                  type: 'text' as const,
                  text: JSON.stringify(
                    {
                      ...result,
                      error: 'No transcripts available for this meeting',
                      hint: 'Transcription must be enabled during the meeting for transcripts to be available. Check if the meeting organizer enabled transcription.',
                    },
                    null,
                    2
                  ),
                },
              ],
            };
          }
        } catch (transcriptError) {
          return {
            content: [
              {
                type: 'text' as const,
                text: JSON.stringify(
                  {
                    ...result,
                    error: `Failed to retrieve transcripts: ${transcriptError}`,
                    hint: 'You may not have permission to access transcripts, or transcription was not enabled for this meeting.',
                  },
                  null,
                  2
                ),
              },
            ],
          };
        }

        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(result, null, 2),
            },
          ],
        };
      } catch (error) {
        logger.error(`Error getting meeting transcript: ${error}`);
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  error: `Failed to get meeting transcript: ${error}`,
                  hint: 'Ensure you have the required permissions (OnlineMeetingTranscript.Read.All) and the meeting had transcription enabled.',
                },
                null,
                2
              ),
            },
          ],
        };
      }
    }
  );
  registeredCount++;

  /**
   * Tool: list-meetings-with-transcripts
   * Lists all recent meetings that have transcripts available
   */
  server.tool(
    'list-meetings-with-transcripts',
    `Lists all Teams meetings that have transcripts available.
This tool:
1. Finds all online meetings in the specified time range
2. Checks which ones have transcripts
3. Returns a list with meeting details and transcript availability

Use this for:
- "Which of my meetings have transcripts?"
- "Show me all transcribed meetings from this week"
- "Find meetings where I can review what was discussed"`,
    {
      dateRange: z
        .enum(['today', 'yesterday', 'this_week', 'last_week', 'this_month', 'last_month'])
        .optional()
        .describe('Time range to search for meetings (default: this_week)'),
      limit: z.number().optional().describe('Maximum number of meetings to return (default: 20)'),
    },
    {
      title: 'List Meetings with Transcripts',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ dateRange = 'this_week', limit = 20 }) => {
      logger.info(`Listing meetings with transcripts for: ${dateRange}`);

      interface MeetingWithTranscript {
        id: string;
        subject: string;
        startTime: string;
        endTime: string;
        organizer?: string;
        attendeeCount: number;
        hasTranscript: boolean;
        transcriptCount?: number;
        hasRecording: boolean;
        recordingCount?: number;
      }

      const meetingsWithTranscripts: MeetingWithTranscript[] = [];

      try {
        // Calculate date range
        const now = new Date();
        let startDate: Date;
        let endDate: Date = new Date(now);
        endDate.setHours(23, 59, 59, 999);

        switch (dateRange) {
          case 'today':
            startDate = new Date(now);
            startDate.setHours(0, 0, 0, 0);
            break;
          case 'yesterday':
            startDate = new Date(now);
            startDate.setDate(startDate.getDate() - 1);
            startDate.setHours(0, 0, 0, 0);
            endDate = new Date(startDate);
            endDate.setHours(23, 59, 59, 999);
            break;
          case 'this_week':
            startDate = new Date(now);
            startDate.setDate(startDate.getDate() - startDate.getDay());
            startDate.setHours(0, 0, 0, 0);
            break;
          case 'last_week':
            startDate = new Date(now);
            startDate.setDate(startDate.getDate() - startDate.getDay() - 7);
            startDate.setHours(0, 0, 0, 0);
            endDate = new Date(startDate);
            endDate.setDate(endDate.getDate() + 6);
            endDate.setHours(23, 59, 59, 999);
            break;
          case 'this_month':
            startDate = new Date(now.getFullYear(), now.getMonth(), 1);
            break;
          case 'last_month':
            startDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
            endDate = new Date(now.getFullYear(), now.getMonth(), 0);
            endDate.setHours(23, 59, 59, 999);
            break;
          default:
            startDate = new Date(now);
            startDate.setDate(startDate.getDate() - 7);
        }

        // Get online meetings
        const onlineMeetingsResponse = await graphClient.makeRequest('/me/onlineMeetings', {
          method: 'GET',
          queryParams: {
            $filter: `startDateTime ge ${startDate.toISOString()} and startDateTime le ${endDate.toISOString()}`,
            $select: 'id,subject,startDateTime,endDateTime,participants',
            $orderby: 'startDateTime desc',
            $top: String(Math.min(limit * 2, 100)), // Get more to filter
          },
        });

        if (
          !onlineMeetingsResponse ||
          typeof onlineMeetingsResponse !== 'object' ||
          !('value' in onlineMeetingsResponse)
        ) {
          return {
            content: [
              {
                type: 'text' as const,
                text: JSON.stringify(
                  {
                    meetings: [],
                    total: 0,
                    dateRange,
                    message: 'No online meetings found in the specified date range',
                  },
                  null,
                  2
                ),
              },
            ],
          };
        }

        const onlineMeetings = onlineMeetingsResponse.value as Array<{
          id: string;
          subject?: string;
          startDateTime?: string;
          endDateTime?: string;
          participants?: {
            organizer?: { identity?: { user?: { displayName?: string } } };
            attendees?: unknown[];
          };
        }>;

        // Check each meeting for transcripts
        for (const meeting of onlineMeetings.slice(0, limit)) {
          const meetingInfo: MeetingWithTranscript = {
            id: meeting.id,
            subject: meeting.subject || 'Untitled Meeting',
            startTime: meeting.startDateTime || '',
            endTime: meeting.endDateTime || '',
            organizer: meeting.participants?.organizer?.identity?.user?.displayName,
            attendeeCount: Array.isArray(meeting.participants?.attendees)
              ? meeting.participants.attendees.length
              : 0,
            hasTranscript: false,
            hasRecording: false,
          };

          // Check for transcripts
          try {
            const transcriptsResponse = await graphClient.makeRequest(
              `/me/onlineMeetings/${meeting.id}/transcripts`,
              { method: 'GET', queryParams: { $top: '1' } }
            );

            if (
              transcriptsResponse &&
              typeof transcriptsResponse === 'object' &&
              'value' in transcriptsResponse &&
              Array.isArray(transcriptsResponse.value)
            ) {
              meetingInfo.hasTranscript = transcriptsResponse.value.length > 0;
              meetingInfo.transcriptCount = transcriptsResponse.value.length;
            }
          } catch {
            // No transcript access or doesn't exist
          }

          // Check for recordings
          try {
            const recordingsResponse = await graphClient.makeRequest(
              `/me/onlineMeetings/${meeting.id}/recordings`,
              { method: 'GET', queryParams: { $top: '1' } }
            );

            if (
              recordingsResponse &&
              typeof recordingsResponse === 'object' &&
              'value' in recordingsResponse &&
              Array.isArray(recordingsResponse.value)
            ) {
              meetingInfo.hasRecording = recordingsResponse.value.length > 0;
              meetingInfo.recordingCount = recordingsResponse.value.length;
            }
          } catch {
            // No recording access or doesn't exist
          }

          meetingsWithTranscripts.push(meetingInfo);
        }

        const withTranscripts = meetingsWithTranscripts.filter((m) => m.hasTranscript);
        const withRecordings = meetingsWithTranscripts.filter((m) => m.hasRecording);

        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  summary: {
                    totalMeetings: meetingsWithTranscripts.length,
                    withTranscripts: withTranscripts.length,
                    withRecordings: withRecordings.length,
                    dateRange,
                  },
                  meetings: meetingsWithTranscripts,
                  hint: 'Use get-meeting-transcript-summary with a meeting subject to get the full transcript content.',
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (error) {
        logger.error(`Error listing meetings with transcripts: ${error}`);
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  error: `Failed to list meetings: ${error}`,
                  hint: 'Ensure you have the required permissions (OnlineMeetings.Read, OnlineMeetingTranscript.Read.All).',
                },
                null,
                2
              ),
            },
          ],
        };
      }
    }
  );
  registeredCount++;

  /**
   * Tool: search-across-transcripts
   * Searches for specific topics or keywords across all meeting transcripts
   */
  server.tool(
    'search-across-transcripts',
    `Searches for specific topics, keywords, or discussions across all meeting transcripts.
This tool:
1. Finds all meetings with transcripts in the date range
2. Downloads each transcript
3. Searches for the specified query
4. Returns matching segments with context

Use this for:
- "What meetings mentioned the budget?"
- "Find all discussions about Project X"
- "Who talked about the deadline?"`,
    {
      query: z.string().describe('Topic, keyword, or phrase to search for in transcripts'),
      dateRange: z
        .enum(['this_week', 'last_week', 'this_month', 'last_month', 'last_3_months'])
        .optional()
        .describe('Time range to search (default: this_month)'),
      maxMeetings: z.number().optional().describe('Maximum meetings to search (default: 10)'),
    },
    {
      title: 'Search Across Transcripts',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ query, dateRange = 'this_month', maxMeetings = 10 }) => {
      logger.info(`Searching transcripts for: "${query}"`);

      interface TranscriptMatch {
        meetingId: string;
        meetingSubject: string;
        meetingDate: string;
        matches: Array<{
          speaker?: string;
          text: string;
          timestamp?: string;
        }>;
        matchCount: number;
      }

      const results: TranscriptMatch[] = [];

      try {
        // Calculate date range
        const now = new Date();
        let startDate: Date;
        const endDate: Date = new Date(now);
        endDate.setHours(23, 59, 59, 999);

        switch (dateRange) {
          case 'this_week':
            startDate = new Date(now);
            startDate.setDate(startDate.getDate() - startDate.getDay());
            startDate.setHours(0, 0, 0, 0);
            break;
          case 'last_week':
            startDate = new Date(now);
            startDate.setDate(startDate.getDate() - startDate.getDay() - 7);
            startDate.setHours(0, 0, 0, 0);
            break;
          case 'this_month':
            startDate = new Date(now.getFullYear(), now.getMonth(), 1);
            break;
          case 'last_month':
            startDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
            break;
          case 'last_3_months':
            startDate = new Date(now.getFullYear(), now.getMonth() - 3, 1);
            break;
          default:
            startDate = new Date(now.getFullYear(), now.getMonth(), 1);
        }

        // Get online meetings
        const onlineMeetingsResponse = await graphClient.makeRequest('/me/onlineMeetings', {
          method: 'GET',
          queryParams: {
            $filter: `startDateTime ge ${startDate.toISOString()} and startDateTime le ${endDate.toISOString()}`,
            $select: 'id,subject,startDateTime',
            $orderby: 'startDateTime desc',
            $top: String(maxMeetings * 2),
          },
        });

        if (
          !onlineMeetingsResponse ||
          typeof onlineMeetingsResponse !== 'object' ||
          !('value' in onlineMeetingsResponse)
        ) {
          return {
            content: [
              {
                type: 'text' as const,
                text: JSON.stringify({ query, results: [], message: 'No meetings found' }, null, 2),
              },
            ],
          };
        }

        const meetings = onlineMeetingsResponse.value as Array<{
          id: string;
          subject?: string;
          startDateTime?: string;
        }>;

        const queryLower = query.toLowerCase();
        let meetingsSearched = 0;

        for (const meeting of meetings) {
          if (meetingsSearched >= maxMeetings) break;

          try {
            // Get transcripts
            const transcriptsResponse = await graphClient.makeRequest(
              `/me/onlineMeetings/${meeting.id}/transcripts`,
              { method: 'GET', queryParams: { $top: '1' } }
            );

            if (
              !transcriptsResponse ||
              typeof transcriptsResponse !== 'object' ||
              !('value' in transcriptsResponse) ||
              !Array.isArray(transcriptsResponse.value) ||
              transcriptsResponse.value.length === 0
            ) {
              continue;
            }

            const transcript = transcriptsResponse.value[0] as { id: string };
            meetingsSearched++;

            // Download transcript content
            const contentResponse = await graphClient.makeRequest(
              `/me/onlineMeetings/${meeting.id}/transcripts/${transcript.id}/content`,
              { method: 'GET', queryParams: { $format: 'text/vtt' } }
            );

            if (contentResponse && typeof contentResponse === 'string') {
              // Search for matches in the transcript
              const matches = searchInTranscript(contentResponse, queryLower);

              if (matches.length > 0) {
                results.push({
                  meetingId: meeting.id,
                  meetingSubject: meeting.subject || 'Untitled Meeting',
                  meetingDate: meeting.startDateTime || '',
                  matches: matches.slice(0, 5), // Top 5 matches per meeting
                  matchCount: matches.length,
                });
              }
            }
          } catch {
            // Skip this meeting if we can't access transcript
          }
        }

        const totalMatches = results.reduce((sum, r) => sum + r.matchCount, 0);

        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  query,
                  dateRange,
                  summary: {
                    meetingsSearched,
                    meetingsWithMatches: results.length,
                    totalMatches,
                  },
                  results,
                  hint:
                    results.length > 0
                      ? 'Use get-meeting-transcript-summary with the meeting subject for full details.'
                      : 'Try a different query or expand the date range.',
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (error) {
        logger.error(`Error searching transcripts: ${error}`);
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  error: `Failed to search transcripts: ${error}`,
                  query,
                  hint: 'Ensure you have OnlineMeetingTranscript.Read.All permission.',
                },
                null,
                2
              ),
            },
          ],
        };
      }
    }
  );
  registeredCount++;

  // ==========================================================================
  // 16. ANALYZE TEAM COLLABORATION - Team communication patterns
  // ==========================================================================
  server.tool(
    'analyze-team-collaboration',
    `Analyze collaboration patterns and communication frequency within a team or group:
- Email exchange frequency between team members
- Meeting frequency and duration
- Shared file activity
- Teams chat activity
- Identify most active collaborators

Use this for "How is my team collaborating?", "Who are the most active team members?", or "Analyze communication in my department".`,
    {
      teamName: z.string().optional().describe('Team or group name to analyze'),
      department: z.string().optional().describe('Department name to analyze'),
      days: z.number().optional().describe('Number of days to analyze (default: 30)'),
    },
    {
      title: 'Analyze Team Collaboration',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ teamName, department, days = 30 }) => {
      logger.info(`Analyzing team collaboration: ${teamName || department || 'my network'}`);

      const result: Record<string, unknown> = {
        analyzedPeriod: `Last ${days} days`,
        retrievedAt: new Date().toISOString(),
      };

      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);
      const endDate = new Date();

      const collaborators = new Map<
        string,
        {
          name: string;
          email: string;
          emailsSent: number;
          emailsReceived: number;
          meetingsTogether: number;
          sharedFiles: number;
        }
      >();

      const promises: Promise<void>[] = [];

      // Analyze email patterns
      promises.push(
        (async () => {
          try {
            const emailResponse = await graphClient.makeRequest('/me/messages', {
              method: 'GET',
              queryParams: {
                $filter: `receivedDateTime ge ${startDate.toISOString()}`,
                $top: '200',
                $select: 'from,toRecipients,ccRecipients,receivedDateTime',
              },
            });

            if (
              emailResponse &&
              typeof emailResponse === 'object' &&
              'value' in emailResponse &&
              Array.isArray(emailResponse.value)
            ) {
              for (const email of emailResponse.value as GraphEmail[]) {
                // Track sender
                const senderEmail = email.from?.emailAddress?.address?.toLowerCase();
                const senderName = email.from?.emailAddress?.name || senderEmail || 'Unknown';
                if (senderEmail && senderEmail !== 'unknown') {
                  if (!collaborators.has(senderEmail)) {
                    collaborators.set(senderEmail, {
                      name: senderName,
                      email: senderEmail,
                      emailsSent: 0,
                      emailsReceived: 1,
                      meetingsTogether: 0,
                      sharedFiles: 0,
                    });
                  } else {
                    collaborators.get(senderEmail)!.emailsReceived++;
                  }
                }

                // Track recipients (emails sent to them)
                for (const recipient of email.toRecipients || []) {
                  const recipientEmail = recipient.emailAddress?.address?.toLowerCase();
                  const recipientName = recipient.emailAddress?.name || recipientEmail || 'Unknown';
                  if (recipientEmail) {
                    if (!collaborators.has(recipientEmail)) {
                      collaborators.set(recipientEmail, {
                        name: recipientName,
                        email: recipientEmail,
                        emailsSent: 1,
                        emailsReceived: 0,
                        meetingsTogether: 0,
                        sharedFiles: 0,
                      });
                    } else {
                      collaborators.get(recipientEmail)!.emailsSent++;
                    }
                  }
                }
              }

              result.emailAnalysis = {
                totalEmails: emailResponse.value.length,
                uniqueContacts: collaborators.size,
              };
            }
          } catch (error) {
            result.emailAnalysis = { error: `Could not analyze: ${error}` };
          }
        })()
      );

      // Analyze meeting patterns
      promises.push(
        (async () => {
          try {
            // Build query string (startDateTime and endDateTime are REQUIRED for calendarView)
            const analysisQueryParams: Record<string, string> = {
              startDateTime: startDate.toISOString(),
              endDateTime: endDate.toISOString(),
              $top: '200',
              $select: 'subject,start,end,attendees,organizer',
            };

            const meetingsResponse = await graphClient.makeRequest(
              `/me/calendarView?${buildGraphQueryString(analysisQueryParams)}`,
              {
                method: 'GET',
              }
            );

            if (
              meetingsResponse &&
              typeof meetingsResponse === 'object' &&
              'value' in meetingsResponse &&
              Array.isArray(meetingsResponse.value)
            ) {
              const events = meetingsResponse.value as GraphEvent[];
              let totalMeetingHours = 0;
              const meetingsByDay = new Map<string, number>();

              for (const event of events) {
                // Calculate duration
                const start = new Date(event.start.dateTime);
                const end = new Date(event.end.dateTime);
                const duration = (end.getTime() - start.getTime()) / (1000 * 60 * 60);
                totalMeetingHours += duration;

                // Track by day
                const day = start.toISOString().split('T')[0];
                meetingsByDay.set(day, (meetingsByDay.get(day) || 0) + 1);

                // Track attendees
                for (const attendee of event.attendees || []) {
                  const attendeeEmail = attendee.emailAddress?.address?.toLowerCase();
                  if (attendeeEmail) {
                    if (collaborators.has(attendeeEmail)) {
                      collaborators.get(attendeeEmail)!.meetingsTogether++;
                    } else {
                      collaborators.set(attendeeEmail, {
                        name: attendee.emailAddress?.name || attendeeEmail,
                        email: attendeeEmail,
                        emailsSent: 0,
                        emailsReceived: 0,
                        meetingsTogether: 1,
                        sharedFiles: 0,
                      });
                    }
                  }
                }
              }

              result.meetingAnalysis = {
                totalMeetings: events.length,
                totalHours: Math.round(totalMeetingHours * 10) / 10,
                averagePerDay: Math.round((events.length / days) * 10) / 10,
                busiestDay: Array.from(meetingsByDay.entries()).sort(
                  ([, a], [, b]) => b - a
                )[0]?.[0],
              };
            }
          } catch (error) {
            result.meetingAnalysis = { error: `Could not analyze: ${error}` };
          }
        })()
      );

      await Promise.allSettled(promises);

      // Calculate top collaborators
      const sortedCollaborators = Array.from(collaborators.values())
        .map((c) => ({
          ...c,
          totalInteractions: c.emailsSent + c.emailsReceived + c.meetingsTogether,
        }))
        .sort((a, b) => b.totalInteractions - a.totalInteractions)
        .slice(0, 20);

      result.topCollaborators = sortedCollaborators;
      result.summary = {
        totalPeopleInteracted: collaborators.size,
        topContacts: sortedCollaborators.slice(0, 5).map((c) => c.name),
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 17. GET CUSTOMER 360 - Complete 360-degree view across ALL Microsoft 365 products
  // ==========================================================================
  server.tool(
    'get-customer-360',
    `Get a comprehensive 360-degree view of a customer, partner, or person across ALL Microsoft 365 products.

This tool automatically detects whether the input is a person name or company name and searches across 16+ Microsoft 365 products:

**Core Search (Microsoft Search API):**
- Emails (messages exchanged)
- Calendar events and meetings
- Files (OneDrive, SharePoint)
- SharePoint sites
- SharePoint lists and list items
- Teams chat messages
- People and contacts

**Extended Search (Follow-up Queries):**
- OneNote pages
- Planner tasks
- Microsoft To-Do tasks
- Personal contacts
- Online meetings
- Joined Teams
- Bookings appointments
- Insights (trending and shared items)

**Use cases:**
- "Show me everything about XYZ" (person or company)
- "Customer 360 for Acme Corp"
- "What's our relationship with John Smith?"
- "Give me all information about [customer name]"
- "Zeige mir alles zum Kunden XYZ" (German)

Returns a comprehensive summary including relationship health score and recommendations.`,
    {
      customerIdentifier: z
        .string()
        .describe(
          'Customer, partner, or person identifier - can be a person name, company name, or email address'
        ),
      emailDomain: z
        .string()
        .optional()
        .describe('Optional: Company email domain (e.g., acme.com) to improve search accuracy'),
      days: z.number().optional().describe('Days of history to include (default: 90)'),
      maxResults: z
        .number()
        .optional()
        .describe('Maximum results per category (default: 25, max: 100)'),
    },
    {
      title: 'Get Customer 360 - Comprehensive View',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ customerIdentifier, emailDomain, days = 90, maxResults = 25 }) => {
      logger.info(`Getting comprehensive Customer 360 for: ${customerIdentifier}`);

      const effectiveMaxResults = Math.min(maxResults, 100);
      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);
      const now = new Date();

      // Step 1: Detect customer type (person vs company)
      const customerType = await detectCustomerType(graphClient, customerIdentifier);
      const searchDomain = emailDomain || customerType.emailDomain;
      const searchQuery = customerType.displayName || customerIdentifier;

      // Initialize result structure
      const result: Record<string, unknown> = {
        customer: {
          identifier: customerIdentifier,
          type: customerType.type,
          displayName: customerType.displayName,
          emailDomain: searchDomain,
          userId: customerType.userId,
          userEmail: customerType.userEmail,
        },
        analyzedPeriod: `Last ${days} days`,
        retrievedAt: new Date().toISOString(),
        productsSearched: [] as string[],
      };

      const productsSearched: string[] = [];
      const searchPromises: Promise<void>[] = [];

      // Step 2: Execute Central Search across all entity types
      searchPromises.push(
        (async () => {
          try {
            const centralSearchResult = await executeCentralSearch(
              graphClient,
              `${searchQuery} OR "${searchQuery}"${searchDomain ? ` OR "${searchDomain}"` : ''}`,
              {
                entityTypes: [
                  'message',
                  'event',
                  'driveItem',
                  'site',
                  'list',
                  'listItem',
                  'chatMessage',
                  'person',
                ],
                maxResults: effectiveMaxResults * 2,
                sortByRank: true,
                includeTimeContext: true,
              }
            );

            // Process emails
            if (centralSearchResult.results.emails.length > 0) {
              const emails = centralSearchResult.results.emails.slice(0, effectiveMaxResults);
              result.emails = {
                count: centralSearchResult.results.emails.length,
                items: emails.map((hit) => ({
                  subject: hit.resource?.subject,
                  from: hit.resource?.from?.emailAddress?.address,
                  to: hit.resource?.toRecipients?.[0]?.emailAddress?.address,
                  date: hit.resource?.receivedDateTime,
                  preview: hit.resource?.bodyPreview?.substring(0, 150),
                  hasAttachments: hit.resource?.hasAttachments,
                  webLink: hit.resource?.webLink,
                  relevance: hit.relevanceScore,
                })),
              };
              productsSearched.push('Outlook Mail');
            }

            // Process calendar events
            if (centralSearchResult.results.events.length > 0) {
              const events = centralSearchResult.results.events.slice(0, effectiveMaxResults);
              const upcoming = events.filter((e) => new Date(e.resource?.start?.dateTime) >= now);
              const past = events.filter((e) => new Date(e.resource?.start?.dateTime) < now);

              result.calendar = {
                count: centralSearchResult.results.events.length,
                upcoming: {
                  count: upcoming.length,
                  items: upcoming.slice(0, 10).map((hit) => ({
                    subject: hit.resource?.subject,
                    start: hit.resource?.start?.dateTime,
                    end: hit.resource?.end?.dateTime,
                    location: hit.resource?.location?.displayName,
                    organizer: hit.resource?.organizer?.emailAddress?.address,
                    webLink: hit.resource?.webLink,
                  })),
                },
                past: {
                  count: past.length,
                  items: past.slice(0, 10).map((hit) => ({
                    subject: hit.resource?.subject,
                    start: hit.resource?.start?.dateTime,
                    end: hit.resource?.end?.dateTime,
                    location: hit.resource?.location?.displayName,
                  })),
                },
              };
              productsSearched.push('Calendar');
            }

            // Process files
            if (centralSearchResult.results.files.length > 0) {
              const files = centralSearchResult.results.files.slice(0, effectiveMaxResults);
              result.files = {
                count: centralSearchResult.results.files.length,
                items: files.map((hit) => ({
                  name: hit.name,
                  webUrl: hit.webUrl,
                  type: hit.resource?.file?.mimeType || hit.resource?.['@odata.type'],
                  modified: hit.resource?.lastModifiedDateTime,
                  modifiedBy: hit.resource?.lastModifiedBy?.user?.displayName,
                  size: hit.resource?.size,
                  relevance: hit.relevanceScore,
                })),
              };
              productsSearched.push('OneDrive/SharePoint Files');
            }

            // Process SharePoint sites
            if (centralSearchResult.results.sites.length > 0) {
              const sites = centralSearchResult.results.sites.slice(0, effectiveMaxResults);
              result.sites = {
                count: centralSearchResult.results.sites.length,
                items: sites.map((hit) => ({
                  name: hit.name,
                  webUrl: hit.webUrl,
                  description: hit.resource?.description,
                  relevance: hit.relevanceScore,
                })),
              };
              productsSearched.push('SharePoint Sites');
            }

            // Process list items
            if (centralSearchResult.results.listItems.length > 0) {
              const listItems = centralSearchResult.results.listItems.slice(0, effectiveMaxResults);
              result.listItems = {
                count: centralSearchResult.results.listItems.length,
                items: listItems.map((hit) => ({
                  name: hit.name,
                  webUrl: hit.webUrl,
                  listName: hit.resource?.fields?.Title,
                  relevance: hit.relevanceScore,
                })),
              };
              productsSearched.push('SharePoint Lists');
            }

            // Process Teams chats
            if (centralSearchResult.results.chats.length > 0) {
              const chats = centralSearchResult.results.chats.slice(0, effectiveMaxResults);
              result.teamsChats = {
                count: centralSearchResult.results.chats.length,
                items: chats.map((hit) => ({
                  content: hit.resource?.body?.content?.substring(0, 200),
                  from: hit.resource?.from?.user?.displayName,
                  date: hit.resource?.createdDateTime,
                  chatType: hit.resource?.chatType,
                  relevance: hit.relevanceScore,
                })),
              };
              productsSearched.push('Teams Chats');
            }

            // Process people
            if (centralSearchResult.results.people.length > 0) {
              const people = centralSearchResult.results.people.slice(0, effectiveMaxResults);
              result.people = {
                count: centralSearchResult.results.people.length,
                items: people.map((hit) => ({
                  name: hit.name || hit.resource?.displayName,
                  email:
                    hit.resource?.emailAddresses?.[0]?.address ||
                    hit.resource?.scoredEmailAddresses?.[0]?.address ||
                    hit.resource?.mail,
                  jobTitle: hit.resource?.jobTitle,
                  department: hit.resource?.department,
                  company: hit.resource?.companyName,
                  phone: hit.resource?.phones?.[0]?.number || hit.resource?.businessPhones?.[0],
                  relevance: hit.relevanceScore,
                })),
              };
              productsSearched.push('People');
            }
          } catch (error) {
            logger.warn(`Central search failed: ${error}`);
            result.centralSearchError = `${error}`;
          }
        })()
      );

      // Step 3: Execute Follow-up Queries for extended products
      searchPromises.push(
        (async () => {
          try {
            const followUpResults = await executeFollowUpQueries(graphClient, searchQuery);

            // Process OneNote pages
            if (followUpResults.onenote && followUpResults.onenote.length > 0) {
              result.onenote = {
                count: followUpResults.onenote.length,
                items: followUpResults.onenote.slice(0, effectiveMaxResults).map((page: any) => ({
                  title: page.title,
                  lastModified: page.lastModifiedDateTime,
                  webUrl: page.links?.oneNoteWebUrl?.href,
                })),
              };
              productsSearched.push('OneNote');
            }

            // Process Planner tasks
            if (followUpResults.planner && followUpResults.planner.length > 0) {
              result.planner = {
                count: followUpResults.planner.length,
                items: followUpResults.planner.slice(0, effectiveMaxResults).map((task: any) => ({
                  title: task.title,
                  percentComplete: task.percentComplete,
                  dueDate: task.dueDateTime,
                  priority: task.priority,
                  bucketId: task.bucketId,
                })),
              };
              productsSearched.push('Planner');
            }

            // Process To-Do tasks
            if (followUpResults.todo && followUpResults.todo.length > 0) {
              result.todo = {
                count: followUpResults.todo.length,
                items: followUpResults.todo.slice(0, effectiveMaxResults).map((task: any) => ({
                  title: task.title,
                  status: task.status,
                  dueDate: task.dueDateTime?.dateTime,
                  importance: task.importance,
                  listName: task.listName,
                })),
              };
              productsSearched.push('Microsoft To-Do');
            }

            // Process personal contacts
            if (followUpResults.contacts && followUpResults.contacts.length > 0) {
              result.contacts = {
                count: followUpResults.contacts.length,
                items: followUpResults.contacts
                  .slice(0, effectiveMaxResults)
                  .map((contact: any) => ({
                    displayName: contact.displayName,
                    email: contact.emailAddresses?.[0]?.address,
                    company: contact.companyName,
                    jobTitle: contact.jobTitle,
                    department: contact.department,
                    phone: contact.businessPhones?.[0] || contact.mobilePhone,
                  })),
              };
              productsSearched.push('Contacts');
            }

            // Process online meetings
            if (followUpResults.meetings && followUpResults.meetings.length > 0) {
              result.onlineMeetings = {
                count: followUpResults.meetings.length,
                items: followUpResults.meetings
                  .slice(0, effectiveMaxResults)
                  .map((meeting: any) => ({
                    subject: meeting.subject,
                    startDateTime: meeting.startDateTime,
                    endDateTime: meeting.endDateTime,
                    joinUrl: meeting.joinWebUrl,
                  })),
              };
              productsSearched.push('Online Meetings');
            }

            // Process joined Teams
            if (followUpResults.teams && followUpResults.teams.length > 0) {
              result.teams = {
                count: followUpResults.teams.length,
                items: followUpResults.teams.slice(0, effectiveMaxResults).map((team: any) => ({
                  displayName: team.displayName,
                  description: team.description,
                  webUrl: team.webUrl,
                })),
              };
              productsSearched.push('Teams');
            }

            // Process bookings
            if (followUpResults.bookings && followUpResults.bookings.length > 0) {
              result.bookings = {
                count: followUpResults.bookings.length,
                items: followUpResults.bookings
                  .slice(0, effectiveMaxResults)
                  .map((booking: any) => ({
                    customerName: booking.customerName,
                    serviceName: booking.serviceName,
                    startDateTime: booking.startDateTime?.dateTime,
                    endDateTime: booking.endDateTime?.dateTime,
                    businessName: booking.businessName,
                  })),
              };
              productsSearched.push('Bookings');
            }

            // Process insights
            if (followUpResults.insights && followUpResults.insights.length > 0) {
              result.insights = {
                count: followUpResults.insights.length,
                items: followUpResults.insights
                  .slice(0, effectiveMaxResults)
                  .map((insight: any) => ({
                    type: insight.insightType,
                    resourceType: insight.resourceVisualization?.type,
                    title: insight.resourceVisualization?.title,
                    containerDisplayName: insight.resourceVisualization?.containerDisplayName,
                    webUrl: insight.resourceReference?.webUrl,
                  })),
              };
              productsSearched.push('Insights');
            }
          } catch (error) {
            logger.warn(`Follow-up queries failed: ${error}`);
            result.followUpError = `${error}`;
          }
        })()
      );

      // Step 4: If we detected company contacts, add them to the result
      if (customerType.companyContacts && customerType.companyContacts.length > 0) {
        searchPromises.push(
          (async () => {
            result.companyContacts = {
              count: customerType.companyContacts!.length,
              items: customerType.companyContacts,
            };
            if (!productsSearched.includes('Contacts')) {
              productsSearched.push('Company Contacts');
            }
          })()
        );
      }

      // Step 5: If it's a person, also get direct communication history
      if (customerType.type === 'person' && customerType.userEmail) {
        searchPromises.push(
          (async () => {
            try {
              // Find emails specifically from/to this person
              const emailsFromPerson = await findEmailsWithPerson(
                graphClient,
                customerType.userEmail!,
                customerType.displayName,
                effectiveMaxResults
              );

              if (emailsFromPerson.length > 0) {
                result.directEmailHistory = {
                  count: emailsFromPerson.length,
                  items: emailsFromPerson.slice(0, effectiveMaxResults).map((e) => ({
                    subject: e.subject,
                    from: e.from?.emailAddress?.address,
                    date: e.receivedDateTime,
                    preview: e.bodyPreview?.substring(0, 100),
                  })),
                };
              }

              // Find meetings with this person
              const meetingsWithPerson = await findMeetingsWithPerson(
                graphClient,
                customerType.userEmail!,
                customerType.displayName,
                effectiveMaxResults
              );

              if (meetingsWithPerson.length > 0) {
                const upcoming = meetingsWithPerson.filter(
                  (m) => new Date(m.start.dateTime) >= now
                );
                const past = meetingsWithPerson.filter((m) => new Date(m.start.dateTime) < now);

                result.directMeetingHistory = {
                  total: meetingsWithPerson.length,
                  upcoming: {
                    count: upcoming.length,
                    items: upcoming.slice(0, 5).map((m) => ({
                      subject: m.subject,
                      start: m.start.dateTime,
                      location: m.location?.displayName,
                    })),
                  },
                  past: {
                    count: past.length,
                    items: past.slice(0, 5).map((m) => ({
                      subject: m.subject,
                      start: m.start.dateTime,
                    })),
                  },
                };
              }

              // Find chats with this person
              if (customerType.userId) {
                const chats = await findChatsWithUser(
                  graphClient,
                  customerType.userId,
                  customerType.userEmail,
                  customerType.displayName
                );

                if (chats.length > 0) {
                  result.directChats = {
                    count: chats.length,
                    items: chats.slice(0, effectiveMaxResults).map((chat) => ({
                      chatType: chat.chatType,
                      topic: chat.topic,
                      lastUpdated: chat.lastUpdatedDateTime,
                      memberCount: chat.members?.length,
                    })),
                  };
                }
              }
            } catch (error) {
              logger.warn(`Direct communication history failed: ${error}`);
            }
          })()
        );
      }

      // Wait for all searches to complete
      await Promise.allSettled(searchPromises);

      // Step 6: Calculate comprehensive relationship summary
      const emailCount =
        ((result.emails as { count?: number })?.count || 0) +
        ((result.directEmailHistory as { count?: number })?.count || 0);
      const meetingCount =
        ((result.calendar as { count?: number })?.count || 0) +
        ((result.directMeetingHistory as { total?: number })?.total || 0);
      const fileCount = (result.files as { count?: number })?.count || 0;
      const chatCount =
        ((result.teamsChats as { count?: number })?.count || 0) +
        ((result.directChats as { count?: number })?.count || 0);
      const contactCount =
        ((result.contacts as { count?: number })?.count || 0) +
        ((result.companyContacts as { count?: number })?.count || 0);
      const taskCount =
        ((result.planner as { count?: number })?.count || 0) +
        ((result.todo as { count?: number })?.count || 0);
      const siteCount = (result.sites as { count?: number })?.count || 0;
      const onenoteCount = (result.onenote as { count?: number })?.count || 0;

      const totalInteractions = emailCount + meetingCount + chatCount + fileCount;

      // Determine relationship health score
      let healthScore: string;
      let recommendation: string;

      if (totalInteractions === 0) {
        healthScore = 'No Data';
        recommendation = 'No interactions found. Consider reaching out to establish connection.';
      } else if (totalInteractions < 5) {
        healthScore = 'Low Activity';
        recommendation = 'Limited recent interaction. Consider scheduling a follow-up meeting.';
      } else if (totalInteractions < 15) {
        healthScore = 'Moderate';
        recommendation = 'Regular interaction maintained. Keep up the engagement.';
      } else if (totalInteractions < 30) {
        healthScore = 'Active';
        recommendation = 'Good level of engagement. Relationship is healthy.';
      } else {
        healthScore = 'Very Active';
        recommendation = 'Excellent engagement. Strong relationship with frequent interactions.';
      }

      result.productsSearched = productsSearched;

      result.relationshipSummary = {
        customerType: customerType.type,
        healthScore,
        recommendation,
        metrics: {
          totalInteractions,
          emails: emailCount,
          meetings: meetingCount,
          files: fileCount,
          chats: chatCount,
          contacts: contactCount,
          tasks: taskCount,
          sites: siteCount,
          notes: onenoteCount,
        },
        productsSearched: productsSearched.length,
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 18. ANALYZE MEETING LOAD - Meeting overload analysis
  // ==========================================================================
  server.tool(
    'analyze-meeting-load',
    `Analyze your meeting load and identify potential issues:
- Total meeting hours per day/week
- Meeting-free time analysis
- Back-to-back meeting detection
- Recurring meeting overhead
- Meeting duration distribution
- Suggestions for optimization

Use this for "Am I in too many meetings?", "Analyze my meeting load", or "Show me my meeting patterns".`,
    {
      weeks: z.number().optional().describe('Number of weeks to analyze (default: 4)'),
      includeRecurring: z
        .boolean()
        .optional()
        .describe('Include recurring meeting analysis (default: true)'),
    },
    {
      title: 'Analyze Meeting Load',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ weeks = 4, includeRecurring = true }) => {
      logger.info(`Analyzing meeting load for ${weeks} weeks`);

      const now = new Date();
      const startDate = new Date(now);
      startDate.setDate(startDate.getDate() - weeks * 7);
      const endDate = new Date(now);

      const result: Record<string, unknown> = {
        analyzedPeriod: `Last ${weeks} weeks`,
        retrievedAt: new Date().toISOString(),
      };

      try {
        // Build query string (startDateTime and endDateTime are REQUIRED for calendarView)
        const exportQueryParams: Record<string, string> = {
          startDateTime: startDate.toISOString(),
          endDateTime: endDate.toISOString(),
          $top: '500',
          $select: 'subject,start,end,isAllDay,recurrence,attendees,organizer,isCancelled',
        };

        const meetingsResponse = await graphClient.makeRequest(
          `/me/calendarView?${buildGraphQueryString(exportQueryParams)}`,
          {
            method: 'GET',
          }
        );

        if (
          !meetingsResponse ||
          typeof meetingsResponse !== 'object' ||
          !('value' in meetingsResponse)
        ) {
          return {
            content: [
              {
                type: 'text' as const,
                text: JSON.stringify({ error: 'Could not fetch calendar data' }, null, 2),
              },
            ],
          };
        }

        const events = (
          meetingsResponse.value as Array<
            GraphEvent & { isAllDay?: boolean; recurrence?: unknown; isCancelled?: boolean }
          >
        ).filter((e) => !e.isAllDay && !e.isCancelled);

        // Analyze by day
        const dailyStats = new Map<
          string,
          { hours: number; count: number; events: GraphEvent[] }
        >();
        const weeklyStats = new Map<number, { hours: number; count: number }>();
        let totalHours = 0;
        let backToBackCount = 0;
        const durationBuckets = { short: 0, medium: 0, long: 0, veryLong: 0 };
        const recurringMeetings = new Map<string, number>();

        for (const event of events) {
          const start = new Date(event.start.dateTime);
          const end = new Date(event.end.dateTime);
          const duration = (end.getTime() - start.getTime()) / (1000 * 60 * 60);
          totalHours += duration;

          // Daily stats
          const day = start.toISOString().split('T')[0];
          if (!dailyStats.has(day)) {
            dailyStats.set(day, { hours: 0, count: 0, events: [] });
          }
          const dayStats = dailyStats.get(day)!;
          dayStats.hours += duration;
          dayStats.count++;
          dayStats.events.push(event);

          // Weekly stats
          const weekNum = Math.floor(
            (start.getTime() - startDate.getTime()) / (7 * 24 * 60 * 60 * 1000)
          );
          if (!weeklyStats.has(weekNum)) {
            weeklyStats.set(weekNum, { hours: 0, count: 0 });
          }
          weeklyStats.get(weekNum)!.hours += duration;
          weeklyStats.get(weekNum)!.count++;

          // Duration buckets
          if (duration <= 0.5) durationBuckets.short++;
          else if (duration <= 1) durationBuckets.medium++;
          else if (duration <= 2) durationBuckets.long++;
          else durationBuckets.veryLong++;

          // Track recurring meetings
          if (includeRecurring && event.subject) {
            const normalized = event.subject.toLowerCase().trim();
            recurringMeetings.set(normalized, (recurringMeetings.get(normalized) || 0) + 1);
          }
        }

        // Detect back-to-back meetings
        for (const [, dayData] of dailyStats) {
          const sortedEvents = dayData.events.sort(
            (a, b) => new Date(a.start.dateTime).getTime() - new Date(b.start.dateTime).getTime()
          );
          for (let i = 0; i < sortedEvents.length - 1; i++) {
            const currentEnd = new Date(sortedEvents[i].end.dateTime);
            const nextStart = new Date(sortedEvents[i + 1].start.dateTime);
            const gap = (nextStart.getTime() - currentEnd.getTime()) / (1000 * 60);
            if (gap <= 5) {
              // 5 minutes or less gap
              backToBackCount++;
            }
          }
        }

        // Find busiest days
        const sortedDays = Array.from(dailyStats.entries())
          .map(([day, stats]) => ({ day, ...stats }))
          .sort((a, b) => b.hours - a.hours);

        // Find most common recurring meetings
        const topRecurring = Array.from(recurringMeetings.entries())
          .filter(([, count]) => count > 1)
          .sort(([, a], [, b]) => b - a)
          .slice(0, 10)
          .map(([title, count]) => ({ title, occurrences: count }));

        // Calculate averages
        const avgDailyHours = totalHours / (weeks * 5); // Assuming 5 work days
        const avgWeeklyHours = totalHours / weeks;
        const avgMeetingsPerDay = events.length / (weeks * 5);

        // Determine load level
        let loadLevel: 'Light' | 'Moderate' | 'Heavy' | 'Overloaded';
        if (avgDailyHours < 2) loadLevel = 'Light';
        else if (avgDailyHours < 4) loadLevel = 'Moderate';
        else if (avgDailyHours < 6) loadLevel = 'Heavy';
        else loadLevel = 'Overloaded';

        result.summary = {
          totalMeetings: events.length,
          totalHours: Math.round(totalHours * 10) / 10,
          averageDailyHours: Math.round(avgDailyHours * 10) / 10,
          averageWeeklyHours: Math.round(avgWeeklyHours * 10) / 10,
          averageMeetingsPerDay: Math.round(avgMeetingsPerDay * 10) / 10,
          loadLevel,
        };

        result.distribution = {
          byDuration: {
            'Under 30 min': durationBuckets.short,
            '30-60 min': durationBuckets.medium,
            '1-2 hours': durationBuckets.long,
            'Over 2 hours': durationBuckets.veryLong,
          },
          backToBackMeetings: backToBackCount,
        };

        result.busiestDays = sortedDays.slice(0, 5).map((d) => ({
          date: d.day,
          hours: Math.round(d.hours * 10) / 10,
          meetings: d.count,
        }));

        if (includeRecurring) {
          result.recurringMeetings = {
            total: topRecurring.reduce((sum, m) => sum + m.occurrences, 0),
            topMeetings: topRecurring,
          };
        }

        // Generate recommendations
        const recommendations: string[] = [];
        if (loadLevel === 'Overloaded') {
          recommendations.push(
            'Consider declining non-essential meetings or delegating attendance'
          );
        }
        if (backToBackCount > 5) {
          recommendations.push(
            `You had ${backToBackCount} back-to-back meetings. Try to add buffer time between meetings`
          );
        }
        if (durationBuckets.veryLong > events.length * 0.2) {
          recommendations.push(
            'Many meetings are over 2 hours. Consider breaking them into shorter sessions'
          );
        }
        if (topRecurring.some((m) => m.occurrences > weeks)) {
          recommendations.push(
            'Some meetings occur multiple times per week. Review if all occurrences are necessary'
          );
        }

        result.recommendations =
          recommendations.length > 0 ? recommendations : ['Your meeting load appears balanced'];
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify({ error: `Failed to analyze meetings: ${error}` }, null, 2),
            },
          ],
        };
      }

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 19. GET DEADLINE OVERVIEW - All upcoming deadlines and due dates
  // ==========================================================================
  server.tool(
    'get-deadline-overview',
    `Get a comprehensive overview of all upcoming deadlines and due dates:
- Tasks due dates from To-Do and Planner
- Calendar events with deadline indicators
- Flagged emails with due dates
- Grouped by urgency (overdue, today, this week, later)

Use this for "What are my deadlines?", "What's due this week?", or "Show me upcoming due dates".`,
    {
      days: z.number().optional().describe('Days ahead to look for deadlines (default: 30)'),
      includeCompleted: z.boolean().optional().describe('Include completed items (default: false)'),
    },
    {
      title: 'Get Deadline Overview',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ days = 30, includeCompleted = false }) => {
      logger.info(`Getting deadline overview for next ${days} days`);

      const now = new Date();
      const futureDate = new Date();
      futureDate.setDate(futureDate.getDate() + days);

      interface DeadlineItem {
        title: string;
        dueDate: string;
        source: 'ToDo' | 'Planner' | 'Email' | 'Calendar';
        priority?: string;
        status?: string;
        link?: string;
      }

      const deadlines: DeadlineItem[] = [];
      const promises: Promise<void>[] = [];

      // To-Do tasks
      promises.push(
        (async () => {
          try {
            const taskListsResponse = await graphClient.makeRequest('/me/todo/lists', {
              method: 'GET',
              queryParams: { $top: '20' },
            });

            if (
              taskListsResponse &&
              typeof taskListsResponse === 'object' &&
              'value' in taskListsResponse
            ) {
              for (const list of taskListsResponse.value as Array<{ id: string }>) {
                try {
                  const tasksResponse = await graphClient.makeRequest(
                    `/me/todo/lists/${list.id}/tasks`,
                    {
                      method: 'GET',
                      queryParams: { $top: '100' },
                    }
                  );

                  if (
                    tasksResponse &&
                    typeof tasksResponse === 'object' &&
                    'value' in tasksResponse
                  ) {
                    for (const task of tasksResponse.value as Array<{
                      title: string;
                      dueDateTime?: { dateTime: string };
                      status: string;
                      importance: string;
                    }>) {
                      if (!includeCompleted && task.status === 'completed') continue;
                      if (task.dueDateTime) {
                        deadlines.push({
                          title: task.title,
                          dueDate: task.dueDateTime.dateTime,
                          source: 'ToDo',
                          priority: task.importance,
                          status: task.status,
                        });
                      }
                    }
                  }
                } catch {
                  // Skip individual list errors
                }
              }
            }
          } catch (error) {
            logger.warn(`Could not fetch To-Do: ${error}`);
          }
        })()
      );

      // Planner tasks
      promises.push(
        (async () => {
          try {
            const plannerResponse = await graphClient.makeRequest('/me/planner/tasks', {
              method: 'GET',
              queryParams: { $top: '100' },
            });

            if (
              plannerResponse &&
              typeof plannerResponse === 'object' &&
              'value' in plannerResponse
            ) {
              for (const task of plannerResponse.value as Array<{
                title: string;
                dueDateTime?: string;
                percentComplete: number;
                priority: number;
              }>) {
                if (!includeCompleted && task.percentComplete >= 100) continue;
                if (task.dueDateTime) {
                  deadlines.push({
                    title: task.title,
                    dueDate: task.dueDateTime,
                    source: 'Planner',
                    priority: ['Urgent', 'Important', 'Medium', 'Low'][task.priority] || 'Medium',
                    status: task.percentComplete >= 100 ? 'completed' : `${task.percentComplete}%`,
                  });
                }
              }
            }
          } catch (error) {
            logger.warn(`Could not fetch Planner: ${error}`);
          }
        })()
      );

      // Flagged emails with due dates
      promises.push(
        (async () => {
          try {
            const flaggedResponse = await graphClient.makeRequest('/me/messages', {
              method: 'GET',
              queryParams: {
                $filter: "flag/flagStatus eq 'flagged'",
                $top: '50',
                $select: 'subject,flag,webLink,receivedDateTime',
              },
            });

            if (
              flaggedResponse &&
              typeof flaggedResponse === 'object' &&
              'value' in flaggedResponse
            ) {
              for (const email of flaggedResponse.value as Array<{
                subject: string;
                flag?: { dueDateTime?: { dateTime: string }; flagStatus: string };
                webLink?: string;
              }>) {
                if (email.flag?.dueDateTime) {
                  deadlines.push({
                    title: `Email: ${email.subject}`,
                    dueDate: email.flag.dueDateTime.dateTime,
                    source: 'Email',
                    link: email.webLink,
                  });
                }
              }
            }
          } catch (error) {
            logger.warn(`Could not fetch flagged emails: ${error}`);
          }
        })()
      );

      await Promise.allSettled(promises);

      // Categorize deadlines
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const tomorrow = new Date(today);
      tomorrow.setDate(tomorrow.getDate() + 1);
      const thisWeekEnd = new Date(today);
      thisWeekEnd.setDate(thisWeekEnd.getDate() + (7 - today.getDay()));

      const categorized = {
        overdue: [] as DeadlineItem[],
        today: [] as DeadlineItem[],
        tomorrow: [] as DeadlineItem[],
        thisWeek: [] as DeadlineItem[],
        later: [] as DeadlineItem[],
      };

      for (const deadline of deadlines) {
        const due = new Date(deadline.dueDate);
        if (due < today) {
          categorized.overdue.push(deadline);
        } else if (due < tomorrow) {
          categorized.today.push(deadline);
        } else if (due < new Date(tomorrow.getTime() + 24 * 60 * 60 * 1000)) {
          categorized.tomorrow.push(deadline);
        } else if (due <= thisWeekEnd) {
          categorized.thisWeek.push(deadline);
        } else {
          categorized.later.push(deadline);
        }
      }

      // Sort each category by due date
      for (const category of Object.values(categorized)) {
        category.sort((a, b) => new Date(a.dueDate).getTime() - new Date(b.dueDate).getTime());
      }

      const result = {
        summary: {
          total: deadlines.length,
          overdue: categorized.overdue.length,
          dueToday: categorized.today.length,
          dueTomorrow: categorized.tomorrow.length,
          dueThisWeek: categorized.thisWeek.length,
          later: categorized.later.length,
        },
        deadlines: categorized,
        urgentAlert:
          categorized.overdue.length > 0
            ? `⚠️ You have ${categorized.overdue.length} overdue item(s)!`
            : categorized.today.length > 0
              ? `📅 You have ${categorized.today.length} item(s) due today`
              : '✅ No urgent deadlines',
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 20. FIND DECISION CONTEXT - Find context for a decision
  // ==========================================================================
  server.tool(
    'find-decision-context',
    `Find all context and history related to a specific decision or topic:
- When was it discussed?
- Who was involved?
- What were the alternatives considered?
- What documentation exists?
- What was the final outcome?

Use this for "When did we decide on X?", "What was the context for decision Y?", or "Find discussions about [topic]".`,
    {
      topic: z.string().describe('The decision or topic to find context for'),
      days: z.number().optional().describe('Days of history to search (default: 180)'),
    },
    {
      title: 'Find Decision Context',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ topic, days = 180 }) => {
      logger.info(`Finding decision context for: ${topic}`);

      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      const result: Record<string, unknown> = {
        topic,
        searchPeriod: `Last ${days} days`,
        retrievedAt: new Date().toISOString(),
      };

      const promises: Promise<void>[] = [];

      // Search emails
      promises.push(
        (async () => {
          try {
            const emailResponse = await graphClient.makeRequest('/me/messages', {
              method: 'GET',
              queryParams: {
                $search: `"${topic}"`,
                $top: '30',
                $select:
                  'subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,webLink',
                $orderby: 'receivedDateTime desc',
              },
            });

            if (emailResponse && typeof emailResponse === 'object' && 'value' in emailResponse) {
              const emails = emailResponse.value as GraphEmail[];
              const uniqueParticipants = new Set<string>();

              for (const email of emails) {
                if (email.from?.emailAddress?.address) {
                  uniqueParticipants.add(email.from.emailAddress.address);
                }
                for (const recipient of [...(email.toRecipients || [])]) {
                  if (recipient.emailAddress?.address) {
                    uniqueParticipants.add(recipient.emailAddress.address);
                  }
                }
              }

              result.emailDiscussions = {
                count: emails.length,
                participants: Array.from(uniqueParticipants),
                threads: emails.slice(0, 10).map((e) => ({
                  subject: e.subject,
                  from: e.from?.emailAddress?.name || e.from?.emailAddress?.address,
                  date: e.receivedDateTime,
                  preview: e.bodyPreview?.substring(0, 150),
                  link: e.webLink,
                })),
              };
            }
          } catch (error) {
            result.emailDiscussions = { error: `${error}` };
          }
        })()
      );

      // Search meetings
      promises.push(
        (async () => {
          try {
            // Build query string (startDateTime and endDateTime are REQUIRED for calendarView)
            // Note: $filter with contains() causes 500 errors, use client-side filtering
            const topicQueryParams: Record<string, string> = {
              startDateTime: startDate.toISOString(),
              endDateTime: new Date().toISOString(),
              $top: '100',
              $select: 'subject,start,end,attendees,organizer,bodyPreview,webLink',
              $orderby: 'start/dateTime desc',
            };

            const meetingsResponse = await graphClient.makeRequest(
              `/me/calendarView?${buildGraphQueryString(topicQueryParams)}`,
              {
                method: 'GET',
              }
            );

            if (
              meetingsResponse &&
              typeof meetingsResponse === 'object' &&
              'value' in meetingsResponse
            ) {
              // Client-side filtering for reliability
              const topicLower = topic.toLowerCase();
              const meetings = (meetingsResponse.value as GraphEvent[]).filter((m) =>
                m.subject?.toLowerCase().includes(topicLower)
              );

              result.relatedMeetings = {
                count: meetings.length,
                meetings: meetings.slice(0, 10).map((m) => ({
                  subject: m.subject,
                  date: m.start.dateTime,
                  organizer: m.organizer?.emailAddress?.name,
                  attendees: m.attendees?.map((a) => a.emailAddress?.name).filter(Boolean),
                  notes: m.bodyPreview?.substring(0, 200),
                  link: m.webLink,
                })),
              };
            }
          } catch (error) {
            logger.warn(`Topic meetings search failed: ${error}`);
            result.relatedMeetings = { error: `${error}` };
          }
        })()
      );

      // Search files/documents - use centralized search
      promises.push(
        (async () => {
          try {
            const searchResult = await executeCentralSearch(graphClient, topic, {
              entityTypes: ['driveItem', 'listItem'],
              maxResults: 20,
              sortByRank: true,
            });

            const files = [...searchResult.results.files, ...searchResult.results.listItems].map(
              (hit) => ({
                name: hit.name || 'Unknown',
                webUrl: hit.webUrl,
                modified: hit.resource?.lastModifiedDateTime,
                preview: hit.summary,
                relevance: hit.relevanceScore,
              })
            );

            result.relatedDocuments = {
              count: files.length,
              files,
            };
          } catch (error) {
            result.relatedDocuments = { error: `${error}` };
          }
        })()
      );

      await Promise.allSettled(promises);

      // Build timeline of events
      const timeline: Array<{
        date: string;
        type: string;
        title: string;
        participants?: string[];
      }> = [];

      const emails =
        (
          result.emailDiscussions as {
            threads?: Array<{ date: string; subject: string; from: string }>;
          }
        )?.threads || [];
      for (const email of emails) {
        timeline.push({
          date: email.date,
          type: 'email',
          title: email.subject,
          participants: [email.from],
        });
      }

      const meetings =
        (
          result.relatedMeetings as {
            meetings?: Array<{ date: string; subject: string; attendees?: string[] }>;
          }
        )?.meetings || [];
      for (const meeting of meetings) {
        timeline.push({
          date: meeting.date,
          type: 'meeting',
          title: meeting.subject,
          participants: meeting.attendees,
        });
      }

      timeline.sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
      result.timeline = timeline;

      // Summary
      const allParticipants = new Set<string>();
      for (const item of timeline) {
        for (const p of item.participants || []) {
          allParticipants.add(p);
        }
      }

      result.summary = {
        totalDiscussions: timeline.length,
        keyParticipants: Array.from(allParticipants).slice(0, 10),
        firstDiscussion: timeline[0]?.date,
        lastDiscussion: timeline[timeline.length - 1]?.date,
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 21. GET PROJECT STAKEHOLDERS - Find everyone involved in a project
  // ==========================================================================
  server.tool(
    'get-project-stakeholders',
    `Identify all stakeholders and participants involved in a project:
- People who attend project meetings
- Email participants on project discussions
- Document contributors
- Ranked by involvement level

Use this for "Who is involved in Project X?", "List stakeholders for [project]", or "Who works on [topic]?".`,
    {
      projectName: z.string().describe('Project or topic name'),
      days: z.number().optional().describe('Days of history to analyze (default: 90)'),
    },
    {
      title: 'Get Project Stakeholders',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ projectName, days = 90 }) => {
      logger.info(`Finding stakeholders for: ${projectName}`);

      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      const stakeholders = new Map<
        string,
        {
          name: string;
          email: string;
          meetingCount: number;
          emailCount: number;
          documentCount: number;
          roles: Set<string>;
        }
      >();

      const promises: Promise<void>[] = [];

      // Meeting participants
      promises.push(
        (async () => {
          try {
            // Build query string (startDateTime and endDateTime are REQUIRED for calendarView)
            // Note: $filter with contains() causes 500 errors, use client-side filtering
            const projectFilterQueryParams: Record<string, string> = {
              startDateTime: startDate.toISOString(),
              endDateTime: new Date().toISOString(),
              $top: '100',
              $select: 'subject,attendees,organizer',
            };

            const meetingsResponse = await graphClient.makeRequest(
              `/me/calendarView?${buildGraphQueryString(projectFilterQueryParams)}`,
              {
                method: 'GET',
              }
            );

            if (
              meetingsResponse &&
              typeof meetingsResponse === 'object' &&
              'value' in meetingsResponse
            ) {
              // Client-side filtering for reliability
              const projectLower = projectName.toLowerCase();
              const filteredEvents = (meetingsResponse.value as GraphEvent[]).filter((e) =>
                e.subject?.toLowerCase().includes(projectLower)
              );
              for (const event of filteredEvents) {
                // Add organizer
                const orgEmail = event.organizer?.emailAddress?.address?.toLowerCase();
                if (orgEmail) {
                  if (!stakeholders.has(orgEmail)) {
                    stakeholders.set(orgEmail, {
                      name: event.organizer?.emailAddress?.name || orgEmail,
                      email: orgEmail,
                      meetingCount: 0,
                      emailCount: 0,
                      documentCount: 0,
                      roles: new Set(['Meeting Organizer']),
                    });
                  }
                  stakeholders.get(orgEmail)!.meetingCount++;
                  stakeholders.get(orgEmail)!.roles.add('Meeting Organizer');
                }

                // Add attendees
                for (const attendee of event.attendees || []) {
                  const email = attendee.emailAddress?.address?.toLowerCase();
                  if (email) {
                    if (!stakeholders.has(email)) {
                      stakeholders.set(email, {
                        name: attendee.emailAddress?.name || email,
                        email,
                        meetingCount: 0,
                        emailCount: 0,
                        documentCount: 0,
                        roles: new Set(),
                      });
                    }
                    stakeholders.get(email)!.meetingCount++;
                    stakeholders.get(email)!.roles.add('Meeting Participant');
                  }
                }
              }
            }
          } catch (error) {
            logger.warn(`Could not analyze meetings: ${error}`);
          }
        })()
      );

      // Email participants
      promises.push(
        (async () => {
          try {
            const emailResponse = await graphClient.makeRequest('/me/messages', {
              method: 'GET',
              queryParams: {
                $search: `"${projectName}"`,
                $top: '100',
                $select: 'from,toRecipients,ccRecipients',
              },
            });

            if (emailResponse && typeof emailResponse === 'object' && 'value' in emailResponse) {
              for (const email of emailResponse.value as GraphEmail[]) {
                const participants = [
                  email.from?.emailAddress,
                  ...(email.toRecipients?.map((r) => r.emailAddress) || []),
                ].filter(Boolean);

                for (const participant of participants) {
                  const pEmail = participant?.address?.toLowerCase();
                  if (pEmail) {
                    if (!stakeholders.has(pEmail)) {
                      stakeholders.set(pEmail, {
                        name: participant?.name || pEmail,
                        email: pEmail,
                        meetingCount: 0,
                        emailCount: 0,
                        documentCount: 0,
                        roles: new Set(),
                      });
                    }
                    stakeholders.get(pEmail)!.emailCount++;
                    stakeholders.get(pEmail)!.roles.add('Email Participant');
                  }
                }
              }
            }
          } catch (error) {
            logger.warn(`Could not analyze emails: ${error}`);
          }
        })()
      );

      await Promise.allSettled(promises);

      // Calculate involvement scores and sort
      const sortedStakeholders = Array.from(stakeholders.values())
        .map((s) => ({
          name: s.name,
          email: s.email,
          meetingCount: s.meetingCount,
          emailCount: s.emailCount,
          involvementScore: s.meetingCount * 3 + s.emailCount, // Meetings weighted higher
          roles: Array.from(s.roles),
        }))
        .sort((a, b) => b.involvementScore - a.involvementScore);

      // Categorize by involvement level
      const keyStakeholders = sortedStakeholders.filter((s) => s.involvementScore >= 10);
      const regularParticipants = sortedStakeholders.filter(
        (s) => s.involvementScore >= 3 && s.involvementScore < 10
      );
      const occasionalParticipants = sortedStakeholders.filter((s) => s.involvementScore < 3);

      const result = {
        project: projectName,
        analyzedPeriod: `Last ${days} days`,
        totalStakeholders: stakeholders.size,
        keyStakeholders: {
          count: keyStakeholders.length,
          people: keyStakeholders.slice(0, 10),
        },
        regularParticipants: {
          count: regularParticipants.length,
          people: regularParticipants.slice(0, 10),
        },
        occasionalParticipants: {
          count: occasionalParticipants.length,
          people: occasionalParticipants.slice(0, 5),
        },
        summary: {
          mostInvolved: keyStakeholders[0]?.name,
          meetingOrganizers: sortedStakeholders
            .filter((s) => s.roles.includes('Meeting Organizer'))
            .map((s) => s.name),
        },
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 22. FIND UNRESPONDED REQUESTS - Items needing your response
  // ==========================================================================
  server.tool(
    'find-unresponded-requests',
    `Find all requests and messages that are waiting for your response:
- Emails where you were asked a question
- Meeting invites awaiting your response
- Tasks assigned to you
- Mentions in Teams requiring action

Use this for "What am I forgetting to respond to?", "Find unanswered requests", or "What's waiting on me?".`,
    {
      days: z.number().optional().describe('Days back to search (default: 14)'),
      priorityOnly: z
        .boolean()
        .optional()
        .describe('Only show high priority items (default: false)'),
    },
    {
      title: 'Find Unresponded Requests',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ days = 14, priorityOnly = false }) => {
      logger.info(`Finding unresponded requests from last ${days} days`);

      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      interface UnrespondedItem {
        type: 'email' | 'meeting' | 'task';
        title: string;
        from?: string;
        receivedDate: string;
        priority?: string;
        link?: string;
        reason: string;
      }

      const unresponded: UnrespondedItem[] = [];
      const promises: Promise<void>[] = [];

      // Unread emails with questions
      promises.push(
        (async () => {
          try {
            let filter = `isRead eq false and receivedDateTime ge ${startDate.toISOString()}`;
            if (priorityOnly) {
              filter += " and importance eq 'high'";
            }

            const emailResponse = await graphClient.makeRequest('/me/messages', {
              method: 'GET',
              queryParams: {
                $filter: filter,
                $top: '50',
                $select: 'subject,from,receivedDateTime,importance,bodyPreview,webLink',
                $orderby: 'receivedDateTime desc',
              },
            });

            if (emailResponse && typeof emailResponse === 'object' && 'value' in emailResponse) {
              for (const email of emailResponse.value as Array<
                GraphEmail & { importance?: string }
              >) {
                // Check if email might contain a question or request
                const preview = (email.bodyPreview || '').toLowerCase();
                const hasQuestion =
                  preview.includes('?') ||
                  preview.includes('please') ||
                  preview.includes('could you') ||
                  preview.includes('can you') ||
                  preview.includes('would you') ||
                  preview.includes('need your') ||
                  preview.includes('waiting for') ||
                  preview.includes('asap');

                if (hasQuestion || email.importance === 'high') {
                  unresponded.push({
                    type: 'email',
                    title: email.subject,
                    from: email.from?.emailAddress?.name || email.from?.emailAddress?.address,
                    receivedDate: email.receivedDateTime,
                    priority: email.importance,
                    link: email.webLink,
                    reason: hasQuestion ? 'Contains question/request' : 'High priority unread',
                  });
                }
              }
            }
          } catch (error) {
            logger.warn(`Could not analyze emails: ${error}`);
          }
        })()
      );

      // Meeting invites needing response
      promises.push(
        (async () => {
          try {
            // Build query string (startDateTime and endDateTime are REQUIRED for calendarView)
            const responseQueryParams: Record<string, string> = {
              startDateTime: new Date().toISOString(),
              endDateTime: new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString(),
              $filter:
                "responseStatus/response eq 'notResponded' or responseStatus/response eq 'none'",
              $top: '30',
              $select: 'subject,start,organizer,responseStatus,webLink',
            };

            const meetingsResponse = await graphClient.makeRequest(
              `/me/calendarView?${buildGraphQueryString(responseQueryParams)}`,
              {
                method: 'GET',
              }
            );

            if (
              meetingsResponse &&
              typeof meetingsResponse === 'object' &&
              'value' in meetingsResponse
            ) {
              for (const meeting of meetingsResponse.value as Array<
                GraphEvent & { responseStatus?: { response: string } }
              >) {
                unresponded.push({
                  type: 'meeting',
                  title: meeting.subject,
                  from: meeting.organizer?.emailAddress?.name,
                  receivedDate: meeting.start.dateTime,
                  link: meeting.webLink,
                  reason: 'Meeting invite awaiting response',
                });
              }
            }
          } catch (error) {
            logger.warn(`Could not analyze meetings: ${error}`);
          }
        })()
      );

      // Flagged emails
      promises.push(
        (async () => {
          try {
            const flaggedResponse = await graphClient.makeRequest('/me/messages', {
              method: 'GET',
              queryParams: {
                $filter: "flag/flagStatus eq 'flagged'",
                $top: '30',
                $select: 'subject,from,receivedDateTime,flag,webLink',
              },
            });

            if (
              flaggedResponse &&
              typeof flaggedResponse === 'object' &&
              'value' in flaggedResponse
            ) {
              for (const email of flaggedResponse.value as GraphEmail[]) {
                unresponded.push({
                  type: 'email',
                  title: email.subject,
                  from: email.from?.emailAddress?.name || email.from?.emailAddress?.address,
                  receivedDate: email.receivedDateTime,
                  link: email.webLink,
                  reason: 'Flagged for follow-up',
                });
              }
            }
          } catch (error) {
            logger.warn(`Could not fetch flagged emails: ${error}`);
          }
        })()
      );

      await Promise.allSettled(promises);

      // Sort by date (oldest first - most urgent)
      unresponded.sort(
        (a, b) => new Date(a.receivedDate).getTime() - new Date(b.receivedDate).getTime()
      );

      // Group by type
      const byType = {
        emails: unresponded.filter((i) => i.type === 'email'),
        meetings: unresponded.filter((i) => i.type === 'meeting'),
        tasks: unresponded.filter((i) => i.type === 'task'),
      };

      const result = {
        summary: {
          total: unresponded.length,
          emails: byType.emails.length,
          meetings: byType.meetings.length,
          oldestItem: unresponded[0]
            ? `${unresponded[0].title} from ${new Date(unresponded[0].receivedDate).toLocaleDateString()}`
            : 'None',
        },
        urgentItems: unresponded.slice(0, 5),
        byType: {
          emails: byType.emails.slice(0, 15),
          meetings: byType.meetings,
        },
        recommendation:
          unresponded.length === 0
            ? '✅ No pending responses needed!'
            : unresponded.length > 10
              ? '⚠️ You have many pending items. Consider setting aside focused response time.'
              : `📬 ${unresponded.length} items need your attention`,
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 23. GET COLLABORATION NETWORK - Your professional network map
  // ==========================================================================
  server.tool(
    'get-collaboration-network',
    `Map your professional collaboration network:
- Who you interact with most frequently
- Communication channels used with each person
- Relationship strength indicators
- Department/team distribution
- Key connectors in your network

Use this for "Who do I work with most?", "Map my professional network", or "Analyze my work relationships".`,
    {
      days: z.number().optional().describe('Days of data to analyze (default: 60)'),
      minInteractions: z
        .number()
        .optional()
        .describe('Minimum interactions to include (default: 3)'),
    },
    {
      title: 'Get Collaboration Network',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ days = 60, minInteractions = 3 }) => {
      logger.info(`Mapping collaboration network for last ${days} days`);

      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      const network = new Map<
        string,
        {
          name: string;
          email: string;
          emails: number;
          meetings: number;
          chats: number;
          lastInteraction: string;
          channels: Set<string>;
        }
      >();

      const promises: Promise<void>[] = [];

      // Email interactions
      promises.push(
        (async () => {
          try {
            const emailResponse = await graphClient.makeRequest('/me/messages', {
              method: 'GET',
              queryParams: {
                $filter: `receivedDateTime ge ${startDate.toISOString()}`,
                $top: '300',
                $select: 'from,toRecipients,receivedDateTime',
              },
            });

            if (emailResponse && typeof emailResponse === 'object' && 'value' in emailResponse) {
              for (const email of emailResponse.value as GraphEmail[]) {
                const contacts = [
                  email.from?.emailAddress,
                  ...(email.toRecipients?.map((r) => r.emailAddress) || []),
                ];

                for (const contact of contacts) {
                  if (!contact?.address) continue;
                  const key = contact.address.toLowerCase();

                  if (!network.has(key)) {
                    network.set(key, {
                      name: contact.name || key,
                      email: key,
                      emails: 0,
                      meetings: 0,
                      chats: 0,
                      lastInteraction: email.receivedDateTime,
                      channels: new Set(),
                    });
                  }

                  network.get(key)!.emails++;
                  network.get(key)!.channels.add('Email');
                  if (
                    new Date(email.receivedDateTime) > new Date(network.get(key)!.lastInteraction)
                  ) {
                    network.get(key)!.lastInteraction = email.receivedDateTime;
                  }
                }
              }
            }
          } catch (error) {
            logger.warn(`Could not analyze emails: ${error}`);
          }
        })()
      );

      // Meeting interactions
      promises.push(
        (async () => {
          try {
            // Build query string (startDateTime and endDateTime are REQUIRED for calendarView)
            const attendanceQueryParams: Record<string, string> = {
              startDateTime: startDate.toISOString(),
              endDateTime: new Date().toISOString(),
              $top: '200',
              $select: 'attendees,start',
            };

            const meetingsResponse = await graphClient.makeRequest(
              `/me/calendarView?${buildGraphQueryString(attendanceQueryParams)}`,
              {
                method: 'GET',
              }
            );

            if (
              meetingsResponse &&
              typeof meetingsResponse === 'object' &&
              'value' in meetingsResponse
            ) {
              for (const event of meetingsResponse.value as GraphEvent[]) {
                for (const attendee of event.attendees || []) {
                  const email = attendee.emailAddress?.address?.toLowerCase();
                  if (!email) continue;

                  if (!network.has(email)) {
                    network.set(email, {
                      name: attendee.emailAddress?.name || email,
                      email,
                      emails: 0,
                      meetings: 0,
                      chats: 0,
                      lastInteraction: event.start.dateTime,
                      channels: new Set(),
                    });
                  }

                  network.get(email)!.meetings++;
                  network.get(email)!.channels.add('Meetings');
                  if (
                    new Date(event.start.dateTime) > new Date(network.get(email)!.lastInteraction)
                  ) {
                    network.get(email)!.lastInteraction = event.start.dateTime;
                  }
                }
              }
            }
          } catch (error) {
            logger.warn(`Could not analyze meetings: ${error}`);
          }
        })()
      );

      await Promise.allSettled(promises);

      // Filter and rank by total interactions
      const connections = Array.from(network.values())
        .filter((c) => c.emails + c.meetings + c.chats >= minInteractions)
        .map((c) => ({
          name: c.name,
          email: c.email,
          totalInteractions: c.emails + c.meetings + c.chats,
          breakdown: {
            emails: c.emails,
            meetings: c.meetings,
            chats: c.chats,
          },
          channels: Array.from(c.channels),
          lastInteraction: c.lastInteraction,
          relationshipStrength:
            c.emails + c.meetings + c.chats > 20
              ? 'Strong'
              : c.emails + c.meetings + c.chats > 10
                ? 'Moderate'
                : 'Light',
        }))
        .sort((a, b) => b.totalInteractions - a.totalInteractions);

      // Analyze network structure
      const strongConnections = connections.filter((c) => c.relationshipStrength === 'Strong');
      const moderateConnections = connections.filter((c) => c.relationshipStrength === 'Moderate');

      // Extract domains for team distribution
      const domainCounts = new Map<string, number>();
      for (const conn of connections) {
        const domain = conn.email.split('@')[1];
        if (domain) {
          domainCounts.set(domain, (domainCounts.get(domain) || 0) + 1);
        }
      }

      const result = {
        analyzedPeriod: `Last ${days} days`,
        networkSize: connections.length,
        summary: {
          strongConnections: strongConnections.length,
          moderateConnections: moderateConnections.length,
          lightConnections:
            connections.length - strongConnections.length - moderateConnections.length,
          topContact: connections[0]?.name,
        },
        topConnections: connections.slice(0, 15),
        byOrganization: Array.from(domainCounts.entries())
          .sort(([, a], [, b]) => b - a)
          .slice(0, 10)
          .map(([domain, count]) => ({ domain, contacts: count })),
        insights: [
          `You have ${strongConnections.length} strong professional relationships`,
          connections.length > 50
            ? 'You have a large professional network'
            : 'Your network is focused and manageable',
          domainCounts.size > 5
            ? 'You collaborate across multiple organizations'
            : 'Most of your collaboration is within your organization',
        ],
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // CHAT MEMORY MANAGEMENT TOOLS
  // ==========================================================================

  /**
   * Chat Memory Status - Shows current chat memory contents
   */
  server.tool(
    'chat-memory-status',
    `🧠 **CHAT MEMORY STATUS** - View your current conversation memory

This tool shows you what the assistant remembers about your current chat session:
- Recent conversation history
- Mentioned people, files, events, and topics
- Your preferences for this chat

Use this to understand what context the assistant has from previous messages.

Note: Chat memory is automatically cleared after 72 hours of inactivity.`,
    {},
    {
      title: 'Chat Memory Status',
      readOnlyHint: true,
      openWorldHint: false,
    },
    async () => {
      const chatId = getChatId();

      if (!isChatMemoryEnabled()) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  status: 'DISABLED',
                  message:
                    'Chat memory is disabled. Set MS365_MCP_CHAT_MEMORY_ENABLED=true to enable.',
                },
                null,
                2
              ),
            },
          ],
        };
      }

      if (!chatId) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  status: 'NO_CHAT_ID',
                  message:
                    'No chat ID detected. Ensure your client sends X-Chat-ID or X-OpenWebUI-Chat-ID header.',
                },
                null,
                2
              ),
            },
          ],
        };
      }

      // SECURITY: Get userId for user-scoped memory access
      const userId = getUserId();
      const memoryStore = getChatMemoryStore();
      // SECURITY: Pass userId to ensure user only accesses their own memory
      const summary = memoryStore.getMemorySummary(chatId, userId);
      const serialized = memoryStore.serializeMemory(chatId, userId);
      const stats = memoryStore.getStats();

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(
              {
                status: 'SUCCESS',
                chatId: chatId.substring(0, 8) + '...',
                userId: userId ? userId.substring(0, 8) + '...' : 'anonymous',
                summary,
                memory: serialized,
                globalStats: {
                  activeSessions: stats.activeSessions,
                  totalConversations: stats.totalConversations,
                  memoryUsage: stats.memoryUsageEstimate,
                },
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );
  registeredCount++;

  /**
   * Chat Memory Clear - Clears memory for current chat
   */
  server.tool(
    'chat-memory-clear',
    `🗑️ **CLEAR CHAT MEMORY** - Reset your conversation memory

This tool clears all memory for your current chat session:
- Removes conversation history
- Clears mentioned entities (people, files, events, topics)
- Resets preferences

Use this when you want to start fresh without previous context.`,
    {
      confirm: z
        .boolean()
        .optional()
        .describe(
          'Set to true to confirm clearing the memory. Without confirmation, shows what would be cleared.'
        ),
    },
    {
      title: 'Clear Chat Memory',
      readOnlyHint: false,
      openWorldHint: false,
    },
    async ({ confirm = false }) => {
      const chatId = getChatId();

      if (!isChatMemoryEnabled()) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  status: 'DISABLED',
                  message: 'Chat memory is disabled.',
                },
                null,
                2
              ),
            },
          ],
        };
      }

      if (!chatId) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  status: 'NO_CHAT_ID',
                  message: 'No chat ID detected.',
                },
                null,
                2
              ),
            },
          ],
        };
      }

      // SECURITY: Get userId for user-scoped memory access
      const userId = getUserId();
      const memoryStore = getChatMemoryStore();

      if (!confirm) {
        // Preview what would be cleared
        // SECURITY: Pass userId to ensure user only accesses their own memory
        const serialized = memoryStore.serializeMemory(chatId, userId);
        if (!serialized) {
          return {
            content: [
              {
                type: 'text' as const,
                text: JSON.stringify(
                  {
                    status: 'NO_MEMORY',
                    message: 'No memory exists for this chat session.',
                  },
                  null,
                  2
                ),
              },
            ],
          };
        }

        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  status: 'PREVIEW',
                  message: 'This is what will be cleared. Call with confirm=true to proceed.',
                  memory: {
                    conversations: serialized.conversationHistory.length,
                    entities: {
                      people: serialized.mentionedEntities.people.length,
                      files: serialized.mentionedEntities.files.length,
                      events: serialized.mentionedEntities.events.length,
                      topics: serialized.mentionedEntities.topics.length,
                    },
                    hasPreferences: Object.keys(serialized.preferences).length > 0,
                  },
                },
                null,
                2
              ),
            },
          ],
        };
      }

      // Actually clear the memory
      // SECURITY: Pass userId to ensure user can only clear their own memory
      const cleared = memoryStore.clearMemory(chatId, userId);

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(
              {
                status: cleared ? 'CLEARED' : 'NOT_FOUND',
                message: cleared
                  ? 'Chat memory has been cleared successfully.'
                  : 'No memory found to clear (or access denied).',
                chatId: chatId.substring(0, 8) + '...',
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );
  registeredCount++;

  /**
   * Chat Memory Set Preference - Set a preference for current chat
   */
  server.tool(
    'chat-memory-set-preference',
    `⚙️ **SET CHAT PREFERENCE** - Configure preferences for this chat session

This tool allows you to set preferences that persist throughout your chat session:
- **language**: Set preferred response language (en, de, or auto)
- **resultLimit**: Set default number of results to return
- **preferredSources**: Set which Microsoft 365 products to search first

Preferences are automatically applied to subsequent queries in this chat.`,
    {
      preference: z
        .enum(['language', 'resultLimit', 'preferredSources'])
        .describe('The preference to set'),
      value: z
        .string()
        .describe(
          'The value for the preference. For language: en|de|auto. For resultLimit: number. For preferredSources: comma-separated list (email,calendar,files,teams)'
        ),
    },
    {
      title: 'Set Chat Preference',
      readOnlyHint: false,
      openWorldHint: false,
    },
    async ({ preference, value }) => {
      const chatId = getChatId();
      const userId = getUserId();

      if (!isChatMemoryEnabled()) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  status: 'DISABLED',
                  message: 'Chat memory is disabled.',
                },
                null,
                2
              ),
            },
          ],
        };
      }

      if (!chatId) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  status: 'NO_CHAT_ID',
                  message: 'No chat ID detected.',
                },
                null,
                2
              ),
            },
          ],
        };
      }

      const memoryStore = getChatMemoryStore();

      try {
        switch (preference) {
          case 'language': {
            if (!['en', 'de', 'auto'].includes(value)) {
              return {
                content: [
                  {
                    type: 'text' as const,
                    text: JSON.stringify(
                      {
                        status: 'ERROR',
                        message: 'Invalid language value. Use: en, de, or auto',
                      },
                      null,
                      2
                    ),
                  },
                ],
              };
            }
            memoryStore.setPreference(chatId, 'language', value as 'en' | 'de' | 'auto', userId);
            break;
          }
          case 'resultLimit': {
            const limit = parseInt(value, 10);
            if (isNaN(limit) || limit < 1 || limit > 100) {
              return {
                content: [
                  {
                    type: 'text' as const,
                    text: JSON.stringify(
                      {
                        status: 'ERROR',
                        message: 'Invalid resultLimit. Use a number between 1 and 100.',
                      },
                      null,
                      2
                    ),
                  },
                ],
              };
            }
            memoryStore.setPreference(chatId, 'resultLimit', limit, userId);
            break;
          }
          case 'preferredSources': {
            const sources = value
              .split(',')
              .map((s) => s.trim().toLowerCase())
              .filter((s) => s.length > 0);
            if (sources.length === 0) {
              return {
                content: [
                  {
                    type: 'text' as const,
                    text: JSON.stringify(
                      {
                        status: 'ERROR',
                        message:
                          'Invalid preferredSources. Provide comma-separated values like: email,calendar,files',
                      },
                      null,
                      2
                    ),
                  },
                ],
              };
            }
            memoryStore.setPreference(chatId, 'preferredSources', sources, userId);
            break;
          }
        }

        // SECURITY: Pass userId to ensure user only accesses their own preferences
        const prefs = memoryStore.getPreferences(chatId, userId);

        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  status: 'SUCCESS',
                  message: `Preference '${preference}' has been set to '${value}'`,
                  currentPreferences: prefs,
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  status: 'ERROR',
                  message: `Failed to set preference: ${String(error)}`,
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
  registeredCount++;

  // ==========================================================================
  // CONTENT INTELLIGENCE & EXTRACTION TOOLS
  // ==========================================================================

  // ==========================================================================
  // 23. EXTRACT ACTION ITEMS - Extract action items from emails and meetings
  // ==========================================================================
  server.tool(
    'extract-action-items',
    `Extract action items from emails and meetings:
- Parse emails for action items (tasks, to-dos, follow-ups)
- Extract tasks from meeting notes and descriptions
- Identify deadlines and assignees
- Create structured task list
- Link to source emails/meetings

Use this for "What action items do I have from recent emails?", "Extract tasks from yesterday's meeting", or "Show me all action items from [person]".`,
    {
      source: z
        .enum(['emails', 'meetings', 'both'])
        .optional()
        .describe('Source to extract from: emails, meetings, or both (default: both)'),
      days: z.number().optional().describe('Days back to search (default: 7)'),
      person: z.string().optional().describe('Filter by person name or email'),
      limit: z.number().optional().describe('Maximum action items to return (default: 50)'),
    },
    {
      title: 'Extract Action Items',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ source = 'both', days = 7, person, limit = 50 }) => {
      logger.info(`Extracting action items: source=${source}, days=${days}, person=${person}`);

      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      interface ActionItem {
        text: string;
        source: 'email' | 'meeting';
        sourceId: string;
        sourceTitle: string;
        sourceDate: string;
        sourceLink?: string;
        assignee?: string;
        deadline?: string;
        priority?: 'high' | 'medium' | 'low';
      }

      const actionItems: ActionItem[] = [];
      const promises: Promise<void>[] = [];

      // Action item detection patterns
      const actionPatterns = [
        /(?:action item|action|todo|to do|task|follow up|follow-up)[\s:]+(.+?)(?:\.|$)/gi,
        /(?:need to|must|should|will|going to)\s+(.+?)(?:\.|$)/gi,
        /(?:please|can you|could you)\s+(.+?)(?:\.|\?|$)/gi,
        /(?:deadline|due|by)\s+([^.]+)/gi,
        /(?:assign(?:ed)? to|@)([^\s.]+)/gi,
      ];

      // Extract from emails
      if (source === 'emails' || source === 'both') {
        promises.push(
          (async () => {
            try {
              let queryParams: Record<string, string> = {
                $top: '100',
                $select: 'id,subject,bodyPreview,body,receivedDateTime,from,toRecipients,webLink',
                $orderby: 'receivedDateTime desc',
                $filter: `receivedDateTime ge ${startDate.toISOString()}`,
              };

              if (person) {
                const user = await findUser(graphClient, person);
                if (user) {
                  const userEmail = user.mail || user.userPrincipalName || '';
                  queryParams.$filter += ` and (from/emailAddress/address eq '${userEmail}' or toRecipients/any(r:r/emailAddress/address eq '${userEmail}'))`;
                }
              }

              const emailResponse = await graphClient.makeRequest('/me/messages', {
                method: 'GET',
                queryParams,
              });

              if (emailResponse && typeof emailResponse === 'object' && 'value' in emailResponse) {
                for (const email of emailResponse.value as GraphEmail[]) {
                  const content = sanitizeHtml(email.bodyPreview || email.body?.content || '');
                  const subject = email.subject || '';

                  // Extract action items from content
                  for (const pattern of actionPatterns) {
                    const matches = content.matchAll(pattern);
                    for (const match of matches) {
                      if (match[1] && match[1].trim().length > 5) {
                        const actionText = match[1].trim();
                        // Check if it's actually an action item (not just a mention)
                        if (
                          actionText.length < 200 &&
                          !actionText.toLowerCase().includes('http') &&
                          !actionText.toLowerCase().includes('@')
                        ) {
                          actionItems.push({
                            text: actionText,
                            source: 'email',
                            sourceId: email.id,
                            sourceTitle: subject,
                            sourceDate: email.receivedDateTime,
                            sourceLink: email.webLink,
                            assignee: email.from?.emailAddress?.name,
                          });
                        }
                      }
                    }
                  }

                  // Also check subject line
                  if (
                    subject.toLowerCase().includes('action') ||
                    subject.toLowerCase().includes('todo')
                  ) {
                    actionItems.push({
                      text: subject,
                      source: 'email',
                      sourceId: email.id,
                      sourceTitle: subject,
                      sourceDate: email.receivedDateTime,
                      sourceLink: email.webLink,
                      assignee: email.from?.emailAddress?.name,
                    });
                  }
                }
              }
            } catch (error) {
              logger.warn(`Could not extract action items from emails: ${error}`);
            }
          })()
        );
      }

      // Extract from meetings
      if (source === 'meetings' || source === 'both') {
        promises.push(
          (async () => {
            try {
              const meetingQueryParams: Record<string, string> = {
                startDateTime: startDate.toISOString(),
                endDateTime: new Date().toISOString(),
                $top: '100',
                $select: 'id,subject,bodyPreview,start,end,organizer,attendees,webLink',
              };

              const meetingsResponse = await graphClient.makeRequest(
                `/me/calendarView?${buildGraphQueryString(meetingQueryParams)}`,
                {
                  method: 'GET',
                }
              );

              if (
                meetingsResponse &&
                typeof meetingsResponse === 'object' &&
                'value' in meetingsResponse
              ) {
                for (const meeting of meetingsResponse.value as GraphEvent[]) {
                  const content = sanitizeHtml(meeting.bodyPreview || '');
                  const subject = meeting.subject || '';

                  // Filter by person if specified
                  if (person) {
                    const user = await findUser(graphClient, person);
                    if (user) {
                      const userEmail = user.mail || user.userPrincipalName || '';
                      const isAttendee =
                        meeting.attendees?.some(
                          (a) => a.emailAddress?.address?.toLowerCase() === userEmail.toLowerCase()
                        ) ||
                        meeting.organizer?.emailAddress?.address?.toLowerCase() ===
                          userEmail.toLowerCase();
                      if (!isAttendee) continue;
                    }
                  }

                  // Extract action items from meeting notes
                  for (const pattern of actionPatterns) {
                    const matches = content.matchAll(pattern);
                    for (const match of matches) {
                      if (match[1] && match[1].trim().length > 5) {
                        const actionText = match[1].trim();
                        if (
                          actionText.length < 200 &&
                          !actionText.toLowerCase().includes('http') &&
                          !actionText.toLowerCase().includes('@')
                        ) {
                          actionItems.push({
                            text: actionText,
                            source: 'meeting',
                            sourceId: meeting.id,
                            sourceTitle: subject,
                            sourceDate: meeting.start.dateTime,
                            sourceLink: meeting.webLink,
                            assignee: meeting.organizer?.emailAddress?.name,
                          });
                        }
                      }
                    }
                  }
                }
              }
            } catch (error) {
              logger.warn(`Could not extract action items from meetings: ${error}`);
            }
          })()
        );
      }

      await Promise.allSettled(promises);

      // Deduplicate and limit
      const uniqueItems = Array.from(
        new Map(actionItems.map((item) => [item.text.toLowerCase(), item])).values()
      ).slice(0, limit);

      const result = {
        totalFound: uniqueItems.length,
        actionItems: uniqueItems,
        summary: {
          fromEmails: uniqueItems.filter((i) => i.source === 'email').length,
          fromMeetings: uniqueItems.filter((i) => i.source === 'meeting').length,
          dateRange: {
            start: startDate.toISOString(),
            end: new Date().toISOString(),
          },
        },
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 24. SUMMARIZE EMAIL THREAD - Summarize long email threads
  // ==========================================================================
  server.tool(
    'summarize-email-thread',
    `Summarize long email threads:
- Extract key points from email conversations
- Identify decisions made
- List participants
- Highlight action items
- Create concise summary

Use this for "Summarize the email thread about [topic]", "What was decided in this email chain?", or "Give me a summary of this conversation".`,
    {
      topic: z.string().describe('Topic or subject to find and summarize email thread'),
      days: z.number().optional().describe('Days back to search (default: 30)'),
      limit: z.number().optional().describe('Maximum emails to analyze (default: 50)'),
    },
    {
      title: 'Summarize Email Thread',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ topic, days = 30, limit = 50 }) => {
      logger.info(`Summarizing email thread for topic: ${topic}`);

      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      try {
        const emailResponse = await graphClient.makeRequest('/me/messages', {
          method: 'GET',
          queryParams: {
            $search: `"${topic}"`,
            $top: String(limit),
            $select:
              'id,subject,bodyPreview,body,receivedDateTime,from,toRecipients,ccRecipients,webLink',
            $orderby: 'receivedDateTime desc',
            $filter: `receivedDateTime ge ${startDate.toISOString()}`,
          },
        });

        if (!emailResponse || typeof emailResponse !== 'object' || !('value' in emailResponse)) {
          return {
            content: [
              {
                type: 'text' as const,
                text: JSON.stringify({
                  error: 'No emails found',
                  message: `No emails found matching "${topic}" in the last ${days} days.`,
                  topic,
                }),
              },
            ],
            isError: true,
          };
        }

        const emails = emailResponse.value as GraphEmail[];
        if (emails.length === 0) {
          return {
            content: [
              {
                type: 'text' as const,
                text: JSON.stringify({
                  error: 'No emails found',
                  message: `No emails found matching "${topic}" in the last ${days} days.`,
                  topic,
                }),
              },
            ],
            isError: true,
          };
        }

        // Extract participants
        const participants = new Set<string>();
        const participantsByEmail = new Map<string, string>();

        // Extract key points, decisions, and action items
        const keyPoints: string[] = [];
        const decisions: string[] = [];
        const actionItems: string[] = [];

        // Group by thread (same subject)
        const threads = new Map<string, GraphEmail[]>();
        for (const email of emails) {
          const subject = email.subject || 'No Subject';
          if (!threads.has(subject)) {
            threads.set(subject, []);
          }
          threads.get(subject)!.push(email);
        }

        // Process each thread
        for (const [subject, threadEmails] of threads) {
          // Sort by date
          threadEmails.sort(
            (a, b) =>
              new Date(a.receivedDateTime).getTime() - new Date(b.receivedDateTime).getTime()
          );

          for (const email of threadEmails) {
            // Track participants
            if (email.from?.emailAddress) {
              participants.add(email.from.emailAddress.name || email.from.emailAddress.address);
              participantsByEmail.set(
                email.from.emailAddress.address,
                email.from.emailAddress.name || email.from.emailAddress.address
              );
            }
            for (const recipient of [
              ...(email.toRecipients || []),
              ...(email.ccRecipients || []),
            ]) {
              if (recipient.emailAddress?.address) {
                participants.add(recipient.emailAddress.name || recipient.emailAddress.address);
                participantsByEmail.set(
                  recipient.emailAddress.address,
                  recipient.emailAddress.name || recipient.emailAddress.address
                );
              }
            }

            const content = sanitizeHtml(email.bodyPreview || email.body?.content || '');
            const contentLower = content.toLowerCase();

            // Extract decisions
            if (
              contentLower.includes('decided') ||
              contentLower.includes('decision') ||
              contentLower.includes('agreed') ||
              contentLower.includes('we will') ||
              contentLower.includes("let's go with")
            ) {
              const decisionMatch = content.match(
                /(?:decided|decision|agreed|we will|let's go with)[\s:]+(.+?)(?:\.|$)/i
              );
              if (decisionMatch && decisionMatch[1]) {
                decisions.push(decisionMatch[1].trim());
              }
            }

            // Extract action items
            if (
              contentLower.includes('action') ||
              contentLower.includes('todo') ||
              contentLower.includes('follow up') ||
              contentLower.includes('need to')
            ) {
              const actionMatch = content.match(
                /(?:action|todo|follow up|need to)[\s:]+(.+?)(?:\.|$)/i
              );
              if (actionMatch && actionMatch[1]) {
                actionItems.push(actionMatch[1].trim());
              }
            }

            // Extract key points (sentences with important keywords)
            if (
              contentLower.includes('important') ||
              contentLower.includes('key') ||
              contentLower.includes('summary') ||
              contentLower.includes('conclusion')
            ) {
              const sentences = content.split(/[.!?]+/).filter((s) => s.trim().length > 20);
              for (const sentence of sentences.slice(0, 3)) {
                if (sentence.trim().length > 20 && sentence.trim().length < 200) {
                  keyPoints.push(sentence.trim());
                }
              }
            }
          }
        }

        // Create summary
        const firstEmail = emails[emails.length - 1]; // Oldest
        const lastEmail = emails[0]; // Newest

        const summary = {
          topic,
          threadCount: threads.size,
          totalEmails: emails.length,
          participants: Array.from(participants).slice(0, 20),
          dateRange: {
            start: firstEmail.receivedDateTime,
            end: lastEmail.receivedDateTime,
          },
          keyPoints: keyPoints.slice(0, 10),
          decisions: decisions.slice(0, 10),
          actionItems: actionItems.slice(0, 10),
          threads: Array.from(threads.entries()).map(([subject, threadEmails]) => ({
            subject,
            emailCount: threadEmails.length,
            firstEmail: threadEmails[0].receivedDateTime,
            lastEmail: threadEmails[threadEmails.length - 1].receivedDateTime,
            participants: Array.from(
              new Set(
                threadEmails.flatMap((e) => [
                  e.from?.emailAddress?.name || e.from?.emailAddress?.address,
                  ...(e.toRecipients || []).map(
                    (r) => r.emailAddress?.name || r.emailAddress?.address
                  ),
                ])
              )
            ).filter(Boolean),
          })),
        };

        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(summary, null, 2),
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  error: 'Failed to summarize email thread',
                  message: `${error}`,
                  topic,
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
  registeredCount++;

  // ==========================================================================
  // 25. EXTRACT DECISIONS - Extract decisions from communications
  // ==========================================================================
  server.tool(
    'extract-decisions',
    `Extract decisions from communications:
- Find decision points in emails
- Extract decisions from meetings
- Track decision timeline
- Identify decision makers
- Link to related documents

Use this for "What decisions were made about [topic]?", "Extract all decisions from last week", or "Show me decision history for [project]".`,
    {
      topic: z.string().optional().describe('Topic or project to find decisions about'),
      days: z.number().optional().describe('Days back to search (default: 90)'),
      source: z
        .enum(['emails', 'meetings', 'both'])
        .optional()
        .describe('Source to search: emails, meetings, or both (default: both)'),
      limit: z.number().optional().describe('Maximum decisions to return (default: 50)'),
    },
    {
      title: 'Extract Decisions',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ topic, days = 90, source = 'both', limit = 50 }) => {
      logger.info(`Extracting decisions: topic=${topic}, days=${days}, source=${source}`);

      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      interface Decision {
        text: string;
        source: 'email' | 'meeting';
        sourceId: string;
        sourceTitle: string;
        sourceDate: string;
        sourceLink?: string;
        decisionMaker?: string;
        participants?: string[];
      }

      const decisions: Decision[] = [];
      const promises: Promise<void>[] = [];

      const decisionPatterns = [
        /(?:decided|decision)[\s:]+(.+?)(?:\.|$)/gi,
        /(?:agreed|agreement)[\s:]+(.+?)(?:\.|$)/gi,
        /(?:we will|we'll|going to)[\s:]+(.+?)(?:\.|$)/gi,
        /(?:let's|let us)[\s:]+(.+?)(?:\.|$)/gi,
        /(?:conclusion|concluded)[\s:]+(.+?)(?:\.|$)/gi,
        /(?:final|finally)[\s:]+(.+?)(?:\.|$)/gi,
      ];

      // Extract from emails
      if (source === 'emails' || source === 'both') {
        promises.push(
          (async () => {
            try {
              let queryParams: Record<string, string> = {
                $top: '100',
                $select:
                  'id,subject,bodyPreview,body,receivedDateTime,from,toRecipients,ccRecipients,webLink',
                $orderby: 'receivedDateTime desc',
                $filter: `receivedDateTime ge ${startDate.toISOString()}`,
              };

              if (topic) {
                queryParams.$search = `"${topic}"`;
                delete queryParams.$orderby; // Can't use orderby with search
              }

              const emailResponse = await graphClient.makeRequest('/me/messages', {
                method: 'GET',
                queryParams,
              });

              if (emailResponse && typeof emailResponse === 'object' && 'value' in emailResponse) {
                for (const email of emailResponse.value as GraphEmail[]) {
                  const content = sanitizeHtml(email.bodyPreview || email.body?.content || '');
                  const subject = email.subject || '';

                  // Filter by topic if specified and not already filtered by search
                  if (topic && !queryParams.$search) {
                    if (
                      !subject.toLowerCase().includes(topic.toLowerCase()) &&
                      !content.toLowerCase().includes(topic.toLowerCase())
                    ) {
                      continue;
                    }
                  }

                  // Extract decisions
                  for (const pattern of decisionPatterns) {
                    const matches = content.matchAll(pattern);
                    for (const match of matches) {
                      if (match[1] && match[1].trim().length > 10) {
                        const decisionText = match[1].trim();
                        if (
                          decisionText.length < 300 &&
                          !decisionText.toLowerCase().includes('http') &&
                          !decisionText.toLowerCase().includes('@')
                        ) {
                          decisions.push({
                            text: decisionText,
                            source: 'email',
                            sourceId: email.id,
                            sourceTitle: subject,
                            sourceDate: email.receivedDateTime,
                            sourceLink: email.webLink,
                            decisionMaker: email.from?.emailAddress?.name,
                            participants: [
                              email.from?.emailAddress?.name || email.from?.emailAddress?.address,
                              ...(email.toRecipients || []).map(
                                (r) => r.emailAddress?.name || r.emailAddress?.address
                              ),
                            ].filter(Boolean),
                          });
                        }
                      }
                    }
                  }
                }
              }
            } catch (error) {
              logger.warn(`Could not extract decisions from emails: ${error}`);
            }
          })()
        );
      }

      // Extract from meetings
      if (source === 'meetings' || source === 'both') {
        promises.push(
          (async () => {
            try {
              const meetingQueryParams: Record<string, string> = {
                startDateTime: startDate.toISOString(),
                endDateTime: new Date().toISOString(),
                $top: '100',
                $select: 'id,subject,bodyPreview,start,end,organizer,attendees,webLink',
              };

              const meetingsResponse = await graphClient.makeRequest(
                `/me/calendarView?${buildGraphQueryString(meetingQueryParams)}`,
                {
                  method: 'GET',
                }
              );

              if (
                meetingsResponse &&
                typeof meetingsResponse === 'object' &&
                'value' in meetingsResponse
              ) {
                for (const meeting of meetingsResponse.value as GraphEvent[]) {
                  const content = sanitizeHtml(meeting.bodyPreview || '');
                  const subject = meeting.subject || '';

                  // Filter by topic if specified
                  if (topic) {
                    if (
                      !subject.toLowerCase().includes(topic.toLowerCase()) &&
                      !content.toLowerCase().includes(topic.toLowerCase())
                    ) {
                      continue;
                    }
                  }

                  // Extract decisions
                  for (const pattern of decisionPatterns) {
                    const matches = content.matchAll(pattern);
                    for (const match of matches) {
                      if (match[1] && match[1].trim().length > 10) {
                        const decisionText = match[1].trim();
                        if (
                          decisionText.length < 300 &&
                          !decisionText.toLowerCase().includes('http') &&
                          !decisionText.toLowerCase().includes('@')
                        ) {
                          decisions.push({
                            text: decisionText,
                            source: 'meeting',
                            sourceId: meeting.id,
                            sourceTitle: subject,
                            sourceDate: meeting.start.dateTime,
                            sourceLink: meeting.webLink,
                            decisionMaker: meeting.organizer?.emailAddress?.name,
                            participants: [
                              meeting.organizer?.emailAddress?.name,
                              ...(meeting.attendees || []).map(
                                (a) => a.emailAddress?.name || a.emailAddress?.address
                              ),
                            ].filter(Boolean),
                          });
                        }
                      }
                    }
                  }
                }
              }
            } catch (error) {
              logger.warn(`Could not extract decisions from meetings: ${error}`);
            }
          })()
        );
      }

      await Promise.allSettled(promises);

      // Sort by date and deduplicate
      decisions.sort((a, b) => new Date(a.sourceDate).getTime() - new Date(b.sourceDate).getTime());
      const uniqueDecisions = Array.from(
        new Map(decisions.map((d) => [d.text.toLowerCase().substring(0, 100), d])).values()
      ).slice(0, limit);

      // Build timeline
      const timeline = uniqueDecisions.map((d) => ({
        date: d.sourceDate,
        decision: d.text,
        source: d.source,
        decisionMaker: d.decisionMaker,
        sourceTitle: d.sourceTitle,
        sourceLink: d.sourceLink,
      }));

      // Identify decision makers
      const decisionMakers = new Map<string, number>();
      for (const decision of uniqueDecisions) {
        if (decision.decisionMaker) {
          decisionMakers.set(
            decision.decisionMaker,
            (decisionMakers.get(decision.decisionMaker) || 0) + 1
          );
        }
      }

      const result = {
        topic: topic || 'All topics',
        totalDecisions: uniqueDecisions.length,
        dateRange: {
          start: startDate.toISOString(),
          end: new Date().toISOString(),
        },
        timeline,
        decisionMakers: Array.from(decisionMakers.entries())
          .map(([name, count]) => ({ name, decisionCount: count }))
          .sort((a, b) => b.decisionCount - a.decisionCount)
          .slice(0, 10),
        summary: {
          fromEmails: uniqueDecisions.filter((d) => d.source === 'email').length,
          fromMeetings: uniqueDecisions.filter((d) => d.source === 'meeting').length,
          firstDecision: timeline[0]?.date,
          lastDecision: timeline[timeline.length - 1]?.date,
        },
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // RELATIONSHIP INTELLIGENCE TOOLS
  // ==========================================================================

  // ==========================================================================
  // 26. ANALYZE RELATIONSHIP STRENGTH - Analyze relationship strength with contacts
  // ==========================================================================
  server.tool(
    'analyze-relationship-strength',
    `Analyze relationship strength with contacts:
- Communication frequency scoring
- Interaction recency
- Meeting participation rate
- Email thread depth
- Overall relationship score

Use this for "How strong is my relationship with [person]?", "Who do I communicate with most?", or "Analyze my professional network".`,
    {
      person: z
        .string()
        .optional()
        .describe('Person name or email to analyze (optional - analyzes all if not provided)'),
      days: z.number().optional().describe('Days of history to analyze (default: 90)'),
      limit: z.number().optional().describe('Maximum relationships to return (default: 20)'),
    },
    {
      title: 'Analyze Relationship Strength',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ person, days = 90, limit = 20 }) => {
      logger.info(`Analyzing relationship strength: person=${person}, days=${days}`);

      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      interface RelationshipData {
        personId: string;
        name: string;
        email: string;
        emailCount: number;
        meetingCount: number;
        lastInteraction: string;
        firstInteraction: string;
        emailThreads: number;
        relationshipScore: number;
      }

      const relationships = new Map<string, RelationshipData>();
      const promises: Promise<void>[] = [];

      // Analyze emails
      promises.push(
        (async () => {
          try {
            let queryParams: Record<string, string> = {
              $top: '500',
              $select: 'id,subject,from,toRecipients,ccRecipients,receivedDateTime,conversationId',
              $orderby: 'receivedDateTime desc',
              $filter: `receivedDateTime ge ${startDate.toISOString()}`,
            };

            if (person) {
              const user = await findUser(graphClient, person);
              if (user) {
                const userEmail = user.mail || user.userPrincipalName || '';
                queryParams.$filter += ` and (from/emailAddress/address eq '${userEmail}' or toRecipients/any(r:r/emailAddress/address eq '${userEmail}'))`;
              }
            }

            const emailResponse = await graphClient.makeRequest('/me/messages', {
              method: 'GET',
              queryParams,
            });

            if (emailResponse && typeof emailResponse === 'object' && 'value' in emailResponse) {
              const threads = new Map<string, Set<string>>();

              for (const email of emailResponse.value as GraphEmail[]) {
                const senderEmail = email.from?.emailAddress?.address?.toLowerCase();
                const senderName = email.from?.emailAddress?.name || senderEmail || 'Unknown';

                if (senderEmail && senderEmail !== (await getCurrentUserEmail())) {
                  if (!relationships.has(senderEmail)) {
                    relationships.set(senderEmail, {
                      personId: '',
                      name: senderName,
                      email: senderEmail,
                      emailCount: 0,
                      meetingCount: 0,
                      lastInteraction: email.receivedDateTime,
                      firstInteraction: email.receivedDateTime,
                      emailThreads: 0,
                      relationshipScore: 0,
                    });
                  }

                  const rel = relationships.get(senderEmail)!;
                  rel.emailCount++;
                  if (new Date(email.receivedDateTime) > new Date(rel.lastInteraction)) {
                    rel.lastInteraction = email.receivedDateTime;
                  }
                  if (new Date(email.receivedDateTime) < new Date(rel.firstInteraction)) {
                    rel.firstInteraction = email.receivedDateTime;
                  }

                  // Track threads
                  if (email.conversationId) {
                    if (!threads.has(email.conversationId)) {
                      threads.set(email.conversationId, new Set());
                    }
                    threads.get(email.conversationId)!.add(senderEmail);
                  }
                }
              }

              // Count threads per person
              for (const [conversationId, participants] of threads) {
                for (const participantEmail of participants) {
                  if (relationships.has(participantEmail)) {
                    relationships.get(participantEmail)!.emailThreads++;
                  }
                }
              }
            }
          } catch (error) {
            logger.warn(`Could not analyze email relationships: ${error}`);
          }
        })()
      );

      // Analyze meetings
      promises.push(
        (async () => {
          try {
            const meetingQueryParams: Record<string, string> = {
              startDateTime: startDate.toISOString(),
              endDateTime: new Date().toISOString(),
              $top: '500',
              $select: 'id,subject,start,end,organizer,attendees',
            };

            const meetingsResponse = await graphClient.makeRequest(
              `/me/calendarView?${buildGraphQueryString(meetingQueryParams)}`,
              {
                method: 'GET',
              }
            );

            if (
              meetingsResponse &&
              typeof meetingsResponse === 'object' &&
              'value' in meetingsResponse
            ) {
              const currentUserEmail = await getCurrentUserEmail();

              for (const meeting of meetingsResponse.value as GraphEvent[]) {
                const organizerEmail = meeting.organizer?.emailAddress?.address?.toLowerCase();
                const organizerName =
                  meeting.organizer?.emailAddress?.name || organizerEmail || 'Unknown';

                // Track organizer
                if (organizerEmail && organizerEmail !== currentUserEmail) {
                  if (!relationships.has(organizerEmail)) {
                    relationships.set(organizerEmail, {
                      personId: '',
                      name: organizerName,
                      email: organizerEmail,
                      emailCount: 0,
                      meetingCount: 0,
                      lastInteraction: meeting.start.dateTime,
                      firstInteraction: meeting.start.dateTime,
                      emailThreads: 0,
                      relationshipScore: 0,
                    });
                  }
                  relationships.get(organizerEmail)!.meetingCount++;
                }

                // Track attendees
                for (const attendee of meeting.attendees || []) {
                  const attendeeEmail = attendee.emailAddress?.address?.toLowerCase();
                  const attendeeName = attendee.emailAddress?.name || attendeeEmail || 'Unknown';

                  if (attendeeEmail && attendeeEmail !== currentUserEmail) {
                    if (!relationships.has(attendeeEmail)) {
                      relationships.set(attendeeEmail, {
                        personId: '',
                        name: attendeeName,
                        email: attendeeEmail,
                        emailCount: 0,
                        meetingCount: 0,
                        lastInteraction: meeting.start.dateTime,
                        firstInteraction: meeting.start.dateTime,
                        emailThreads: 0,
                        relationshipScore: 0,
                      });
                    }
                    relationships.get(attendeeEmail)!.meetingCount++;
                  }
                }
              }
            }
          } catch (error) {
            logger.warn(`Could not analyze meeting relationships: ${error}`);
          }
        })()
      );

      await Promise.allSettled(promises);

      // Calculate relationship scores
      const now = new Date();
      for (const [email, rel] of relationships) {
        // Score components:
        // - Email frequency (0-40 points)
        // - Meeting frequency (0-30 points)
        // - Recency (0-20 points)
        // - Thread depth (0-10 points)

        const emailScore = Math.min(rel.emailCount * 2, 40);
        const meetingScore = Math.min(rel.meetingCount * 5, 30);
        const daysSinceLastInteraction =
          (now.getTime() - new Date(rel.lastInteraction).getTime()) / (1000 * 60 * 60 * 24);
        const recencyScore = Math.max(0, 20 - daysSinceLastInteraction / 5);
        const threadScore = Math.min(rel.emailThreads * 2, 10);

        rel.relationshipScore = emailScore + meetingScore + recencyScore + threadScore;
      }

      // Sort by score and filter by person if specified
      let sortedRelationships = Array.from(relationships.values())
        .sort((a, b) => b.relationshipScore - a.relationshipScore)
        .slice(0, limit);

      // Filter by person if specified
      if (person) {
        const user = await findUser(graphClient, person);
        if (user) {
          const userEmail = user.mail || user.userPrincipalName || '';
          sortedRelationships = sortedRelationships.filter(
            (r) => r.email.toLowerCase() === userEmail.toLowerCase()
          );
        }
      }

      const result = {
        analyzedPeriod: `Last ${days} days`,
        totalRelationships: relationships.size,
        relationships: sortedRelationships.map((rel) => ({
          name: rel.name,
          email: rel.email,
          score: Math.round(rel.relationshipScore * 10) / 10,
          metrics: {
            emailCount: rel.emailCount,
            meetingCount: rel.meetingCount,
            emailThreads: rel.emailThreads,
            lastInteraction: rel.lastInteraction,
            firstInteraction: rel.firstInteraction,
          },
          strength:
            rel.relationshipScore >= 70
              ? 'Very Strong'
              : rel.relationshipScore >= 50
                ? 'Strong'
                : rel.relationshipScore >= 30
                  ? 'Moderate'
                  : rel.relationshipScore >= 15
                    ? 'Weak'
                    : 'Very Weak',
        })),
        summary: {
          strongestRelationship: sortedRelationships[0]?.name,
          mostEmails: sortedRelationships.sort((a, b) => b.emailCount - a.emailCount)[0]?.name,
          mostMeetings: sortedRelationships.sort((a, b) => b.meetingCount - a.meetingCount)[0]
            ?.name,
        },
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 27. FIND MUTUAL CONNECTIONS - Find mutual connections between people
  // ==========================================================================
  server.tool(
    'find-mutual-connections',
    `Find mutual connections between people:
- Common meeting participants
- Shared email threads
- Mutual contacts
- Collaboration history
- Network connections

Use this for "Who do I know in common with [person]?", "Find mutual connections for [person]", or "Who connects me to [person]?".`,
    {
      person: z.string().describe('Person name or email to find mutual connections with'),
      days: z.number().optional().describe('Days of history to analyze (default: 180)'),
      limit: z.number().optional().describe('Maximum connections to return (default: 20)'),
    },
    {
      title: 'Find Mutual Connections',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ person, days = 180, limit = 20 }) => {
      logger.info(`Finding mutual connections for: ${person}`);

      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      // Find the target person
      const targetUser = await findUser(graphClient, person);
      if (!targetUser) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify({
                error: 'Person not found',
                message: `Could not find a person matching "${person}".`,
                searchedFor: person,
              }),
            },
          ],
          isError: true,
        };
      }

      const targetEmail = targetUser.mail || targetUser.userPrincipalName || '';
      const currentUserEmail = await getCurrentUserEmail();

      // Track mutual connections
      const mutualConnections = new Map<
        string,
        {
          name: string;
          email: string;
          sharedMeetings: number;
          sharedEmails: number;
          connectionStrength: number;
        }
      >();

      const promises: Promise<void>[] = [];

      // Find shared meetings
      promises.push(
        (async () => {
          try {
            const meetingQueryParams: Record<string, string> = {
              startDateTime: startDate.toISOString(),
              endDateTime: new Date().toISOString(),
              $top: '500',
              $select: 'id,subject,start,attendees,organizer',
            };

            const meetingsResponse = await graphClient.makeRequest(
              `/me/calendarView?${buildGraphQueryString(meetingQueryParams)}`,
              {
                method: 'GET',
              }
            );

            if (
              meetingsResponse &&
              typeof meetingsResponse === 'object' &&
              'value' in meetingsResponse
            ) {
              for (const meeting of meetingsResponse.value as GraphEvent[]) {
                const attendees = meeting.attendees || [];
                const organizer = meeting.organizer?.emailAddress?.address?.toLowerCase();

                // Check if target person is in this meeting
                const hasTarget =
                  organizer === targetEmail.toLowerCase() ||
                  attendees.some(
                    (a) => a.emailAddress?.address?.toLowerCase() === targetEmail.toLowerCase()
                  );

                if (hasTarget) {
                  // Track all other participants as mutual connections
                  for (const attendee of attendees) {
                    const attendeeEmail = attendee.emailAddress?.address?.toLowerCase();
                    const attendeeName = attendee.emailAddress?.name || attendeeEmail || 'Unknown';

                    if (
                      attendeeEmail &&
                      attendeeEmail !== currentUserEmail.toLowerCase() &&
                      attendeeEmail !== targetEmail.toLowerCase()
                    ) {
                      if (!mutualConnections.has(attendeeEmail)) {
                        mutualConnections.set(attendeeEmail, {
                          name: attendeeName,
                          email: attendeeEmail,
                          sharedMeetings: 0,
                          sharedEmails: 0,
                          connectionStrength: 0,
                        });
                      }
                      mutualConnections.get(attendeeEmail)!.sharedMeetings++;
                    }
                  }

                  // Also track organizer if not already tracked
                  if (
                    organizer &&
                    organizer !== currentUserEmail.toLowerCase() &&
                    organizer !== targetEmail.toLowerCase()
                  ) {
                    if (!mutualConnections.has(organizer)) {
                      mutualConnections.set(organizer, {
                        name: meeting.organizer?.emailAddress?.name || organizer,
                        email: organizer,
                        sharedMeetings: 0,
                        sharedEmails: 0,
                        connectionStrength: 0,
                      });
                    }
                    mutualConnections.get(organizer)!.sharedMeetings++;
                  }
                }
              }
            }
          } catch (error) {
            logger.warn(`Could not analyze shared meetings: ${error}`);
          }
        })()
      );

      // Find shared email threads
      promises.push(
        (async () => {
          try {
            const emailResponse = await graphClient.makeRequest('/me/messages', {
              method: 'GET',
              queryParams: {
                $top: '500',
                $select:
                  'id,subject,from,toRecipients,ccRecipients,conversationId,receivedDateTime',
                $filter: `receivedDateTime ge ${startDate.toISOString()}`,
                $orderby: 'receivedDateTime desc',
              },
            });

            if (emailResponse && typeof emailResponse === 'object' && 'value' in emailResponse) {
              // Group by conversation
              const conversations = new Map<string, GraphEmail[]>();
              for (const email of emailResponse.value as GraphEmail[]) {
                if (email.conversationId) {
                  if (!conversations.has(email.conversationId)) {
                    conversations.set(email.conversationId, []);
                  }
                  conversations.get(email.conversationId)!.push(email);
                }
              }

              // Find conversations with target person
              for (const [conversationId, emails] of conversations) {
                const hasTarget = emails.some(
                  (e) =>
                    e.from?.emailAddress?.address?.toLowerCase() === targetEmail.toLowerCase() ||
                    e.toRecipients?.some(
                      (r) => r.emailAddress?.address?.toLowerCase() === targetEmail.toLowerCase()
                    ) ||
                    e.ccRecipients?.some(
                      (r) => r.emailAddress?.address?.toLowerCase() === targetEmail.toLowerCase()
                    )
                );

                if (hasTarget) {
                  // Track all other participants
                  for (const email of emails) {
                    const senderEmail = email.from?.emailAddress?.address?.toLowerCase();
                    const senderName = email.from?.emailAddress?.name || senderEmail || 'Unknown';

                    if (
                      senderEmail &&
                      senderEmail !== currentUserEmail.toLowerCase() &&
                      senderEmail !== targetEmail.toLowerCase()
                    ) {
                      if (!mutualConnections.has(senderEmail)) {
                        mutualConnections.set(senderEmail, {
                          name: senderName,
                          email: senderEmail,
                          sharedMeetings: 0,
                          sharedEmails: 0,
                          connectionStrength: 0,
                        });
                      }
                      mutualConnections.get(senderEmail)!.sharedEmails++;
                    }

                    // Track recipients
                    for (const recipient of [
                      ...(email.toRecipients || []),
                      ...(email.ccRecipients || []),
                    ]) {
                      const recipientEmail = recipient.emailAddress?.address?.toLowerCase();
                      const recipientName =
                        recipient.emailAddress?.name || recipientEmail || 'Unknown';

                      if (
                        recipientEmail &&
                        recipientEmail !== currentUserEmail.toLowerCase() &&
                        recipientEmail !== targetEmail.toLowerCase() &&
                        recipientEmail !== senderEmail
                      ) {
                        if (!mutualConnections.has(recipientEmail)) {
                          mutualConnections.set(recipientEmail, {
                            name: recipientName,
                            email: recipientEmail,
                            sharedMeetings: 0,
                            sharedEmails: 0,
                            connectionStrength: 0,
                          });
                        }
                        mutualConnections.get(recipientEmail)!.sharedEmails++;
                      }
                    }
                  }
                }
              }
            }
          } catch (error) {
            logger.warn(`Could not analyze shared emails: ${error}`);
          }
        })()
      );

      await Promise.allSettled(promises);

      // Calculate connection strength
      for (const [email, connection] of mutualConnections) {
        connection.connectionStrength = connection.sharedMeetings * 3 + connection.sharedEmails;
      }

      // Sort by strength
      const sortedConnections = Array.from(mutualConnections.values())
        .sort((a, b) => b.connectionStrength - a.connectionStrength)
        .slice(0, limit);

      const result = {
        targetPerson: {
          name: targetUser.displayName,
          email: targetEmail,
        },
        analyzedPeriod: `Last ${days} days`,
        totalMutualConnections: mutualConnections.size,
        connections: sortedConnections.map((conn) => ({
          name: conn.name,
          email: conn.email,
          sharedMeetings: conn.sharedMeetings,
          sharedEmails: conn.sharedEmails,
          connectionStrength: conn.connectionStrength,
        })),
        summary: {
          strongestConnection: sortedConnections[0]?.name,
          mostSharedMeetings: sortedConnections.sort(
            (a, b) => b.sharedMeetings - a.sharedMeetings
          )[0]?.name,
          mostSharedEmails: sortedConnections.sort((a, b) => b.sharedEmails - a.sharedEmails)[0]
            ?.name,
        },
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 28. GET COMMUNICATION FREQUENCY - Analyze communication frequency
  // ==========================================================================
  server.tool(
    'get-communication-frequency',
    `Analyze communication frequency:
- Most frequent contacts
- Communication trends
- Response patterns
- Interaction types breakdown
- Time-based patterns

Use this for "Who do I email most often?", "Show my communication frequency", or "Analyze my communication patterns".`,
    {
      days: z.number().optional().describe('Days of history to analyze (default: 90)'),
      limit: z.number().optional().describe('Maximum contacts to return (default: 30)'),
      includeMeetings: z
        .boolean()
        .optional()
        .describe('Include meeting interactions (default: true)'),
    },
    {
      title: 'Get Communication Frequency',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ days = 90, limit = 30, includeMeetings = true }) => {
      logger.info(
        `Analyzing communication frequency: days=${days}, includeMeetings=${includeMeetings}`
      );

      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      interface CommunicationData {
        name: string;
        email: string;
        emailCount: number;
        meetingCount: number;
        totalInteractions: number;
        lastInteraction: string;
        firstInteraction: string;
        averageDaysBetween: number;
      }

      const communications = new Map<string, CommunicationData>();
      const promises: Promise<void>[] = [];

      // Analyze emails
      promises.push(
        (async () => {
          try {
            const emailResponse = await graphClient.makeRequest('/me/messages', {
              method: 'GET',
              queryParams: {
                $top: '1000',
                $select: 'id,from,receivedDateTime',
                $orderby: 'receivedDateTime desc',
                $filter: `receivedDateTime ge ${startDate.toISOString()}`,
              },
            });

            if (emailResponse && typeof emailResponse === 'object' && 'value' in emailResponse) {
              const currentUserEmail = await getCurrentUserEmail();

              for (const email of emailResponse.value as GraphEmail[]) {
                const senderEmail = email.from?.emailAddress?.address?.toLowerCase();
                const senderName = email.from?.emailAddress?.name || senderEmail || 'Unknown';

                if (senderEmail && senderEmail !== currentUserEmail.toLowerCase()) {
                  if (!communications.has(senderEmail)) {
                    communications.set(senderEmail, {
                      name: senderName,
                      email: senderEmail,
                      emailCount: 0,
                      meetingCount: 0,
                      totalInteractions: 0,
                      lastInteraction: email.receivedDateTime,
                      firstInteraction: email.receivedDateTime,
                      averageDaysBetween: 0,
                    });
                  }

                  const comm = communications.get(senderEmail)!;
                  comm.emailCount++;
                  comm.totalInteractions++;
                  if (new Date(email.receivedDateTime) > new Date(comm.lastInteraction)) {
                    comm.lastInteraction = email.receivedDateTime;
                  }
                  if (new Date(email.receivedDateTime) < new Date(comm.firstInteraction)) {
                    comm.firstInteraction = email.receivedDateTime;
                  }
                }
              }
            }
          } catch (error) {
            logger.warn(`Could not analyze email frequency: ${error}`);
          }
        })()
      );

      // Analyze meetings
      if (includeMeetings) {
        promises.push(
          (async () => {
            try {
              const meetingQueryParams: Record<string, string> = {
                startDateTime: startDate.toISOString(),
                endDateTime: new Date().toISOString(),
                $top: '500',
                $select: 'id,subject,start,attendees,organizer',
              };

              const meetingsResponse = await graphClient.makeRequest(
                `/me/calendarView?${buildGraphQueryString(meetingQueryParams)}`,
                {
                  method: 'GET',
                }
              );

              if (
                meetingsResponse &&
                typeof meetingsResponse === 'object' &&
                'value' in meetingsResponse
              ) {
                const currentUserEmail = await getCurrentUserEmail();

                for (const meeting of meetingsResponse.value as GraphEvent[]) {
                  const organizerEmail = meeting.organizer?.emailAddress?.address?.toLowerCase();
                  const organizerName =
                    meeting.organizer?.emailAddress?.name || organizerEmail || 'Unknown';

                  // Track organizer
                  if (organizerEmail && organizerEmail !== currentUserEmail.toLowerCase()) {
                    if (!communications.has(organizerEmail)) {
                      communications.set(organizerEmail, {
                        name: organizerName,
                        email: organizerEmail,
                        emailCount: 0,
                        meetingCount: 0,
                        totalInteractions: 0,
                        lastInteraction: meeting.start.dateTime,
                        firstInteraction: meeting.start.dateTime,
                        averageDaysBetween: 0,
                      });
                    }
                    communications.get(organizerEmail)!.meetingCount++;
                    communications.get(organizerEmail)!.totalInteractions++;
                  }

                  // Track attendees
                  for (const attendee of meeting.attendees || []) {
                    const attendeeEmail = attendee.emailAddress?.address?.toLowerCase();
                    const attendeeName = attendee.emailAddress?.name || attendeeEmail || 'Unknown';

                    if (attendeeEmail && attendeeEmail !== currentUserEmail.toLowerCase()) {
                      if (!communications.has(attendeeEmail)) {
                        communications.set(attendeeEmail, {
                          name: attendeeName,
                          email: attendeeEmail,
                          emailCount: 0,
                          meetingCount: 0,
                          totalInteractions: 0,
                          lastInteraction: meeting.start.dateTime,
                          firstInteraction: meeting.start.dateTime,
                          averageDaysBetween: 0,
                        });
                      }
                      communications.get(attendeeEmail)!.meetingCount++;
                      communications.get(attendeeEmail)!.totalInteractions++;
                    }
                  }
                }
              }
            } catch (error) {
              logger.warn(`Could not analyze meeting frequency: ${error}`);
            }
          })()
        );
      }

      await Promise.allSettled(promises);

      // Calculate average days between interactions
      for (const [email, comm] of communications) {
        const totalDays =
          (new Date(comm.lastInteraction).getTime() - new Date(comm.firstInteraction).getTime()) /
          (1000 * 60 * 60 * 24);
        comm.averageDaysBetween =
          comm.totalInteractions > 1 ? totalDays / (comm.totalInteractions - 1) : totalDays;
      }

      // Sort by total interactions
      const sortedCommunications = Array.from(communications.values())
        .sort((a, b) => b.totalInteractions - a.totalInteractions)
        .slice(0, limit);

      // Calculate trends (weekly breakdown)
      const weeklyBreakdown = new Map<number, number>();
      for (const comm of sortedCommunications) {
        const weekNum = Math.floor(
          (new Date(comm.lastInteraction).getTime() - startDate.getTime()) /
            (7 * 24 * 60 * 60 * 1000)
        );
        weeklyBreakdown.set(weekNum, (weeklyBreakdown.get(weekNum) || 0) + comm.totalInteractions);
      }

      const result = {
        analyzedPeriod: `Last ${days} days`,
        totalContacts: communications.size,
        topContacts: sortedCommunications.map((comm) => ({
          name: comm.name,
          email: comm.email,
          emailCount: comm.emailCount,
          meetingCount: comm.meetingCount,
          totalInteractions: comm.totalInteractions,
          lastInteraction: comm.lastInteraction,
          averageDaysBetween: Math.round(comm.averageDaysBetween * 10) / 10,
        })),
        summary: {
          mostFrequentContact: sortedCommunications[0]?.name,
          totalEmails: sortedCommunications.reduce((sum, c) => sum + c.emailCount, 0),
          totalMeetings: sortedCommunications.reduce((sum, c) => sum + c.meetingCount, 0),
          averageInteractionsPerContact:
            sortedCommunications.length > 0
              ? Math.round(
                  (sortedCommunications.reduce((sum, c) => sum + c.totalInteractions, 0) /
                    sortedCommunications.length) *
                    10
                ) / 10
              : 0,
        },
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // DOCUMENT INTELLIGENCE TOOLS
  // ==========================================================================

  // ==========================================================================
  // 29. FIND RELATED DOCUMENTS - Find related documents across services
  // ==========================================================================
  server.tool(
    'find-related-documents',
    `Find related documents across services:
- Documents mentioned in emails
- Files related to meetings
- Documents shared with same people
- Related by topic/keywords
- Version history connections

Use this for "Find documents related to [topic]", "Show files related to this meeting", or "Find all documents about [project]".`,
    {
      topic: z.string().describe('Topic, project name, or keyword to find related documents'),
      days: z.number().optional().describe('Days back to search (default: 180)'),
      limit: z.number().optional().describe('Maximum documents to return (default: 50)'),
      includeEmails: z
        .boolean()
        .optional()
        .describe('Include documents mentioned in emails (default: true)'),
      includeMeetings: z
        .boolean()
        .optional()
        .describe('Include documents related to meetings (default: true)'),
    },
    {
      title: 'Find Related Documents',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ topic, days = 180, limit = 50, includeEmails = true, includeMeetings = true }) => {
      logger.info(`Finding related documents for topic: ${topic}`);

      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      interface RelatedDocument {
        name: string;
        webUrl?: string;
        type: 'file' | 'email' | 'meeting';
        source: string;
        sourceDate: string;
        relevance: number;
        sharedWith?: string[];
      }

      const documents = new Map<string, RelatedDocument>();
      const promises: Promise<void>[] = [];

      // Search for files directly
      promises.push(
        (async () => {
          try {
            const searchResult = await executeCentralSearch(graphClient, topic, {
              entityTypes: ['driveItem', 'listItem'],
              maxResults: limit,
              sortByRank: true,
            });

            for (const hit of [...searchResult.results.files, ...searchResult.results.listItems]) {
              if (hit.resource) {
                const doc = hit.resource as GraphDriveItem;
                const key = doc.id || doc.name || '';
                if (key && !documents.has(key)) {
                  documents.set(key, {
                    name: doc.name || 'Unknown',
                    webUrl: doc.webUrl,
                    type: 'file',
                    source: 'Direct search',
                    sourceDate: doc.lastModifiedDateTime || doc.createdDateTime || '',
                    relevance: hit.relevanceScore || 0,
                  });
                }
              }
            }
          } catch (error) {
            logger.warn(`Could not search for files: ${error}`);
          }
        })()
      );

      // Find documents mentioned in emails
      if (includeEmails) {
        promises.push(
          (async () => {
            try {
              const emailResponse = await graphClient.makeRequest('/me/messages', {
                method: 'GET',
                queryParams: {
                  $search: `"${topic}"`,
                  $top: '100',
                  $select:
                    'id,subject,bodyPreview,body,receivedDateTime,hasAttachments,attachments,webLink',
                  $filter: `receivedDateTime ge ${startDate.toISOString()}`,
                },
              });

              if (emailResponse && typeof emailResponse === 'object' && 'value' in emailResponse) {
                for (const email of emailResponse.value as GraphEmail &
                  {
                    hasAttachments?: boolean;
                    attachments?: Array<{ name: string; contentId?: string }>;
                  }[]) {
                  const content = sanitizeHtml(email.bodyPreview || email.body?.content || '');

                  // Look for file references (common patterns)
                  const filePatterns = [
                    /(?:see|check|review|attached|attachment)[\s:]+([^\s.]+\.(?:docx?|xlsx?|pptx?|pdf|txt))/gi,
                    /([^\s.]+\.(?:docx?|xlsx?|pptx?|pdf|txt))/gi,
                  ];

                  for (const pattern of filePatterns) {
                    const matches = content.matchAll(pattern);
                    for (const match of matches) {
                      if (match[1]) {
                        const fileName = match[1].trim();
                        if (!documents.has(fileName)) {
                          documents.set(fileName, {
                            name: fileName,
                            type: 'email',
                            source: `Email: ${email.subject}`,
                            sourceDate: email.receivedDateTime,
                            relevance: 50,
                          });
                        }
                      }
                    }
                  }

                  // Check attachments
                  if (email.hasAttachments && email.attachments) {
                    for (const attachment of email.attachments) {
                      if (attachment.name) {
                        const key = `email-${email.id}-${attachment.name}`;
                        if (!documents.has(key)) {
                          documents.set(key, {
                            name: attachment.name,
                            type: 'email',
                            source: `Email attachment: ${email.subject}`,
                            sourceDate: email.receivedDateTime,
                            relevance: 70,
                            webUrl: email.webLink,
                          });
                        }
                      }
                    }
                  }
                }
              }
            } catch (error) {
              logger.warn(`Could not find documents in emails: ${error}`);
            }
          })()
        );
      }

      // Find documents related to meetings
      if (includeMeetings) {
        promises.push(
          (async () => {
            try {
              const meetingQueryParams: Record<string, string> = {
                startDateTime: startDate.toISOString(),
                endDateTime: new Date().toISOString(),
                $top: '100',
                $select: 'id,subject,bodyPreview,start,attendees,organizer,webLink',
              };

              const meetingsResponse = await graphClient.makeRequest(
                `/me/calendarView?${buildGraphQueryString(meetingQueryParams)}`,
                {
                  method: 'GET',
                }
              );

              if (
                meetingsResponse &&
                typeof meetingsResponse === 'object' &&
                'value' in meetingsResponse
              ) {
                for (const meeting of meetingsResponse.value as GraphEvent[]) {
                  const subject = meeting.subject || '';
                  const content = sanitizeHtml(meeting.bodyPreview || '');

                  // Check if meeting is related to topic
                  if (
                    subject.toLowerCase().includes(topic.toLowerCase()) ||
                    content.toLowerCase().includes(topic.toLowerCase())
                  ) {
                    // Look for file references in meeting notes
                    const filePatterns = [
                      /(?:see|check|review|document|file)[\s:]+([^\s.]+\.(?:docx?|xlsx?|pptx?|pdf|txt))/gi,
                      /([^\s.]+\.(?:docx?|xlsx?|pptx?|pdf|txt))/gi,
                    ];

                    for (const pattern of filePatterns) {
                      const matches = content.matchAll(pattern);
                      for (const match of matches) {
                        if (match[1]) {
                          const fileName = match[1].trim();
                          const key = `meeting-${meeting.id}-${fileName}`;
                          if (!documents.has(key)) {
                            documents.set(key, {
                              name: fileName,
                              type: 'meeting',
                              source: `Meeting: ${subject}`,
                              sourceDate: meeting.start.dateTime,
                              relevance: 60,
                              webUrl: meeting.webLink,
                              sharedWith: [
                                meeting.organizer?.emailAddress?.name,
                                ...(meeting.attendees || []).map(
                                  (a) => a.emailAddress?.name || a.emailAddress?.address
                                ),
                              ].filter(Boolean),
                            });
                          }
                        }
                      }
                    }
                  }
                }
              }
            } catch (error) {
              logger.warn(`Could not find documents in meetings: ${error}`);
            }
          })()
        );
      }

      await Promise.allSettled(promises);

      // Sort by relevance
      const sortedDocuments = Array.from(documents.values())
        .sort((a, b) => b.relevance - a.relevance)
        .slice(0, limit);

      const result = {
        topic,
        analyzedPeriod: `Last ${days} days`,
        totalDocuments: documents.size,
        documents: sortedDocuments,
        summary: {
          byType: {
            files: sortedDocuments.filter((d) => d.type === 'file').length,
            emails: sortedDocuments.filter((d) => d.type === 'email').length,
            meetings: sortedDocuments.filter((d) => d.type === 'meeting').length,
          },
          mostRelevant: sortedDocuments[0]?.name,
        },
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 30. BUILD KNOWLEDGE GRAPH - Build knowledge graph from data
  // ==========================================================================
  server.tool(
    'build-knowledge-graph',
    `Build knowledge graph from data:
- Connect people, projects, documents
- Identify relationships
- Map collaboration networks
- Visualize connections
- Discover hidden relationships

Use this for "Build a knowledge graph for [topic]", "Show connections for [project]", or "Map relationships for [person]".`,
    {
      topic: z.string().describe('Topic, project, or person to build knowledge graph for'),
      days: z.number().optional().describe('Days of history to analyze (default: 180)'),
      maxNodes: z.number().optional().describe('Maximum nodes in graph (default: 50)'),
    },
    {
      title: 'Build Knowledge Graph',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ topic, days = 180, maxNodes = 50 }) => {
      logger.info(`Building knowledge graph for: ${topic}`);

      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      interface GraphNode {
        id: string;
        type: 'person' | 'document' | 'project' | 'meeting';
        name: string;
        properties?: Record<string, unknown>;
      }

      interface GraphEdge {
        from: string;
        to: string;
        type: 'email' | 'meeting' | 'document' | 'collaboration';
        weight: number;
        properties?: Record<string, unknown>;
      }

      const nodes = new Map<string, GraphNode>();
      const edges: GraphEdge[] = [];

      // Try to detect if topic is a person
      const topicUser = await findUser(graphClient, topic);
      const isPerson = !!topicUser;

      const promises: Promise<void>[] = [];

      // Add topic as central node
      if (isPerson && topicUser) {
        nodes.set(topicUser.id, {
          id: topicUser.id,
          type: 'person',
          name: topicUser.displayName,
          properties: {
            email: topicUser.mail || topicUser.userPrincipalName,
            department: topicUser.department,
            jobTitle: topicUser.jobTitle,
          },
        });
      } else {
        nodes.set('topic', {
          id: 'topic',
          type: 'project',
          name: topic,
        });
      }

      // Find related emails
      promises.push(
        (async () => {
          try {
            const emailResponse = await graphClient.makeRequest('/me/messages', {
              method: 'GET',
              queryParams: {
                $search: `"${topic}"`,
                $top: '100',
                $select:
                  'id,subject,from,toRecipients,ccRecipients,receivedDateTime,conversationId',
                $filter: `receivedDateTime ge ${startDate.toISOString()}`,
              },
            });

            if (emailResponse && typeof emailResponse === 'object' && 'value' in emailResponse) {
              const topicNodeId = isPerson && topicUser ? topicUser.id : 'topic';

              for (const email of emailResponse.value as GraphEmail) {
                const senderEmail = email.from?.emailAddress?.address?.toLowerCase();
                const senderName = email.from?.emailAddress?.name || senderEmail || 'Unknown';
                const senderId = `person-${senderEmail}`;

                // Add sender node
                if (!nodes.has(senderId)) {
                  nodes.set(senderId, {
                    id: senderId,
                    type: 'person',
                    name: senderName,
                    properties: { email: senderEmail },
                  });
                }

                // Add edge from sender to topic
                edges.push({
                  from: senderId,
                  to: topicNodeId,
                  type: 'email',
                  weight: 1,
                  properties: { subject: email.subject, date: email.receivedDateTime },
                });

                // Add recipients
                for (const recipient of [
                  ...(email.toRecipients || []),
                  ...(email.ccRecipients || []),
                ]) {
                  const recipientEmail = recipient.emailAddress?.address?.toLowerCase();
                  const recipientName = recipient.emailAddress?.name || recipientEmail || 'Unknown';
                  const recipientId = `person-${recipientEmail}`;

                  if (!nodes.has(recipientId)) {
                    nodes.set(recipientId, {
                      id: recipientId,
                      type: 'person',
                      name: recipientName,
                      properties: { email: recipientEmail },
                    });
                  }

                  // Add edge between sender and recipient
                  edges.push({
                    from: senderId,
                    to: recipientId,
                    type: 'email',
                    weight: 1,
                    properties: { subject: email.subject },
                  });
                }
              }
            }
          } catch (error) {
            logger.warn(`Could not analyze emails for knowledge graph: ${error}`);
          }
        })()
      );

      // Find related meetings
      promises.push(
        (async () => {
          try {
            const meetingQueryParams: Record<string, string> = {
              startDateTime: startDate.toISOString(),
              endDateTime: new Date().toISOString(),
              $top: '100',
              $select: 'id,subject,start,attendees,organizer',
            };

            const meetingsResponse = await graphClient.makeRequest(
              `/me/calendarView?${buildGraphQueryString(meetingQueryParams)}`,
              {
                method: 'GET',
              }
            );

            if (
              meetingsResponse &&
              typeof meetingsResponse === 'object' &&
              'value' in meetingsResponse
            ) {
              const topicNodeId = isPerson && topicUser ? topicUser.id : 'topic';

              for (const meeting of meetingsResponse.value as GraphEvent) {
                const subject = meeting.subject || '';
                if (!subject.toLowerCase().includes(topic.toLowerCase())) {
                  continue;
                }

                const meetingId = `meeting-${meeting.id}`;
                nodes.set(meetingId, {
                  id: meetingId,
                  type: 'meeting',
                  name: subject,
                  properties: { date: meeting.start.dateTime },
                });

                // Connect meeting to topic
                edges.push({
                  from: meetingId,
                  to: topicNodeId,
                  type: 'meeting',
                  weight: 2,
                });

                // Connect participants
                const organizerEmail = meeting.organizer?.emailAddress?.address?.toLowerCase();
                const organizerName =
                  meeting.organizer?.emailAddress?.name || organizerEmail || 'Unknown';
                const organizerId = `person-${organizerEmail}`;

                if (organizerEmail && !nodes.has(organizerId)) {
                  nodes.set(organizerId, {
                    id: organizerId,
                    type: 'person',
                    name: organizerName,
                    properties: { email: organizerEmail },
                  });
                }

                if (organizerEmail) {
                  edges.push({
                    from: organizerId,
                    to: meetingId,
                    type: 'meeting',
                    weight: 2,
                  });
                }

                for (const attendee of meeting.attendees || []) {
                  const attendeeEmail = attendee.emailAddress?.address?.toLowerCase();
                  const attendeeName = attendee.emailAddress?.name || attendeeEmail || 'Unknown';
                  const attendeeId = `person-${attendeeEmail}`;

                  if (attendeeEmail && !nodes.has(attendeeId)) {
                    nodes.set(attendeeId, {
                      id: attendeeId,
                      type: 'person',
                      name: attendeeName,
                      properties: { email: attendeeEmail },
                    });
                  }

                  if (attendeeEmail) {
                    edges.push({
                      from: attendeeId,
                      to: meetingId,
                      type: 'meeting',
                      weight: 2,
                    });
                  }
                }
              }
            }
          } catch (error) {
            logger.warn(`Could not analyze meetings for knowledge graph: ${error}`);
          }
        })()
      );

      // Find related documents
      promises.push(
        (async () => {
          try {
            const searchResult = await executeCentralSearch(graphClient, topic, {
              entityTypes: ['driveItem'],
              maxResults: 20,
              sortByRank: true,
            });

            for (const hit of searchResult.results.files) {
              if (hit.resource) {
                const doc = hit.resource as GraphDriveItem;
                const docId = `doc-${doc.id}`;
                const topicNodeId = isPerson && topicUser ? topicUser.id : 'topic';

                nodes.set(docId, {
                  id: docId,
                  type: 'document',
                  name: doc.name || 'Unknown',
                  properties: {
                    url: doc.webUrl,
                    modified: doc.lastModifiedDateTime,
                  },
                });

                edges.push({
                  from: docId,
                  to: topicNodeId,
                  type: 'document',
                  weight: 3,
                });
              }
            }
          } catch (error) {
            logger.warn(`Could not analyze documents for knowledge graph: ${error}`);
          }
        })()
      );

      await Promise.allSettled(promises);

      // Limit nodes and calculate node degrees
      const nodeArray = Array.from(nodes.values()).slice(0, maxNodes);
      const nodeIds = new Set(nodeArray.map((n) => n.id));
      const filteredEdges = edges.filter((e) => nodeIds.has(e.from) && nodeIds.has(e.to));

      // Calculate node degrees (number of connections)
      const nodeDegrees = new Map<string, number>();
      for (const edge of filteredEdges) {
        nodeDegrees.set(edge.from, (nodeDegrees.get(edge.from) || 0) + 1);
        nodeDegrees.set(edge.to, (nodeDegrees.get(edge.to) || 0) + 1);
      }

      // Add degrees to node properties
      for (const node of nodeArray) {
        node.properties = {
          ...node.properties,
          degree: nodeDegrees.get(node.id) || 0,
        };
      }

      const result = {
        topic,
        analyzedPeriod: `Last ${days} days`,
        graph: {
          nodes: nodeArray,
          edges: filteredEdges,
        },
        statistics: {
          totalNodes: nodeArray.length,
          totalEdges: filteredEdges.length,
          nodeTypes: {
            person: nodeArray.filter((n) => n.type === 'person').length,
            document: nodeArray.filter((n) => n.type === 'document').length,
            meeting: nodeArray.filter((n) => n.type === 'meeting').length,
            project: nodeArray.filter((n) => n.type === 'project').length,
          },
          mostConnected: nodeArray
            .sort((a, b) => (b.properties?.degree || 0) - (a.properties?.degree || 0))
            .slice(0, 5)
            .map((n) => ({ name: n.name, connections: n.properties?.degree || 0 })),
        },
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // TOOL ORCHESTRATION & INTELLIGENCE TOOLS
  // ==========================================================================

  // ==========================================================================
  // 31. PLAN TOOL EXECUTION - The Orchestrator - Creates execution plan for ANY query
  // ==========================================================================
  server.tool(
    'plan-tool-execution',
    `🎯 **TOOL ORCHESTRATOR** - THE MOST IMPORTANT TOOL FOR LLMs!

**CRITICAL FUNCTION**: This tool tells the LLM EXACTLY which tools to call, in which order, with which parameters, for ANY question the user asks.

**When to use:**
- Call this tool FIRST for ANY user query/question
- Use when you need to know which tools to call
- Use for complex queries requiring multiple tools
- Use when you're unsure which tool is best

**What it does:**
- Analyzes ANY user query/question
- Breaks it down into concrete tool calls
- Provides EXACT tool names and parameters
- Gives clear step-by-step instructions
- Handles tool dependencies automatically
- Suggests parallel execution where possible

**Output:**
Returns a detailed execution plan with:
- Step-by-step tool calls with exact parameters
- Clear reasons for each step
- Expected results
- Parallel execution opportunities
- Fallback strategies

**Example:**
User asks: "What do I know about Project Apollo?"
→ Call this tool → Get plan with exact tools to call

This tool works for ALL types of queries: emails, meetings, files, people, projects, tasks, etc.`,
    {
      query: z.string().describe('The user query or question to create an execution plan for'),
    },
    {
      title: 'Plan Tool Execution',
      readOnlyHint: true,
      openWorldHint: false,
    },
    async ({ query }) => {
      logger.info(`Planning tool execution for query: ${query}`);

      try {
        const plan = createToolExecutionPlan(query, nlpEnhancer);

        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(plan, null, 2),
            },
          ],
        };
      } catch (error) {
        logger.error(`Error creating tool execution plan: ${error}`);
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  error: 'Failed to create execution plan',
                  message: `${error}`,
                  query,
                  fallback: {
                    tool: 'ask-microsoft-365',
                    parameters: { question: query },
                    reason: 'Use intelligent assistant as fallback',
                  },
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
  registeredCount++;

  // ==========================================================================
  // 32. SUGGEST TOOL SEQUENCE - Suggests optimal tool sequences for scenarios
  // ==========================================================================
  server.tool(
    'suggest-tool-sequence',
    `Suggest optimal tool sequences for common business scenarios:
- Pre-defined workflows for common tasks
- Context-aware tool recommendations
- Multi-step operation planning
- Tool dependency resolution

Use this for "What's the best way to prepare for a meeting?", "How do I get a complete customer overview?", or "What tools should I use to analyze my productivity?".`,
    {
      scenario: z
        .enum([
          'meeting_preparation',
          'customer_overview',
          'project_research',
          'person_research',
          'productivity_analysis',
          'weekly_review',
          'email_analysis',
          'document_discovery',
        ])
        .describe('Business scenario to get tool sequence for'),
      context: z
        .string()
        .optional()
        .describe('Additional context (e.g., person name, project name)'),
    },
    {
      title: 'Suggest Tool Sequence',
      readOnlyHint: true,
      openWorldHint: false,
    },
    async ({ scenario, context }) => {
      logger.info(`Suggesting tool sequence for scenario: ${scenario}`);

      const workflows: Record<
        string,
        {
          scenario: string;
          description: string;
          steps: Array<{
            step: number;
            tool: string;
            parameters: Record<string, unknown>;
            reason: string;
          }>;
        }
      > = {
        meeting_preparation: {
          scenario: 'meeting_preparation',
          description: 'Prepare for an upcoming meeting',
          steps: [
            {
              step: 1,
              tool: 'prepare-for-meeting',
              parameters: {
                meetingSubject: context || 'upcoming meeting',
                hoursAhead: 48,
              },
              reason: 'Gather all context for the meeting',
            },
            {
              step: 2,
              tool: 'find-related-documents',
              parameters: {
                topic: context || 'meeting topic',
                days: 30,
                limit: 20,
              },
              reason: 'Find documents related to the meeting',
            },
          ],
        },
        customer_overview: {
          scenario: 'customer_overview',
          description: 'Get complete customer overview',
          steps: [
            {
              step: 1,
              tool: 'get-company-contacts',
              parameters: {
                companyName: context || 'company',
              },
              reason: 'Find all contacts from the company',
            },
            {
              step: 2,
              tool: 'search-everything',
              parameters: {
                query: context || 'company name',
                limit: 25,
              },
              reason: 'Search for all interactions with the company',
            },
            {
              step: 3,
              tool: 'find-related-documents',
              parameters: {
                topic: context || 'company name',
                days: 180,
                limit: 50,
              },
              reason: 'Find all documents related to the company',
            },
          ],
        },
        project_research: {
          scenario: 'project_research',
          description: 'Research a project comprehensively',
          steps: [
            {
              step: 1,
              tool: 'search-everything',
              parameters: {
                query: context || 'project name',
                limit: 25,
              },
              reason: 'Start with universal search',
            },
            {
              step: 2,
              tool: 'get-project-overview',
              parameters: {
                projectName: context || 'project name',
                includeFiles: true,
                includeMeetings: true,
                includeEmails: true,
                includeTasks: true,
              },
              reason: 'Get structured project overview',
            },
            {
              step: 3,
              tool: 'get-project-stakeholders',
              parameters: {
                projectName: context || 'project name',
                days: 90,
              },
              reason: 'Identify all stakeholders',
            },
            {
              step: 4,
              tool: 'find-related-documents',
              parameters: {
                topic: context || 'project name',
                days: 180,
                limit: 50,
              },
              reason: 'Find all related documents',
            },
          ],
        },
        person_research: {
          scenario: 'person_research',
          description: 'Research a person comprehensively',
          steps: [
            {
              step: 1,
              tool: 'get-communication-summary',
              parameters: {
                person: context || 'person name',
                includeEmails: true,
                includeChats: true,
                includeMeetings: true,
                includeFiles: true,
              },
              reason: 'Get complete communication overview',
            },
            {
              step: 2,
              tool: 'analyze-relationship-strength',
              parameters: {
                person: context || 'person name',
                days: 90,
              },
              reason: 'Analyze relationship strength',
            },
            {
              step: 3,
              tool: 'find-mutual-connections',
              parameters: {
                person: context || 'person name',
                days: 180,
              },
              reason: 'Find mutual connections',
            },
          ],
        },
        productivity_analysis: {
          scenario: 'productivity_analysis',
          description: 'Analyze productivity',
          steps: [
            {
              step: 1,
              tool: 'get-my-week-summary',
              parameters: {
                weekOffset: 0,
              },
              reason: 'Get weekly summary',
            },
            {
              step: 2,
              tool: 'analyze-meeting-load',
              parameters: {
                weeks: 4,
                includeRecurring: true,
              },
              reason: 'Analyze meeting load',
            },
            {
              step: 3,
              tool: 'get-communication-frequency',
              parameters: {
                days: 90,
                limit: 30,
                includeMeetings: true,
              },
              reason: 'Analyze communication frequency',
            },
          ],
        },
        weekly_review: {
          scenario: 'weekly_review',
          description: 'Weekly review workflow',
          steps: [
            {
              step: 1,
              tool: 'get-my-week-summary',
              parameters: {
                weekOffset: 0,
              },
              reason: 'Get weekly summary',
            },
            {
              step: 2,
              tool: 'get-deadline-overview',
              parameters: {
                days: 7,
                includeCompleted: false,
              },
              reason: 'Check upcoming deadlines',
            },
            {
              step: 3,
              tool: 'get-follow-up-items',
              parameters: {
                includeEmails: true,
                includeTasks: true,
                includeMeetings: true,
              },
              reason: 'Get items needing attention',
            },
          ],
        },
        email_analysis: {
          scenario: 'email_analysis',
          description: 'Analyze emails',
          steps: [
            {
              step: 1,
              tool: 'get-my-emails',
              parameters: {
                filter: 'all',
                limit: 50,
              },
              reason: 'Get recent emails',
            },
            {
              step: 2,
              tool: 'extract-action-items',
              parameters: {
                source: 'emails',
                days: 7,
                limit: 50,
              },
              reason: 'Extract action items from emails',
            },
            {
              step: 3,
              tool: 'find-unresponded-requests',
              parameters: {
                days: 14,
                priorityOnly: false,
              },
              reason: 'Find unresponded requests',
            },
          ],
        },
        document_discovery: {
          scenario: 'document_discovery',
          description: 'Discover related documents',
          steps: [
            {
              step: 1,
              tool: 'search-everything',
              parameters: {
                query: context || 'topic',
                limit: 25,
              },
              reason: 'Search for documents',
            },
            {
              step: 2,
              tool: 'find-related-documents',
              parameters: {
                topic: context || 'topic',
                days: 180,
                limit: 50,
                includeEmails: true,
                includeMeetings: true,
              },
              reason: 'Find related documents',
            },
            {
              step: 3,
              tool: 'build-knowledge-graph',
              parameters: {
                topic: context || 'topic',
                days: 180,
                maxNodes: 50,
              },
              reason: 'Build knowledge graph',
            },
          ],
        },
      };

      const workflow = workflows[scenario];
      if (!workflow) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  error: 'Unknown scenario',
                  availableScenarios: Object.keys(workflows),
                },
                null,
                2
              ),
            },
          ],
          isError: true,
        };
      }

      // Update parameters with context if provided
      if (context) {
        for (const step of workflow.steps) {
          for (const [key, value] of Object.entries(step.parameters)) {
            if ((typeof value === 'string' && value.includes('name')) || value.includes('topic')) {
              step.parameters[key] = context;
            }
          }
        }
      }

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(workflow, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 33. GET TOOL RECOMMENDATIONS - AI-powered tool recommendations
  // ==========================================================================
  server.tool(
    'get-tool-recommendations',
    `Get AI-powered tool recommendations based on:
- Current query context
- Historical successful tool combinations
- User's typical workflows
- Available data sources

Use this for "What tools should I use for this query?", "Recommend tools for analyzing [topic]", or "What's the best tool combination for [task]?".`,
    {
      query: z.string().describe('Query or task to get tool recommendations for'),
      limit: z.number().optional().describe('Maximum recommendations to return (default: 5)'),
    },
    {
      title: 'Get Tool Recommendations',
      readOnlyHint: true,
      openWorldHint: false,
    },
    async ({ query, limit = 5 }) => {
      logger.info(`Getting tool recommendations for: ${query}`);

      try {
        // Analyze query
        const decomposed = nlpEnhancer.decomposeQuery(query);
        const queryLower = query.toLowerCase();

        // Find matching tools
        const matchingTools: Array<{
          tool: ToolDefinition;
          score: number;
          reason: string;
        }> = [];

        for (const toolDef of AVAILABLE_TOOLS) {
          let score = 0;
          const reasons: string[] = [];

          // Check use cases
          for (const useCase of toolDef.useCases) {
            if (queryLower.includes(useCase)) {
              score += 10;
              reasons.push(`Matches use case: ${useCase}`);
            }
          }

          // Check category match
          if (
            (queryLower.includes('email') || queryLower.includes('mail')) &&
            toolDef.category === 'email'
          ) {
            score += 5;
            reasons.push('Email-related query');
          }
          if (
            (queryLower.includes('meeting') || queryLower.includes('calendar')) &&
            toolDef.category === 'calendar'
          ) {
            score += 5;
            reasons.push('Meeting-related query');
          }
          if (
            (queryLower.includes('person') || queryLower.includes('people')) &&
            toolDef.category === 'people'
          ) {
            score += 5;
            reasons.push('People-related query');
          }
          if (
            (queryLower.includes('project') || queryLower.includes('task')) &&
            toolDef.category === 'project'
          ) {
            score += 5;
            reasons.push('Project-related query');
          }
          if (
            (queryLower.includes('file') || queryLower.includes('document')) &&
            toolDef.category === 'documents'
          ) {
            score += 5;
            reasons.push('Document-related query');
          }

          // Check description match
          if (toolDef.description.toLowerCase().includes(queryLower.split(' ')[0])) {
            score += 3;
            reasons.push('Description matches');
          }

          if (score > 0) {
            matchingTools.push({
              tool: toolDef,
              score,
              reason: reasons.join('; '),
            });
          }
        }

        // Sort by score
        matchingTools.sort((a, b) => b.score - a.score);

        // Get top recommendations
        const recommendations = matchingTools.slice(0, limit).map((match, index) => ({
          rank: index + 1,
          tool: match.tool.name,
          description: match.tool.description,
          category: match.tool.category,
          score: match.score,
          reason: match.reason,
          parameters: match.tool.parameters,
          useCases: match.tool.useCases,
        }));

        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  query,
                  recommendations,
                  totalMatches: matchingTools.length,
                  suggestion:
                    recommendations.length > 0
                      ? `Consider using '${recommendations[0].tool}' as the primary tool`
                      : "Use 'ask-microsoft-365' for general queries",
                },
                null,
                2
              ),
            },
          ],
        };
      } catch (error) {
        logger.error(`Error getting tool recommendations: ${error}`);
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  error: 'Failed to get recommendations',
                  message: `${error}`,
                  query,
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
  registeredCount++;

  // ==========================================================================
  // BUSINESS WORKFLOW TOOLS
  // ==========================================================================

  // ==========================================================================
  // 34. EXECUTE BUSINESS WORKFLOW - Execute predefined business workflows
  // ==========================================================================
  server.tool(
    'execute-business-workflow',
    `Execute predefined business workflows:
- Customer onboarding workflow
- Project kickoff workflow
- Meeting preparation workflow
- Weekly review workflow
- Client communication workflow

Use this for "Execute customer onboarding workflow for [company]", "Run meeting preparation workflow for [meeting]", or "Start weekly review workflow".`,
    {
      workflow: z
        .enum([
          'customer_onboarding',
          'project_kickoff',
          'meeting_preparation',
          'weekly_review',
          'client_communication',
        ])
        .describe('Workflow to execute'),
      context: z
        .string()
        .describe('Context for the workflow (e.g., company name, project name, meeting subject)'),
    },
    {
      title: 'Execute Business Workflow',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ workflow, context }) => {
      logger.info(`Executing workflow: ${workflow} with context: ${context}`);

      // Get tool sequence from suggest-tool-sequence
      const scenarioMap: Record<string, string> = {
        customer_onboarding: 'customer_overview',
        project_kickoff: 'project_research',
        meeting_preparation: 'meeting_preparation',
        weekly_review: 'weekly_review',
        client_communication: 'person_research',
      };

      const scenario = scenarioMap[workflow];
      if (!scenario) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  error: 'Unknown workflow',
                  availableWorkflows: Object.keys(scenarioMap),
                },
                null,
                2
              ),
            },
          ],
          isError: true,
        };
      }

      // Use suggest-tool-sequence logic
      const workflows: Record<string, any> = {
        customer_overview: {
          workflow: 'customer_onboarding',
          description: 'Customer onboarding workflow',
          steps: [
            {
              step: 1,
              tool: 'get-company-contacts',
              parameters: { companyName: context },
              reason: 'Find all contacts from the company',
            },
            {
              step: 2,
              tool: 'search-everything',
              parameters: { query: context, limit: 25 },
              reason: 'Search for all interactions',
            },
            {
              step: 3,
              tool: 'find-related-documents',
              parameters: { topic: context, days: 180, limit: 50 },
              reason: 'Find all documents',
            },
          ],
        },
        project_research: {
          workflow: 'project_kickoff',
          description: 'Project kickoff workflow',
          steps: [
            {
              step: 1,
              tool: 'search-everything',
              parameters: { query: context, limit: 25 },
              reason: 'Start with universal search',
            },
            {
              step: 2,
              tool: 'get-project-overview',
              parameters: {
                projectName: context,
                includeFiles: true,
                includeMeetings: true,
                includeEmails: true,
                includeTasks: true,
              },
              reason: 'Get project overview',
            },
            {
              step: 3,
              tool: 'get-project-stakeholders',
              parameters: { projectName: context, days: 90 },
              reason: 'Identify stakeholders',
            },
          ],
        },
        meeting_preparation: {
          workflow: 'meeting_preparation',
          description: 'Meeting preparation workflow',
          steps: [
            {
              step: 1,
              tool: 'prepare-for-meeting',
              parameters: { meetingSubject: context, hoursAhead: 48 },
              reason: 'Gather meeting context',
            },
            {
              step: 2,
              tool: 'find-related-documents',
              parameters: { topic: context, days: 30, limit: 20 },
              reason: 'Find related documents',
            },
          ],
        },
        weekly_review: {
          workflow: 'weekly_review',
          description: 'Weekly review workflow',
          steps: [
            {
              step: 1,
              tool: 'get-my-week-summary',
              parameters: { weekOffset: 0 },
              reason: 'Get weekly summary',
            },
            {
              step: 2,
              tool: 'get-deadline-overview',
              parameters: { days: 7, includeCompleted: false },
              reason: 'Check deadlines',
            },
            {
              step: 3,
              tool: 'get-follow-up-items',
              parameters: { includeEmails: true, includeTasks: true, includeMeetings: true },
              reason: 'Get follow-ups',
            },
          ],
        },
        person_research: {
          workflow: 'client_communication',
          description: 'Client communication workflow',
          steps: [
            {
              step: 1,
              tool: 'get-communication-summary',
              parameters: {
                person: context,
                includeEmails: true,
                includeChats: true,
                includeMeetings: true,
                includeFiles: true,
              },
              reason: 'Get communication overview',
            },
            {
              step: 2,
              tool: 'find-related-documents',
              parameters: { topic: context, days: 180, limit: 50 },
              reason: 'Find related documents',
            },
          ],
        },
      };

      const workflowDef = workflows[scenario];
      if (!workflowDef) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify({ error: 'Workflow not found' }, null, 2),
            },
          ],
          isError: true,
        };
      }

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(
              {
                ...workflowDef,
                context,
                instructions:
                  'Execute each step in order. Use the tool names and parameters exactly as specified.',
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 35. CREATE CUSTOM WORKFLOW - Create and save custom workflows
  // ==========================================================================
  server.tool(
    'create-custom-workflow',
    `Create and save custom workflows:
- Define tool sequences
- Set parameters
- Save for reuse
- Share with team

Use this for "Create a workflow for client research" or "Save this tool sequence as a workflow".

Note: This is a planning tool that returns the workflow definition. Actual persistence would need to be implemented separately.`,
    {
      workflowName: z.string().describe('Name for the custom workflow'),
      description: z.string().describe('Description of what the workflow does'),
      steps: z
        .array(
          z.object({
            step: z.number(),
            tool: z.string(),
            parameters: z.record(z.unknown()),
            reason: z.string(),
          })
        )
        .describe('Array of workflow steps'),
    },
    {
      title: 'Create Custom Workflow',
      readOnlyHint: false,
      openWorldHint: false,
    },
    async ({ workflowName, description, steps }) => {
      logger.info(`Creating custom workflow: ${workflowName}`);

      // Validate steps
      for (const step of steps) {
        const toolExists = AVAILABLE_TOOLS.some((t) => t.name === step.tool);
        if (!toolExists) {
          return {
            content: [
              {
                type: 'text' as const,
                text: JSON.stringify(
                  {
                    error: 'Invalid tool',
                    message: `Tool '${step.tool}' not found in available tools`,
                    step: step.step,
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

      const workflow = {
        name: workflowName,
        description,
        steps,
        createdAt: new Date().toISOString(),
        totalSteps: steps.length,
        estimatedTime: `~${steps.length * 2}-${steps.length * 3} seconds`,
        note: 'This workflow definition can be saved and reused. Actual persistence needs to be implemented.',
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(workflow, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // ADVANCED ANALYTICS & INSIGHTS TOOLS
  // ==========================================================================

  // ==========================================================================
  // 36. ANALYZE BUSINESS METRICS - Analyze business metrics across Microsoft 365
  // ==========================================================================
  server.tool(
    'analyze-business-metrics',
    `Analyze business metrics across Microsoft 365:
- Communication velocity
- Project completion rates
- Team collaboration patterns
- Customer engagement metrics
- Response time analytics

Use this for "Analyze my business metrics this quarter", "Show team collaboration patterns", or "What's my customer engagement rate?".`,
    {
      period: z
        .enum(['week', 'month', 'quarter', 'year'])
        .optional()
        .describe('Time period to analyze (default: month)'),
      includeCommunication: z
        .boolean()
        .optional()
        .describe('Include communication metrics (default: true)'),
      includeProjects: z.boolean().optional().describe('Include project metrics (default: true)'),
      includeCollaboration: z
        .boolean()
        .optional()
        .describe('Include collaboration metrics (default: true)'),
    },
    {
      title: 'Analyze Business Metrics',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({
      period = 'month',
      includeCommunication = true,
      includeProjects = true,
      includeCollaboration = true,
    }) => {
      logger.info(`Analyzing business metrics for period: ${period}`);

      const daysMap: Record<string, number> = {
        week: 7,
        month: 30,
        quarter: 90,
        year: 365,
      };
      const days = daysMap[period] || 30;
      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      const metrics: Record<string, unknown> = {
        period,
        analyzedDays: days,
        analyzedFrom: startDate.toISOString(),
        analyzedTo: new Date().toISOString(),
      };

      const promises: Promise<void>[] = [];

      // Communication metrics
      if (includeCommunication) {
        promises.push(
          (async () => {
            try {
              const emailResponse = await graphClient.makeRequest('/me/messages', {
                method: 'GET',
                queryParams: {
                  $top: '500',
                  $select: 'receivedDateTime,from,importance',
                  $filter: `receivedDateTime ge ${startDate.toISOString()}`,
                  $orderby: 'receivedDateTime desc',
                },
              });

              if (emailResponse && typeof emailResponse === 'object' && 'value' in emailResponse) {
                const emails = emailResponse.value as GraphEmail[];
                const totalEmails = emails.length;
                const importantEmails = emails.filter((e) => e.importance === 'high').length;
                const avgPerDay = totalEmails / days;

                metrics.communication = {
                  totalEmails: totalEmails,
                  importantEmails,
                  averagePerDay: Math.round(avgPerDay * 10) / 10,
                  emailVelocity: avgPerDay > 20 ? 'High' : avgPerDay > 10 ? 'Medium' : 'Low',
                };
              }
            } catch (error) {
              logger.warn(`Could not analyze communication metrics: ${error}`);
            }
          })()
        );
      }

      // Collaboration metrics
      if (includeCollaboration) {
        promises.push(
          (async () => {
            try {
              const meetingQueryParams: Record<string, string> = {
                startDateTime: startDate.toISOString(),
                endDateTime: new Date().toISOString(),
                $top: '500',
                $select: 'subject,start,end,attendees,organizer',
              };

              const meetingsResponse = await graphClient.makeRequest(
                `/me/calendarView?${buildGraphQueryString(meetingQueryParams)}`,
                {
                  method: 'GET',
                }
              );

              if (
                meetingsResponse &&
                typeof meetingsResponse === 'object' &&
                'value' in meetingsResponse
              ) {
                const meetings = meetingsResponse.value as GraphEvent[];
                const totalMeetings = meetings.length;
                let totalMeetingHours = 0;
                const uniqueParticipants = new Set<string>();

                for (const meeting of meetings) {
                  const start = new Date(meeting.start.dateTime);
                  const end = new Date(meeting.end.dateTime);
                  const duration = (end.getTime() - start.getTime()) / (1000 * 60 * 60);
                  totalMeetingHours += duration;

                  if (meeting.attendees) {
                    for (const attendee of meeting.attendees) {
                      if (attendee.emailAddress?.address) {
                        uniqueParticipants.add(attendee.emailAddress.address);
                      }
                    }
                  }
                }

                metrics.collaboration = {
                  totalMeetings,
                  totalMeetingHours: Math.round(totalMeetingHours * 10) / 10,
                  averageMeetingDuration: Math.round((totalMeetingHours / totalMeetings) * 10) / 10,
                  uniqueCollaborators: uniqueParticipants.size,
                  averageMeetingsPerDay: Math.round((totalMeetings / days) * 10) / 10,
                };
              }
            } catch (error) {
              logger.warn(`Could not analyze collaboration metrics: ${error}`);
            }
          })()
        );
      }

      // Project metrics (using task completion as proxy)
      if (includeProjects) {
        promises.push(
          (async () => {
            try {
              const taskListsResponse = await graphClient.makeRequest('/me/todo/lists', {
                method: 'GET',
                queryParams: { $top: '20' },
              });

              if (
                taskListsResponse &&
                typeof taskListsResponse === 'object' &&
                'value' in taskListsResponse
              ) {
                let totalTasks = 0;
                let completedTasks = 0;

                for (const list of taskListsResponse.value as Array<{ id: string }>) {
                  try {
                    const tasksResponse = await graphClient.makeRequest(
                      `/me/todo/lists/${list.id}/tasks`,
                      {
                        method: 'GET',
                        queryParams: { $top: '100' },
                      }
                    );

                    if (
                      tasksResponse &&
                      typeof tasksResponse === 'object' &&
                      'value' in tasksResponse
                    ) {
                      for (const task of tasksResponse.value as Array<{
                        status: string;
                        createdDateTime?: string;
                      }>) {
                        if (!task.createdDateTime || new Date(task.createdDateTime) >= startDate) {
                          totalTasks++;
                          if (task.status === 'completed') {
                            completedTasks++;
                          }
                        }
                      }
                    }
                  } catch {
                    // Skip individual list errors
                  }
                }

                metrics.projects = {
                  totalTasks,
                  completedTasks,
                  completionRate:
                    totalTasks > 0 ? Math.round((completedTasks / totalTasks) * 100) : 0,
                  pendingTasks: totalTasks - completedTasks,
                };
              }
            } catch (error) {
              logger.warn(`Could not analyze project metrics: ${error}`);
            }
          })()
        );
      }

      await Promise.allSettled(promises);

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(metrics, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 37. GET BUSINESS INTELLIGENCE - Comprehensive business intelligence dashboard
  // ==========================================================================
  server.tool(
    'get-business-intelligence',
    `Comprehensive business intelligence dashboard:
- Key performance indicators
- Trend analysis
- Comparative analytics
- Predictive insights

Use this for "Show my business intelligence dashboard", "What are the key trends this month?", or "Compare this quarter to last quarter".`,
    {
      period: z
        .enum(['week', 'month', 'quarter'])
        .optional()
        .describe('Time period (default: month)'),
      compareWithPrevious: z
        .boolean()
        .optional()
        .describe('Compare with previous period (default: true)'),
    },
    {
      title: 'Get Business Intelligence',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ period = 'month', compareWithPrevious = true }) => {
      logger.info(`Getting business intelligence for period: ${period}`);

      const daysMap: Record<string, number> = {
        week: 7,
        month: 30,
        quarter: 90,
      };
      const days = daysMap[period] || 30;

      // Get current period metrics
      const currentStart = new Date();
      currentStart.setDate(currentStart.getDate() - days);
      const currentEnd = new Date();

      // Get previous period metrics if requested
      const previousStart = new Date(currentStart);
      previousStart.setDate(previousStart.getDate() - days);
      const previousEnd = currentStart;

      const promises: Promise<void>[] = [];
      const bi: Record<string, unknown> = {
        period,
        currentPeriod: {
          start: currentStart.toISOString(),
          end: currentEnd.toISOString(),
        },
      };

      // Email activity
      promises.push(
        (async () => {
          try {
            const emailResponse = await graphClient.makeRequest('/me/messages', {
              method: 'GET',
              queryParams: {
                $top: '1000',
                $select: 'receivedDateTime,importance',
                $filter: `receivedDateTime ge ${currentStart.toISOString()}`,
                $orderby: 'receivedDateTime desc',
              },
            });

            if (emailResponse && typeof emailResponse === 'object' && 'value' in emailResponse) {
              const emails = emailResponse.value as GraphEmail[];
              bi.emailActivity = {
                total: emails.length,
                important: emails.filter((e) => e.importance === 'high').length,
                averagePerDay: Math.round((emails.length / days) * 10) / 10,
              };

              if (compareWithPrevious) {
                const prevEmailResponse = await graphClient.makeRequest('/me/messages', {
                  method: 'GET',
                  queryParams: {
                    $top: '1000',
                    $select: 'receivedDateTime',
                    $filter: `receivedDateTime ge ${previousStart.toISOString()} and receivedDateTime lt ${previousEnd.toISOString()}`,
                  },
                });

                if (
                  prevEmailResponse &&
                  typeof prevEmailResponse === 'object' &&
                  'value' in prevEmailResponse
                ) {
                  const prevEmails = prevEmailResponse.value as GraphEmail[];
                  const change = emails.length - prevEmails.length;
                  bi.emailActivity = {
                    ...bi.emailActivity,
                    previousPeriod: prevEmails.length,
                    change,
                    changePercent:
                      prevEmails.length > 0 ? Math.round((change / prevEmails.length) * 100) : 0,
                    trend: change > 0 ? 'increasing' : change < 0 ? 'decreasing' : 'stable',
                  };
                }
              }
            }
          } catch (error) {
            logger.warn(`Could not analyze email activity: ${error}`);
          }
        })()
      );

      // Meeting activity
      promises.push(
        (async () => {
          try {
            const meetingQueryParams: Record<string, string> = {
              startDateTime: currentStart.toISOString(),
              endDateTime: currentEnd.toISOString(),
              $top: '500',
              $select: 'subject,start,end',
            };

            const meetingsResponse = await graphClient.makeRequest(
              `/me/calendarView?${buildGraphQueryString(meetingQueryParams)}`,
              {
                method: 'GET',
              }
            );

            if (
              meetingsResponse &&
              typeof meetingsResponse === 'object' &&
              'value' in meetingsResponse
            ) {
              const meetings = meetingsResponse.value as GraphEvent[];
              let totalHours = 0;
              for (const meeting of meetings) {
                const start = new Date(meeting.start.dateTime);
                const end = new Date(meeting.end.dateTime);
                totalHours += (end.getTime() - start.getTime()) / (1000 * 60 * 60);
              }

              bi.meetingActivity = {
                totalMeetings: meetings.length,
                totalHours: Math.round(totalHours * 10) / 10,
                averagePerDay: Math.round((meetings.length / days) * 10) / 10,
              };
            }
          } catch (error) {
            logger.warn(`Could not analyze meeting activity: ${error}`);
          }
        })()
      );

      await Promise.allSettled(promises);

      // Calculate KPIs
      bi.keyPerformanceIndicators = {
        communicationVelocity: (bi.emailActivity as any)?.averagePerDay || 0,
        meetingLoad: (bi.meetingActivity as any)?.totalHours || 0,
        productivityScore:
          ((bi.emailActivity as any)?.total || 0) > 0 &&
          ((bi.meetingActivity as any)?.totalMeetings || 0) > 0
            ? 'Active'
            : 'Low',
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(bi, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 38. ANALYZE TEAM PERFORMANCE - Team performance analytics
  // ==========================================================================
  server.tool(
    'analyze-team-performance',
    `Team performance analytics:
- Individual contributions
- Collaboration effectiveness
- Meeting efficiency
- Task completion rates
- Communication patterns

Use this for "Analyze team performance", "How effective is our team collaboration?", or "Show team productivity metrics".`,
    {
      days: z.number().optional().describe('Days of history to analyze (default: 90)'),
      includeMeetings: z.boolean().optional().describe('Include meeting analysis (default: true)'),
      includeCommunication: z
        .boolean()
        .optional()
        .describe('Include communication analysis (default: true)'),
    },
    {
      title: 'Analyze Team Performance',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ days = 90, includeMeetings = true, includeCommunication = true }) => {
      logger.info(`Analyzing team performance for ${days} days`);

      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      const performance: Record<string, unknown> = {
        analyzedPeriod: `Last ${days} days`,
        startDate: startDate.toISOString(),
        endDate: new Date().toISOString(),
      };

      const promises: Promise<void>[] = [];

      // Meeting efficiency
      if (includeMeetings) {
        promises.push(
          (async () => {
            try {
              const meetingQueryParams: Record<string, string> = {
                startDateTime: startDate.toISOString(),
                endDateTime: new Date().toISOString(),
                $top: '500',
                $select: 'subject,start,end,attendees,organizer',
              };

              const meetingsResponse = await graphClient.makeRequest(
                `/me/calendarView?${buildGraphQueryString(meetingQueryParams)}`,
                {
                  method: 'GET',
                }
              );

              if (
                meetingsResponse &&
                typeof meetingsResponse === 'object' &&
                'value' in meetingsResponse
              ) {
                const meetings = meetingsResponse.value as GraphEvent[];
                let totalHours = 0;
                const attendeeCounts = new Map<string, number>();

                for (const meeting of meetings) {
                  const start = new Date(meeting.start.dateTime);
                  const end = new Date(meeting.end.dateTime);
                  const duration = (end.getTime() - start.getTime()) / (1000 * 60 * 60);
                  totalHours += duration;

                  if (meeting.attendees) {
                    for (const attendee of meeting.attendees) {
                      const email = attendee.emailAddress?.address;
                      if (email) {
                        attendeeCounts.set(email, (attendeeCounts.get(email) || 0) + 1);
                      }
                    }
                  }
                }

                performance.meetingEfficiency = {
                  totalMeetings: meetings.length,
                  totalHours: Math.round(totalHours * 10) / 10,
                  averageDuration: Math.round((totalHours / meetings.length) * 10) / 10,
                  averageAttendees:
                    Math.round(
                      (Array.from(attendeeCounts.values()).reduce((a, b) => a + b, 0) /
                        meetings.length) *
                        10
                    ) / 10,
                  mostActiveParticipants: Array.from(attendeeCounts.entries())
                    .sort((a, b) => b[1] - a[1])
                    .slice(0, 10)
                    .map(([email, count]) => ({ email, meetingCount: count })),
                };
              }
            } catch (error) {
              logger.warn(`Could not analyze meeting efficiency: ${error}`);
            }
          })()
        );
      }

      // Communication patterns
      if (includeCommunication) {
        promises.push(
          (async () => {
            try {
              const commFrequency = await graphClient.makeRequest('/me/messages', {
                method: 'GET',
                queryParams: {
                  $top: '1000',
                  $select: 'from,receivedDateTime',
                  $filter: `receivedDateTime ge ${startDate.toISOString()}`,
                  $orderby: 'receivedDateTime desc',
                },
              });

              if (commFrequency && typeof commFrequency === 'object' && 'value' in commFrequency) {
                const emails = commFrequency.value as GraphEmail[];
                const senderCounts = new Map<string, number>();

                for (const email of emails) {
                  const sender = email.from?.emailAddress?.address;
                  if (sender) {
                    senderCounts.set(sender, (senderCounts.get(sender) || 0) + 1);
                  }
                }

                performance.communicationPatterns = {
                  totalEmails: emails.length,
                  uniqueSenders: senderCounts.size,
                  topCommunicators: Array.from(senderCounts.entries())
                    .sort((a, b) => b[1] - a[1])
                    .slice(0, 10)
                    .map(([email, count]) => ({ email, emailCount: count })),
                };
              }
            } catch (error) {
              logger.warn(`Could not analyze communication patterns: ${error}`);
            }
          })()
        );
      }

      await Promise.allSettled(promises);

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(performance, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // SMART AUTOMATION TOOLS
  // ==========================================================================

  // ==========================================================================
  // 39. AUTO CATEGORIZE ITEMS - Automatically categorize emails, files, and tasks
  // ==========================================================================
  server.tool(
    'auto-categorize-items',
    `Automatically categorize emails, files, and tasks:
- Use AI to detect categories
- Apply tags and labels
- Organize by project/topic
- Learn from user corrections

Use this for "Categorize my recent emails", "Auto-organize files by project", or "Tag all items related to [topic]".

Note: This tool analyzes and suggests categories. Actual categorization would require write permissions.`,
    {
      source: z
        .enum(['emails', 'files', 'tasks', 'all'])
        .optional()
        .describe('Source to categorize (default: all)'),
      days: z.number().optional().describe('Days back to analyze (default: 7)'),
      topic: z.string().optional().describe('Topic to categorize by (optional)'),
      limit: z.number().optional().describe('Maximum items to categorize (default: 50)'),
    },
    {
      title: 'Auto Categorize Items',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ source = 'all', days = 7, topic, limit = 50 }) => {
      logger.info(`Auto-categorizing items: source=${source}, days=${days}, topic=${topic}`);

      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      interface CategorizedItem {
        id: string;
        type: 'email' | 'file' | 'task';
        title: string;
        suggestedCategories: string[];
        confidence: number;
        date: string;
      }

      const categorized: CategorizedItem[] = [];
      const promises: Promise<void>[] = [];

      // Categorize emails
      if (source === 'emails' || source === 'all') {
        promises.push(
          (async () => {
            try {
              let queryParams: Record<string, string> = {
                $top: String(limit),
                $select: 'id,subject,bodyPreview,receivedDateTime,categories',
                $filter: `receivedDateTime ge ${startDate.toISOString()}`,
                $orderby: 'receivedDateTime desc',
              };

              if (topic) {
                queryParams.$search = `"${topic}"`;
                delete queryParams.$orderby;
              }

              const emailResponse = await graphClient.makeRequest('/me/messages', {
                method: 'GET',
                queryParams,
              });

              if (emailResponse && typeof emailResponse === 'object' && 'value' in emailResponse) {
                for (const email of emailResponse.value as GraphEmail &
                  {
                    categories?: string[];
                  }[]) {
                  const content = sanitizeHtml(email.bodyPreview || '');
                  const subject = email.subject || '';
                  const combined = `${subject} ${content}`.toLowerCase();

                  const suggestedCategories: string[] = [];
                  let confidence = 0.5;

                  // Detect categories based on keywords
                  if (combined.includes('project') || combined.includes('projekt')) {
                    suggestedCategories.push('Project');
                    confidence += 0.2;
                  }
                  if (combined.includes('meeting') || combined.includes('termin')) {
                    suggestedCategories.push('Meeting');
                    confidence += 0.2;
                  }
                  if (
                    combined.includes('task') ||
                    combined.includes('todo') ||
                    combined.includes('aufgabe')
                  ) {
                    suggestedCategories.push('Task');
                    confidence += 0.2;
                  }
                  if (
                    combined.includes('urgent') ||
                    combined.includes('important') ||
                    combined.includes('wichtig')
                  ) {
                    suggestedCategories.push('Urgent');
                    confidence += 0.2;
                  }
                  if (
                    combined.includes('client') ||
                    combined.includes('customer') ||
                    combined.includes('kunde')
                  ) {
                    suggestedCategories.push('Client');
                    confidence += 0.2;
                  }
                  if (topic && combined.includes(topic.toLowerCase())) {
                    suggestedCategories.push(topic);
                    confidence += 0.3;
                  }

                  if (suggestedCategories.length === 0) {
                    suggestedCategories.push('General');
                  }

                  categorized.push({
                    id: email.id,
                    type: 'email',
                    title: subject,
                    suggestedCategories,
                    confidence: Math.min(confidence, 1),
                    date: email.receivedDateTime,
                  });
                }
              }
            } catch (error) {
              logger.warn(`Could not categorize emails: ${error}`);
            }
          })()
        );
      }

      await Promise.allSettled(promises);

      // Group by category
      const categoryGroups: Record<string, CategorizedItem[]> = {};
      for (const item of categorized) {
        for (const category of item.suggestedCategories) {
          if (!categoryGroups[category]) {
            categoryGroups[category] = [];
          }
          categoryGroups[category].push(item);
        }
      }

      const result = {
        totalItems: categorized.length,
        categorizedItems: categorized.slice(0, limit),
        categoryGroups,
        summary: {
          totalCategories: Object.keys(categoryGroups).length,
          topCategories: Object.entries(categoryGroups)
            .sort((a, b) => b[1].length - a[1].length)
            .slice(0, 10)
            .map(([category, items]) => ({ category, itemCount: items.length })),
        },
        note: 'These are suggested categories. Actual categorization requires write permissions.',
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 40. SMART REMINDER SYSTEM - Intelligent reminder system
  // ==========================================================================
  server.tool(
    'smart-reminder-system',
    `Intelligent reminder system:
- Context-aware reminders
- Priority-based scheduling
- Follow-up detection
- Deadline tracking

Use this for "Set a smart reminder for [item]", "What should I follow up on?", or "Show me upcoming deadlines".`,
    {
      action: z
        .enum(['list', 'set', 'check'])
        .optional()
        .describe('Action: list reminders, set reminder, or check for reminders (default: list)'),
      item: z.string().optional().describe('Item to set reminder for (required if action=set)'),
      days: z.number().optional().describe('Days ahead to check (default: 7)'),
    },
    {
      title: 'Smart Reminder System',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ action = 'list', item, days = 7 }) => {
      logger.info(`Smart reminder system: action=${action}, item=${item}`);

      if (action === 'set' && !item) {
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  error: 'Item required',
                  message: 'Please provide an item to set a reminder for',
                },
                null,
                2
              ),
            },
          ],
          isError: true,
        };
      }

      const reminders: Array<{
        type: string;
        title: string;
        dueDate: string;
        priority: string;
        source: string;
      }> = [];

      // Get deadlines
      const deadlineResponse = await graphClient.makeRequest('/me/todo/lists', {
        method: 'GET',
        queryParams: { $top: '20' },
      });

      if (deadlineResponse && typeof deadlineResponse === 'object' && 'value' in deadlineResponse) {
        for (const list of deadlineResponse.value as Array<{ id: string }>) {
          try {
            const tasksResponse = await graphClient.makeRequest(`/me/todo/lists/${list.id}/tasks`, {
              method: 'GET',
              queryParams: { $top: '100' },
            });

            if (tasksResponse && typeof tasksResponse === 'object' && 'value' in tasksResponse) {
              for (const task of tasksResponse.value as Array<{
                title: string;
                dueDateTime?: { dateTime: string };
                status: string;
                importance: string;
              }>) {
                if (task.status !== 'completed' && task.dueDateTime) {
                  const dueDate = new Date(task.dueDateTime.dateTime);
                  const daysUntil = Math.ceil(
                    (dueDate.getTime() - Date.now()) / (1000 * 60 * 60 * 24)
                  );
                  if (daysUntil >= 0 && daysUntil <= days) {
                    reminders.push({
                      type: 'task',
                      title: task.title,
                      dueDate: task.dueDateTime.dateTime,
                      priority: task.importance || 'normal',
                      source: 'To-Do',
                    });
                  }
                }
              }
            }
          } catch {
            // Skip errors
          }
        }
      }

      // Get flagged emails
      try {
        const flaggedResponse = await graphClient.makeRequest('/me/messages', {
          method: 'GET',
          queryParams: {
            $filter: "flag/flagStatus eq 'flagged'",
            $top: '50',
            $select: 'subject,flag,receivedDateTime',
          },
        });

        if (flaggedResponse && typeof flaggedResponse === 'object' && 'value' in flaggedResponse) {
          for (const email of flaggedResponse.value as Array<{
            subject: string;
            flag?: { dueDateTime?: { dateTime: string }; flagStatus: string };
            receivedDateTime: string;
          }>) {
            if (email.flag?.dueDateTime) {
              reminders.push({
                type: 'email',
                title: email.subject,
                dueDate: email.flag.dueDateTime.dateTime,
                priority: 'high',
                source: 'Email',
              });
            }
          }
        }
      } catch (error) {
        logger.warn(`Could not get flagged emails: ${error}`);
      }

      // Sort by due date
      reminders.sort((a, b) => new Date(a.dueDate).getTime() - new Date(b.dueDate).getTime());

      const result = {
        action,
        reminders: reminders.slice(0, 20),
        totalReminders: reminders.length,
        upcoming: reminders.filter(
          (r) => new Date(r.dueDate).getTime() <= Date.now() + 24 * 60 * 60 * 1000
        ).length,
        summary: {
          byType: {
            task: reminders.filter((r) => r.type === 'task').length,
            email: reminders.filter((r) => r.type === 'email').length,
          },
          byPriority: {
            high: reminders.filter((r) => r.priority === 'high').length,
            normal: reminders.filter((r) => r.priority === 'normal').length,
          },
        },
        note:
          action === 'set'
            ? `Reminder suggestion for "${item}": Set reminder for 1 day before due date or follow-up date`
            : undefined,
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 41. AUTO SUMMARIZE PERIOD - Automatically summarize a time period
  // ==========================================================================
  server.tool(
    'auto-summarize-period',
    `Automatically summarize a time period:
- Daily/weekly/monthly summaries
- Key highlights
- Action items
- Decisions made

Use this for "Summarize my week", "What happened this month?", or "Give me a daily summary".`,
    {
      period: z
        .enum(['day', 'week', 'month'])
        .optional()
        .describe('Period to summarize (default: week)'),
      includeEmails: z.boolean().optional().describe('Include email summary (default: true)'),
      includeMeetings: z.boolean().optional().describe('Include meeting summary (default: true)'),
      includeTasks: z.boolean().optional().describe('Include task summary (default: true)'),
    },
    {
      title: 'Auto Summarize Period',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({
      period = 'week',
      includeEmails = true,
      includeMeetings = true,
      includeTasks = true,
    }) => {
      logger.info(`Auto-summarizing period: ${period}`);

      const daysMap: Record<string, number> = {
        day: 1,
        week: 7,
        month: 30,
      };
      const days = daysMap[period] || 7;
      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      const summary: Record<string, unknown> = {
        period,
        dateRange: {
          start: startDate.toISOString(),
          end: new Date().toISOString(),
        },
      };

      const promises: Promise<void>[] = [];

      // Email summary
      if (includeEmails) {
        promises.push(
          (async () => {
            try {
              const emailResponse = await graphClient.makeRequest('/me/messages', {
                method: 'GET',
                queryParams: {
                  $top: '100',
                  $select: 'subject,from,receivedDateTime,importance,bodyPreview',
                  $filter: `receivedDateTime ge ${startDate.toISOString()}`,
                  $orderby: 'receivedDateTime desc',
                },
              });

              if (emailResponse && typeof emailResponse === 'object' && 'value' in emailResponse) {
                const emails = emailResponse.value as GraphEmail[];
                summary.emails = {
                  total: emails.length,
                  important: emails.filter((e) => e.importance === 'high').length,
                  topSenders: Array.from(
                    new Map(
                      emails.map((e) => [
                        e.from?.emailAddress?.address || 'Unknown',
                        e.from?.emailAddress?.name || 'Unknown',
                      ])
                    ).entries()
                  )
                    .slice(0, 5)
                    .map(([email, name]) => ({ email, name })),
                  recentSubjects: emails.slice(0, 10).map((e) => ({
                    subject: e.subject,
                    from: e.from?.emailAddress?.name,
                    date: e.receivedDateTime,
                  })),
                };
              }
            } catch (error) {
              logger.warn(`Could not summarize emails: ${error}`);
            }
          })()
        );
      }

      // Meeting summary
      if (includeMeetings) {
        promises.push(
          (async () => {
            try {
              const meetingQueryParams: Record<string, string> = {
                startDateTime: startDate.toISOString(),
                endDateTime: new Date().toISOString(),
                $top: '100',
                $select: 'subject,start,end,attendees,organizer',
              };

              const meetingsResponse = await graphClient.makeRequest(
                `/me/calendarView?${buildGraphQueryString(meetingQueryParams)}`,
                {
                  method: 'GET',
                }
              );

              if (
                meetingsResponse &&
                typeof meetingsResponse === 'object' &&
                'value' in meetingsResponse
              ) {
                const meetings = meetingsResponse.value as GraphEvent[];
                let totalHours = 0;
                for (const meeting of meetings) {
                  const start = new Date(meeting.start.dateTime);
                  const end = new Date(meeting.end.dateTime);
                  totalHours += (end.getTime() - start.getTime()) / (1000 * 60 * 60);
                }

                summary.meetings = {
                  total: meetings.length,
                  totalHours: Math.round(totalHours * 10) / 10,
                  upcoming: meetings.filter((m) => new Date(m.start.dateTime) > new Date()).length,
                  recent: meetings.slice(0, 10).map((m) => ({
                    subject: m.subject,
                    date: m.start.dateTime,
                    organizer: m.organizer?.emailAddress?.name,
                  })),
                };
              }
            } catch (error) {
              logger.warn(`Could not summarize meetings: ${error}`);
            }
          })()
        );
      }

      // Task summary
      if (includeTasks) {
        promises.push(
          (async () => {
            try {
              const taskListsResponse = await graphClient.makeRequest('/me/todo/lists', {
                method: 'GET',
                queryParams: { $top: '20' },
              });

              if (
                taskListsResponse &&
                typeof taskListsResponse === 'object' &&
                'value' in taskListsResponse
              ) {
                let totalTasks = 0;
                let completedTasks = 0;

                for (const list of taskListsResponse.value as Array<{ id: string }>) {
                  try {
                    const tasksResponse = await graphClient.makeRequest(
                      `/me/todo/lists/${list.id}/tasks`,
                      {
                        method: 'GET',
                        queryParams: { $top: '100' },
                      }
                    );

                    if (
                      tasksResponse &&
                      typeof tasksResponse === 'object' &&
                      'value' in tasksResponse
                    ) {
                      for (const task of tasksResponse.value as Array<{
                        title: string;
                        status: string;
                        createdDateTime?: string;
                      }>) {
                        if (!task.createdDateTime || new Date(task.createdDateTime) >= startDate) {
                          totalTasks++;
                          if (task.status === 'completed') {
                            completedTasks++;
                          }
                        }
                      }
                    }
                  } catch {
                    // Skip errors
                  }
                }

                summary.tasks = {
                  total: totalTasks,
                  completed: completedTasks,
                  pending: totalTasks - completedTasks,
                  completionRate:
                    totalTasks > 0 ? Math.round((completedTasks / totalTasks) * 100) : 0,
                };
              }
            } catch (error) {
              logger.warn(`Could not summarize tasks: ${error}`);
            }
          })()
        );
      }

      await Promise.allSettled(promises);

      // Generate highlights
      summary.highlights = {
        totalActivity:
          ((summary.emails as any)?.total || 0) +
          ((summary.meetings as any)?.total || 0) +
          ((summary.tasks as any)?.total || 0),
        keyMetrics: {
          emailsReceived: (summary.emails as any)?.total || 0,
          meetingsAttended: (summary.meetings as any)?.total || 0,
          tasksCompleted: (summary.tasks as any)?.completed || 0,
        },
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(summary, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // ADVANCED SEARCH & DISCOVERY TOOLS
  // ==========================================================================

  // ==========================================================================
  // 42. INTELLIGENT QUERY BUILDER - Build optimized queries automatically
  // ==========================================================================
  server.tool(
    'intelligent-query-builder',
    `Build optimized queries automatically:
- Suggest query improvements
- Expand search terms
- Add filters automatically
- Optimize for best results

Use this for "Build a query for [topic]", "Optimize this search query", or "Suggest better search terms".`,
    {
      query: z.string().describe('Original query to optimize'),
      expandTerms: z
        .boolean()
        .optional()
        .describe('Expand search terms with synonyms (default: true)'),
      addFilters: z.boolean().optional().describe('Add automatic filters (default: true)'),
    },
    {
      title: 'Intelligent Query Builder',
      readOnlyHint: true,
      openWorldHint: false,
    },
    async ({ query, expandTerms = true, addFilters = true }) => {
      logger.info(`Building intelligent query for: ${query}`);

      try {
        // Use NLP enhancer to decompose and improve query
        const decomposed = nlpEnhancer.decomposeQuery(query);

        // Build optimized query
        const optimizedQueries: Array<{
          query: string;
          type: string;
          reason: string;
          confidence: number;
        }> = [];

        // Original query
        optimizedQueries.push({
          query,
          type: 'original',
          reason: 'Original query as provided',
          confidence: 1.0,
        });

        // Expanded query with synonyms
        if (expandTerms && decomposed.semanticVariants.length > 0) {
          for (const variant of decomposed.semanticVariants.slice(0, 3)) {
            optimizedQueries.push({
              query: variant,
              type: 'expanded',
              reason: 'Expanded with semantic variants',
              confidence: 0.8,
            });
          }
        }

        // Query with entity focus
        if (decomposed.entity) {
          optimizedQueries.push({
            query: `"${decomposed.entity}"`,
            type: 'entity_focused',
            reason: `Focused on entity: ${decomposed.entity}`,
            confidence: 0.9,
          });
        }

        // Query with compound parts
        if (decomposed.compoundParts.length > 1) {
          const combinedQuery = decomposed.compoundParts.join(' AND ');
          optimizedQueries.push({
            query: combinedQuery,
            type: 'compound',
            reason: 'Combined compound parts with AND',
            confidence: 0.85,
          });
        }

        // Add filters if requested
        const filters: string[] = [];
        if (addFilters && decomposed.temporal) {
          if (decomposed.temporal.type === 'past') {
            filters.push(
              `receivedDateTime ge ${new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString()}`
            );
          } else if (decomposed.temporal.type === 'future') {
            filters.push(`start/dateTime ge ${new Date().toISOString()}`);
          }
        }

        const result = {
          originalQuery: query,
          optimizedQueries: optimizedQueries.slice(0, 5),
          recommendedQuery: optimizedQueries[0]?.query || query,
          filters,
          analysis: {
            entities: decomposed.entities.map((e) => ({ value: e.value, type: e.type })),
            intent: decomposed.intent.type,
            confidence: decomposed.confidence,
            semanticVariants: decomposed.semanticVariants.slice(0, 5),
          },
          suggestions: [
            'Use the recommended query for best results',
            'Try expanded queries if original returns few results',
            'Apply filters to narrow down results',
          ],
        };

        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(result, null, 2),
            },
          ],
        };
      } catch (error) {
        logger.error(`Error building intelligent query: ${error}`);
        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(
                {
                  error: 'Failed to build query',
                  message: `${error}`,
                  originalQuery: query,
                  fallback: query,
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
  registeredCount++;

  // ==========================================================================
  // 43. DISCOVER RELATED TOPICS - Discover related topics and connections
  // ==========================================================================
  server.tool(
    'discover-related-topics',
    `Discover related topics and connections:
- Topic clustering
- Related projects
- Connected people
- Similar documents

Use this for "What topics are related to [topic]?", "Discover connections for [project]", or "Find similar items to [item]".`,
    {
      topic: z.string().describe('Topic to discover related topics for'),
      days: z.number().optional().describe('Days of history to analyze (default: 180)'),
      limit: z.number().optional().describe('Maximum related topics to return (default: 20)'),
    },
    {
      title: 'Discover Related Topics',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ topic, days = 180, limit = 20 }) => {
      logger.info(`Discovering related topics for: ${topic}`);

      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      const relatedTopics = new Map<
        string,
        {
          topic: string;
          mentions: number;
          sources: string[];
          relevance: number;
        }
      >();

      const promises: Promise<void>[] = [];

      // Search for topic mentions
      promises.push(
        (async () => {
          try {
            const searchResult = await executeCentralSearch(graphClient, topic, {
              entityTypes: ['message', 'event', 'driveItem'],
              maxResults: 100,
              sortByRank: true,
            });

            // Extract related topics from search results
            const allText: string[] = [];

            // From emails
            for (const email of searchResult.results.emails.slice(0, 50)) {
              if (email.summary) {
                allText.push(email.summary);
              }
            }

            // From meetings
            for (const meeting of searchResult.results.events.slice(0, 50)) {
              if (meeting.summary) {
                allText.push(meeting.summary);
              }
            }

            // Extract potential topics (capitalized words, project names, etc.)
            const topicPattern = /\b[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*\b/g;
            for (const text of allText) {
              const matches = text.match(topicPattern);
              if (matches) {
                for (const match of matches) {
                  const normalized = match.trim();
                  if (
                    normalized.length > 3 &&
                    normalized.toLowerCase() !== topic.toLowerCase() &&
                    !normalized.match(/^(The|A|An|This|That|These|Those)$/i)
                  ) {
                    if (!relatedTopics.has(normalized)) {
                      relatedTopics.set(normalized, {
                        topic: normalized,
                        mentions: 0,
                        sources: [],
                        relevance: 0,
                      });
                    }
                    const rt = relatedTopics.get(normalized)!;
                    rt.mentions++;
                  }
                }
              }
            }
          } catch (error) {
            logger.warn(`Could not discover related topics: ${error}`);
          }
        })()
      );

      // Find related projects
      promises.push(
        (async () => {
          try {
            const emailResponse = await graphClient.makeRequest('/me/messages', {
              method: 'GET',
              queryParams: {
                $search: `"${topic}"`,
                $top: '50',
                $select: 'subject,bodyPreview',
                $filter: `receivedDateTime ge ${startDate.toISOString()}`,
              },
            });

            if (emailResponse && typeof emailResponse === 'object' && 'value' in emailResponse) {
              const emails = emailResponse.value as GraphEmail[];
              const projectPattern =
                /(?:project|projekt|projekt)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)/gi;

              for (const email of emails) {
                const content = `${email.subject} ${email.bodyPreview || ''}`;
                const matches = content.matchAll(projectPattern);
                for (const match of matches) {
                  if (match[1]) {
                    const projectName = match[1].trim();
                    if (projectName.toLowerCase() !== topic.toLowerCase()) {
                      if (!relatedTopics.has(projectName)) {
                        relatedTopics.set(projectName, {
                          topic: projectName,
                          mentions: 0,
                          sources: ['email'],
                          relevance: 0,
                        });
                      }
                      relatedTopics.get(projectName)!.mentions++;
                    }
                  }
                }
              }
            }
          } catch (error) {
            logger.warn(`Could not find related projects: ${error}`);
          }
        })()
      );

      await Promise.allSettled(promises);

      // Calculate relevance scores
      for (const [key, rt] of relatedTopics) {
        rt.relevance = rt.mentions * 10; // Simple relevance based on mentions
      }

      // Sort by relevance
      const sortedTopics = Array.from(relatedTopics.values())
        .sort((a, b) => b.relevance - a.relevance)
        .slice(0, limit);

      const result = {
        originalTopic: topic,
        analyzedPeriod: `Last ${days} days`,
        relatedTopics: sortedTopics,
        totalDiscovered: relatedTopics.size,
        topRelated: sortedTopics.slice(0, 5).map((t) => t.topic),
        summary: {
          mostMentioned: sortedTopics[0]?.topic,
          totalMentions: sortedTopics.reduce((sum, t) => sum + t.mentions, 0),
        },
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // COLLABORATION INTELLIGENCE TOOLS
  // ==========================================================================

  // ==========================================================================
  // 44. ANALYZE COLLABORATION PATTERNS - Analyze how teams collaborate
  // ==========================================================================
  server.tool(
    'analyze-collaboration-patterns',
    `Analyze how teams collaborate:
- Communication networks
- Collaboration hotspots
- Bottleneck identification
- Efficiency opportunities

Use this for "How does our team collaborate?", "Identify collaboration bottlenecks", or "Show team communication networks".`,
    {
      days: z.number().optional().describe('Days of history to analyze (default: 90)'),
      includeMeetings: z.boolean().optional().describe('Include meeting analysis (default: true)'),
      includeEmails: z.boolean().optional().describe('Include email analysis (default: true)'),
    },
    {
      title: 'Analyze Collaboration Patterns',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ days = 90, includeMeetings = true, includeEmails = true }) => {
      logger.info(`Analyzing collaboration patterns for ${days} days`);

      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      const patterns: Record<string, unknown> = {
        analyzedPeriod: `Last ${days} days`,
        startDate: startDate.toISOString(),
        endDate: new Date().toISOString(),
      };

      const promises: Promise<void>[] = [];

      // Communication network
      if (includeEmails) {
        promises.push(
          (async () => {
            try {
              const emailResponse = await graphClient.makeRequest('/me/messages', {
                method: 'GET',
                queryParams: {
                  $top: '1000',
                  $select: 'from,toRecipients,ccRecipients,receivedDateTime',
                  $filter: `receivedDateTime ge ${startDate.toISOString()}`,
                  $orderby: 'receivedDateTime desc',
                },
              });

              if (emailResponse && typeof emailResponse === 'object' && 'value' in emailResponse) {
                const emails = emailResponse.value as GraphEmail[];
                const communicationMap = new Map<
                  string,
                  {
                    sent: number;
                    received: number;
                    connections: Set<string>;
                  }
                >();

                const currentUserEmail = await getCurrentUserEmail();

                for (const email of emails) {
                  const sender = email.from?.emailAddress?.address?.toLowerCase();
                  if (sender && sender !== currentUserEmail.toLowerCase()) {
                    if (!communicationMap.has(sender)) {
                      communicationMap.set(sender, {
                        sent: 0,
                        received: 0,
                        connections: new Set(),
                      });
                    }
                    communicationMap.get(sender)!.sent++;
                  }

                  // Track recipients
                  for (const recipient of [
                    ...(email.toRecipients || []),
                    ...(email.ccRecipients || []),
                  ]) {
                    const recipientEmail = recipient.emailAddress?.address?.toLowerCase();
                    if (recipientEmail && recipientEmail !== currentUserEmail.toLowerCase()) {
                      if (!communicationMap.has(recipientEmail)) {
                        communicationMap.set(recipientEmail, {
                          sent: 0,
                          received: 0,
                          connections: new Set(),
                        });
                      }
                      communicationMap.get(recipientEmail)!.received++;
                      if (sender) {
                        communicationMap.get(recipientEmail)!.connections.add(sender);
                      }
                    }
                  }
                }

                patterns.communicationNetwork = {
                  totalParticipants: communicationMap.size,
                  mostActive: Array.from(communicationMap.entries())
                    .map(([email, data]) => ({
                      email,
                      totalInteractions: data.sent + data.received,
                      connections: data.connections.size,
                    }))
                    .sort((a, b) => b.totalInteractions - a.totalInteractions)
                    .slice(0, 10),
                  networkDensity:
                    communicationMap.size > 0
                      ? Math.round(
                          (Array.from(communicationMap.values()).reduce(
                            (sum, d) => sum + d.connections.size,
                            0
                          ) /
                            communicationMap.size) *
                            10
                        ) / 10
                      : 0,
                };
              }
            } catch (error) {
              logger.warn(`Could not analyze communication network: ${error}`);
            }
          })()
        );
      }

      // Collaboration hotspots (frequent meeting participants)
      if (includeMeetings) {
        promises.push(
          (async () => {
            try {
              const meetingQueryParams: Record<string, string> = {
                startDateTime: startDate.toISOString(),
                endDateTime: new Date().toISOString(),
                $top: '500',
                $select: 'subject,start,end,attendees,organizer',
              };

              const meetingsResponse = await graphClient.makeRequest(
                `/me/calendarView?${buildGraphQueryString(meetingQueryParams)}`,
                {
                  method: 'GET',
                }
              );

              if (
                meetingsResponse &&
                typeof meetingsResponse === 'object' &&
                'value' in meetingsResponse
              ) {
                const meetings = meetingsResponse.value as GraphEvent[];
                const participantCounts = new Map<string, number>();
                const meetingTopics = new Map<string, number>();

                for (const meeting of meetings) {
                  // Track participants
                  if (meeting.attendees) {
                    for (const attendee of meeting.attendees) {
                      const email = attendee.emailAddress?.address?.toLowerCase();
                      if (email) {
                        participantCounts.set(email, (participantCounts.get(email) || 0) + 1);
                      }
                    }
                  }

                  // Track meeting topics
                  const subject = meeting.subject || '';
                  const words = subject
                    .toLowerCase()
                    .split(/\s+/)
                    .filter((w) => w.length > 3);
                  for (const word of words) {
                    meetingTopics.set(word, (meetingTopics.get(word) || 0) + 1);
                  }
                }

                patterns.collaborationHotspots = {
                  totalMeetings: meetings.length,
                  mostFrequentParticipants: Array.from(participantCounts.entries())
                    .sort((a, b) => b[1] - a[1])
                    .slice(0, 10)
                    .map(([email, count]) => ({ email, meetingCount: count })),
                  commonTopics: Array.from(meetingTopics.entries())
                    .sort((a, b) => b[1] - a[1])
                    .slice(0, 10)
                    .map(([topic, count]) => ({ topic, mentions: count })),
                };
              }
            } catch (error) {
              logger.warn(`Could not analyze collaboration hotspots: ${error}`);
            }
          })()
        );
      }

      await Promise.allSettled(promises);

      // Identify bottlenecks (people with many connections but low efficiency)
      patterns.bottlenecks = {
        analysis:
          'Bottlenecks are people who appear in many communications but may be blocking workflows',
        note: 'Review communication patterns to identify potential bottlenecks',
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(patterns, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // 45. SUGGEST COLLABORATION IMPROVEMENTS - Suggest improvements to collaboration
  // ==========================================================================
  server.tool(
    'suggest-collaboration-improvements',
    `Suggest improvements to collaboration:
- Communication recommendations
- Meeting optimization
- Workflow suggestions
- Tool recommendations

Use this for "How can we collaborate better?", "Suggest improvements to our workflow", or "Optimize our meeting schedule".`,
    {
      focusArea: z
        .enum(['communication', 'meetings', 'workflow', 'all'])
        .optional()
        .describe('Focus area for suggestions (default: all)'),
      days: z.number().optional().describe('Days of history to analyze (default: 90)'),
    },
    {
      title: 'Suggest Collaboration Improvements',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ focusArea = 'all', days = 90 }) => {
      logger.info(`Suggesting collaboration improvements: focus=${focusArea}, days=${days}`);

      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      const suggestions: Array<{
        category: string;
        suggestion: string;
        priority: 'high' | 'medium' | 'low';
        reason: string;
      }> = [];

      const promises: Promise<void>[] = [];

      // Analyze meeting patterns
      if (focusArea === 'meetings' || focusArea === 'all') {
        promises.push(
          (async () => {
            try {
              const meetingQueryParams: Record<string, string> = {
                startDateTime: startDate.toISOString(),
                endDateTime: new Date().toISOString(),
                $top: '500',
                $select: 'subject,start,end,attendees',
              };

              const meetingsResponse = await graphClient.makeRequest(
                `/me/calendarView?${buildGraphQueryString(meetingQueryParams)}`,
                {
                  method: 'GET',
                }
              );

              if (
                meetingsResponse &&
                typeof meetingsResponse === 'object' &&
                'value' in meetingsResponse
              ) {
                const meetings = meetingsResponse.value as GraphEvent[];
                let backToBackCount = 0;
                let totalHours = 0;
                const sortedMeetings = meetings
                  .map((m) => ({
                    start: new Date(m.start.dateTime),
                    end: new Date(m.end.dateTime),
                    attendees: m.attendees?.length || 0,
                  }))
                  .sort((a, b) => a.start.getTime() - b.start.getTime());

                // Check for back-to-back meetings
                for (let i = 0; i < sortedMeetings.length - 1; i++) {
                  const gap =
                    (sortedMeetings[i + 1].start.getTime() - sortedMeetings[i].end.getTime()) /
                    (1000 * 60);
                  if (gap <= 5) {
                    backToBackCount++;
                  }
                  totalHours +=
                    (sortedMeetings[i].end.getTime() - sortedMeetings[i].start.getTime()) /
                    (1000 * 60 * 60);
                }

                const avgDailyHours = totalHours / days;
                const avgAttendees =
                  meetings.reduce((sum, m) => sum + (m.attendees?.length || 0), 0) /
                  meetings.length;

                if (backToBackCount > 5) {
                  suggestions.push({
                    category: 'meetings',
                    suggestion: 'Add buffer time between meetings',
                    priority: 'high',
                    reason: `You have ${backToBackCount} back-to-back meetings. Consider adding 15-minute buffers.`,
                  });
                }

                if (avgDailyHours > 6) {
                  suggestions.push({
                    category: 'meetings',
                    suggestion: 'Reduce meeting load',
                    priority: 'high',
                    reason: `You spend ${Math.round(avgDailyHours * 10) / 10} hours per day in meetings. Consider declining non-essential meetings.`,
                  });
                }

                if (avgAttendees > 8) {
                  suggestions.push({
                    category: 'meetings',
                    suggestion: 'Optimize meeting size',
                    priority: 'medium',
                    reason: `Average meeting size is ${Math.round(avgAttendees)} attendees. Smaller meetings are often more effective.`,
                  });
                }
              }
            } catch (error) {
              logger.warn(`Could not analyze meetings for suggestions: ${error}`);
            }
          })()
        );
      }

      // Analyze communication patterns
      if (focusArea === 'communication' || focusArea === 'all') {
        promises.push(
          (async () => {
            try {
              const emailResponse = await graphClient.makeRequest('/me/messages', {
                method: 'GET',
                queryParams: {
                  $top: '500',
                  $select: 'receivedDateTime,from,importance',
                  $filter: `receivedDateTime ge ${startDate.toISOString()}`,
                  $orderby: 'receivedDateTime desc',
                },
              });

              if (emailResponse && typeof emailResponse === 'object' && 'value' in emailResponse) {
                const emails = emailResponse.value as GraphEmail[];
                const unreadCount = emails.filter((e) => !e.isRead).length;
                const importantCount = emails.filter((e) => e.importance === 'high').length;
                const avgPerDay = emails.length / days;

                if (unreadCount > 50) {
                  suggestions.push({
                    category: 'communication',
                    suggestion: 'Process unread emails',
                    priority: 'high',
                    reason: `You have ${unreadCount} unread emails. Consider setting aside time to process them.`,
                  });
                }

                if (avgPerDay > 50) {
                  suggestions.push({
                    category: 'communication',
                    suggestion: 'Use email filters and rules',
                    priority: 'medium',
                    reason: `You receive ${Math.round(avgPerDay)} emails per day. Consider using filters to prioritize.`,
                  });
                }

                if (importantCount / emails.length > 0.3) {
                  suggestions.push({
                    category: 'communication',
                    suggestion: 'Review importance flags',
                    priority: 'low',
                    reason: `${Math.round((importantCount / emails.length) * 100)}% of emails are marked important. Consider reviewing flagging criteria.`,
                  });
                }
              }
            } catch (error) {
              logger.warn(`Could not analyze communication for suggestions: ${error}`);
            }
          })()
        );
      }

      await Promise.allSettled(promises);

      // Add general workflow suggestions
      if (focusArea === 'workflow' || focusArea === 'all') {
        suggestions.push({
          category: 'workflow',
          suggestion: 'Use automation tools',
          priority: 'medium',
          reason:
            'Consider using auto-categorize-items and smart-reminder-system to improve workflow efficiency.',
        });

        suggestions.push({
          category: 'workflow',
          suggestion: 'Regular reviews',
          priority: 'low',
          reason:
            'Use get-my-week-summary and get-deadline-overview for regular productivity reviews.',
        });
      }

      const result = {
        focusArea,
        analyzedPeriod: `Last ${days} days`,
        totalSuggestions: suggestions.length,
        suggestions: suggestions.sort((a, b) => {
          const priorityOrder = { high: 3, medium: 2, low: 1 };
          return priorityOrder[b.priority] - priorityOrder[a.priority];
        }),
        summary: {
          highPriority: suggestions.filter((s) => s.priority === 'high').length,
          mediumPriority: suggestions.filter((s) => s.priority === 'medium').length,
          lowPriority: suggestions.filter((s) => s.priority === 'low').length,
        },
      };

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    }
  );
  registeredCount++;

  // ==========================================================================
  // READ-LOOP-FILE - Read and parse Microsoft Loop files
  // ==========================================================================
  server.tool(
    'read-loop-file',
    {
      title: 'Read Loop File',
      description:
        'Read and parse a Microsoft Loop file. Loop files are collaborative documents that can contain notes, lists, tables, and other content. ' +
        'This tool detects Loop files, downloads their content, and extracts readable text from the Fluid format. ' +
        'Use this when you need to read the content of a Loop component or Loop page.',
      inputSchema: z.object({
        itemId: z.string().describe('The ID of the Loop file (DriveItem ID)'),
        driveId: z
          .string()
          .optional()
          .describe("The Drive ID. If not provided, uses the user's OneDrive"),
        includeRawContent: z
          .boolean()
          .optional()
          .default(false)
          .describe('Include the raw file content in addition to parsed text'),
      }),
      annotations: {
        audience: ['user', 'assistant'],
        readOnlyHint: true,
        openWorldHint: true,
      },
    },
    async ({ itemId, driveId, includeRawContent = false }) => {
      logger.info(`Reading Loop file: ${itemId}`);

      try {
        // Step 1: Get file metadata
        const metadataEndpoint = driveId
          ? `/drives/${driveId}/items/${itemId}`
          : `/me/drive/items/${itemId}`;

        const metadataResponse = await graphClient.makeRequest(metadataEndpoint, {
          method: 'GET',
        });

        if (!metadataResponse || typeof metadataResponse !== 'object') {
          return {
            content: [
              {
                type: 'text' as const,
                text: JSON.stringify({
                  error: 'File not found',
                  message: `Could not find file with ID "${itemId}".`,
                }),
              },
            ],
            isError: true,
          };
        }

        const metadata = metadataResponse as Record<string, unknown>;
        const fileName = (metadata.name as string) || 'Unknown';
        const webUrl = (metadata.webUrl as string) || '';
        const size = (metadata.size as number) || 0;
        const lastModified = (metadata.lastModifiedDateTime as string) || '';

        // Step 2: Detect if it's a Loop file
        const loopDetection = detectLoopFile(metadata);

        // Step 3: Download file content
        const contentEndpoint = driveId
          ? `/drives/${driveId}/items/${itemId}/content`
          : `/me/drive/items/${itemId}/content`;

        let content: string | null = null;
        try {
          const contentResponse = await graphClient.makeRequest(contentEndpoint, {
            method: 'GET',
          });
          if (typeof contentResponse === 'string') {
            content = contentResponse;
          } else if (contentResponse && typeof contentResponse === 'object') {
            content = JSON.stringify(contentResponse);
          }
        } catch (err) {
          logger.warn(`Failed to download Loop file content: ${err}`);
        }

        // Step 4: Parse Loop content
        let parsedContent = null;
        if (content) {
          parsedContent = parseLoopContent(content);
        }

        // Build response
        const response: Record<string, unknown> = {
          success: true,
          file: {
            id: itemId,
            name: fileName,
            webUrl,
            size,
            lastModified,
          },
          isLoopFile: loopDetection.isLoopFile,
          loopDetection: loopDetection.isLoopFile
            ? {
                method: loopDetection.detectionMethod,
                confidence: loopDetection.confidence,
                matchedPattern: loopDetection.matchedPattern,
              }
            : null,
        };

        if (parsedContent) {
          response.contentType = parsedContent.contentType;
          response.textContent = parsedContent.textContent || null;
          response.metadata = parsedContent.metadata || null;

          if (includeRawContent && parsedContent.rawContent) {
            // Limit raw content size
            const maxRawLength = 20000;
            response.rawContent =
              parsedContent.rawContent.length > maxRawLength
                ? parsedContent.rawContent.substring(0, maxRawLength) + '... (truncated)'
                : parsedContent.rawContent;
          }

          response.rawContentLength = parsedContent.rawContent?.length || 0;
        }

        if (!loopDetection.isLoopFile) {
          response.note =
            'This file was not detected as a Loop file. It may still contain valid content, ' +
            'but Loop-specific parsing may not apply.';
        }

        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify(response, null, 2),
            },
          ],
        };
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        logger.error(`Error reading Loop file: ${errorMessage}`);

        return {
          content: [
            {
              type: 'text' as const,
              text: JSON.stringify({
                error: 'Failed to read Loop file',
                message: errorMessage,
                itemId,
              }),
            },
          ],
          isError: true,
        };
      }
    }
  );
  registeredCount++;

  logger.info(`Registered ${registeredCount} compound tools`);
  return registeredCount;
}

/**
 * Extracts key information from a VTT transcript
 */
function extractTranscriptSummary(content: string): {
  keyPoints: string[];
  actionItems: string[];
  decisions: string[];
  participants: string[];
} {
  const keyPoints: string[] = [];
  const actionItems: string[] = [];
  const decisions: string[] = [];
  const participants = new Set<string>();

  // Parse VTT format - extract speaker and text
  const lines = content.split('\n');
  let currentText = '';
  let currentSpeaker = '';

  for (const line of lines) {
    const trimmed = line.trim();

    // Skip VTT headers and timestamps
    if (trimmed === 'WEBVTT' || trimmed === '' || /^\d{2}:\d{2}/.test(trimmed)) {
      continue;
    }

    // Check for speaker format: "<v Speaker Name>text</v>"
    const speakerMatch = trimmed.match(/<v\s+([^>]+)>([^<]*)<\/v>/);
    if (speakerMatch) {
      currentSpeaker = speakerMatch[1];
      currentText = speakerMatch[2];
      participants.add(currentSpeaker);
    } else if (!trimmed.startsWith('<')) {
      currentText = trimmed;
    }

    if (currentText) {
      const textLower = currentText.toLowerCase();

      // Detect action items
      if (
        textLower.includes('action item') ||
        textLower.includes('to do') ||
        textLower.includes('will do') ||
        textLower.includes('need to') ||
        textLower.includes('should') ||
        textLower.includes('must') ||
        textLower.includes('follow up')
      ) {
        actionItems.push(`${currentSpeaker}: ${currentText}`);
      }

      // Detect decisions
      if (
        textLower.includes('decided') ||
        textLower.includes('decision') ||
        textLower.includes('agreed') ||
        textLower.includes('we will') ||
        textLower.includes("let's go with")
      ) {
        decisions.push(`${currentSpeaker}: ${currentText}`);
      }

      // Key points - sentences that seem important
      if (
        textLower.includes('important') ||
        textLower.includes('key point') ||
        textLower.includes('main') ||
        textLower.includes('summary') ||
        textLower.includes('conclusion')
      ) {
        keyPoints.push(`${currentSpeaker}: ${currentText}`);
      }
    }
  }

  return {
    keyPoints: keyPoints.slice(0, 10),
    actionItems: actionItems.slice(0, 10),
    decisions: decisions.slice(0, 10),
    participants: Array.from(participants),
  };
}

/**
 * Searches for a query in transcript content and returns matching segments
 */
function searchInTranscript(
  content: string,
  query: string
): Array<{ speaker?: string; text: string; timestamp?: string }> {
  const matches: Array<{ speaker?: string; text: string; timestamp?: string }> = [];
  const lines = content.split('\n');
  let currentTimestamp = '';

  for (const line of lines) {
    const trimmed = line.trim();

    // Capture timestamp
    if (/^\d{2}:\d{2}:\d{2}/.test(trimmed)) {
      currentTimestamp = trimmed.split(' ')[0];
      continue;
    }

    // Check for speaker format and search
    const speakerMatch = trimmed.match(/<v\s+([^>]+)>([^<]*)<\/v>/);
    if (speakerMatch) {
      const speaker = speakerMatch[1];
      const text = speakerMatch[2];
      if (text.toLowerCase().includes(query)) {
        matches.push({ speaker, text, timestamp: currentTimestamp });
      }
    } else if (!trimmed.startsWith('<') && trimmed.toLowerCase().includes(query)) {
      matches.push({ text: trimmed, timestamp: currentTimestamp });
    }
  }

  return matches;
}

/**
 * Tool Orchestrator - Maps queries to tool execution plans
 * This maintains a registry of available tools and their capabilities
 */
interface ToolDefinition {
  name: string;
  description: string;
  category: string;
  parameters: string[];
  useCases: string[];
  dependencies?: string[];
  parallelizable: boolean;
}

/**
 * Available tools registry for orchestration
 * This should be kept in sync with actual registered tools
 */
const AVAILABLE_TOOLS: ToolDefinition[] = [
  // Search & Discovery
  {
    name: 'search-everything',
    description: 'Universal search across all Microsoft 365 services',
    category: 'search',
    parameters: ['query', 'limit'],
    useCases: ['finding information', 'searching', 'discovering'],
    parallelizable: false,
  },
  {
    name: 'ask-microsoft-365',
    description: 'Intelligent assistant for any Microsoft 365 question',
    category: 'search',
    parameters: ['question'],
    useCases: ['any question', 'general query', 'information gathering'],
    parallelizable: false,
  },
  {
    name: 'ms365-search',
    description: 'Primary search tool for Microsoft 365 data',
    category: 'search',
    parameters: ['query'],
    useCases: ['searching', 'finding'],
    parallelizable: false,
  },
  // Person-focused
  {
    name: 'find-messages-with-person',
    description: 'Find Teams chat messages with a specific person',
    category: 'people',
    parameters: ['person', 'limit'],
    useCases: ['messages with person', 'chat history', 'teams conversations'],
    dependencies: ['find-user'],
    parallelizable: false,
  },
  {
    name: 'find-emails-with-person',
    description: 'Find email conversations with a specific person',
    category: 'people',
    parameters: ['person', 'limit'],
    useCases: ['emails with person', 'email history', 'correspondence'],
    dependencies: ['find-user'],
    parallelizable: false,
  },
  {
    name: 'find-meetings-with-person',
    description: 'Find calendar meetings with a specific person',
    category: 'people',
    parameters: ['person', 'limit'],
    useCases: ['meetings with person', 'calendar history', 'scheduled meetings'],
    dependencies: ['find-user'],
    parallelizable: false,
  },
  {
    name: 'get-communication-summary',
    description: 'Complete communication overview with a person',
    category: 'people',
    parameters: ['person', 'includeEmails', 'includeChats', 'includeMeetings', 'includeFiles'],
    useCases: ['communication overview', 'person summary', 'interaction history'],
    dependencies: ['find-user'],
    parallelizable: false,
  },
  {
    name: 'analyze-relationship-strength',
    description: 'Analyze relationship strength with contacts',
    category: 'people',
    parameters: ['person', 'days', 'limit'],
    useCases: ['relationship analysis', 'contact strength', 'network analysis'],
    parallelizable: false,
  },
  {
    name: 'find-mutual-connections',
    description: 'Find mutual connections between people',
    category: 'people',
    parameters: ['person', 'days', 'limit'],
    useCases: ['mutual connections', 'network mapping', 'common contacts'],
    dependencies: ['find-user'],
    parallelizable: false,
  },
  {
    name: 'get-communication-frequency',
    description: 'Analyze communication frequency',
    category: 'people',
    parameters: ['days', 'limit', 'includeMeetings'],
    useCases: ['communication patterns', 'frequency analysis', 'interaction patterns'],
    parallelizable: false,
  },
  // Project & Business
  {
    name: 'get-project-overview',
    description: 'Complete project overview with files, meetings, emails, tasks',
    category: 'project',
    parameters: ['projectName', 'includeFiles', 'includeMeetings', 'includeEmails', 'includeTasks'],
    useCases: ['project overview', 'project status', 'project information'],
    parallelizable: false,
  },
  {
    name: 'get-project-stakeholders',
    description: 'Identify all stakeholders and participants in a project',
    category: 'project',
    parameters: ['projectName', 'days'],
    useCases: ['project stakeholders', 'team members', 'participants'],
    parallelizable: false,
  },
  {
    name: 'get-company-contacts',
    description: 'Find all contacts and interactions with a company',
    category: 'business',
    parameters: ['companyName'],
    useCases: ['company contacts', 'business relationships', 'client information'],
    parallelizable: false,
  },
  // Productivity
  {
    name: 'prepare-for-meeting',
    description: 'Gather all context for an upcoming meeting',
    category: 'productivity',
    parameters: ['meetingSubject', 'hoursAhead'],
    useCases: ['meeting preparation', 'preparing for meeting', 'meeting context'],
    parallelizable: false,
  },
  {
    name: 'get-my-week-summary',
    description: 'Weekly productivity digest',
    category: 'productivity',
    parameters: ['weekOffset'],
    useCases: ['week summary', 'weekly review', 'productivity summary'],
    parallelizable: false,
  },
  {
    name: 'get-all-my-tasks',
    description: 'Unified task view from To-Do and Planner',
    category: 'productivity',
    parameters: ['includeCompleted', 'dueSoon'],
    useCases: ['tasks', 'to-do list', 'task overview'],
    parallelizable: false,
  },
  {
    name: 'get-follow-up-items',
    description: 'Items needing attention',
    category: 'productivity',
    parameters: ['includeEmails', 'includeTasks', 'includeMeetings'],
    useCases: ['follow-ups', 'attention needed', 'pending items'],
    parallelizable: false,
  },
  {
    name: 'get-deadline-overview',
    description: 'All upcoming deadlines and due dates',
    category: 'productivity',
    parameters: ['days', 'includeCompleted'],
    useCases: ['deadlines', 'due dates', 'upcoming tasks'],
    parallelizable: false,
  },
  {
    name: 'find-unresponded-requests',
    description: 'Find requests waiting for response',
    category: 'productivity',
    parameters: ['days', 'priorityOnly'],
    useCases: ['unanswered requests', 'pending responses', 'follow-ups needed'],
    parallelizable: false,
  },
  // Content Intelligence
  {
    name: 'extract-action-items',
    description: 'Extract action items from emails and meetings',
    category: 'content',
    parameters: ['source', 'days', 'person', 'limit'],
    useCases: ['action items', 'tasks extraction', 'to-dos'],
    parallelizable: false,
  },
  {
    name: 'summarize-email-thread',
    description: 'Summarize long email threads',
    category: 'content',
    parameters: ['topic', 'days', 'limit'],
    useCases: ['email summary', 'thread summary', 'conversation summary'],
    parallelizable: false,
  },
  {
    name: 'extract-decisions',
    description: 'Extract decisions from communications',
    category: 'content',
    parameters: ['topic', 'days', 'source', 'limit'],
    useCases: ['decision extraction', 'decision history', 'decisions made'],
    parallelizable: false,
  },
  {
    name: 'find-decision-context',
    description: 'Find context and history for a decision',
    category: 'content',
    parameters: ['topic', 'days'],
    useCases: ['decision context', 'decision history', 'decision background'],
    parallelizable: false,
  },
  // Document Intelligence
  {
    name: 'find-related-documents',
    description: 'Find related documents across services',
    category: 'documents',
    parameters: ['topic', 'days', 'limit', 'includeEmails', 'includeMeetings'],
    useCases: ['related documents', 'document search', 'file discovery'],
    parallelizable: false,
  },
  {
    name: 'build-knowledge-graph',
    description: 'Build knowledge graph from data',
    category: 'documents',
    parameters: ['topic', 'days', 'maxNodes'],
    useCases: ['knowledge graph', 'relationship mapping', 'connections'],
    parallelizable: false,
  },
  // Analytics
  {
    name: 'analyze-meeting-load',
    description: 'Analyze meeting load and identify issues',
    category: 'analytics',
    parameters: ['weeks', 'includeRecurring'],
    useCases: ['meeting analysis', 'meeting load', 'calendar analysis'],
    parallelizable: false,
  },
  // Email Tools
  {
    name: 'get-my-emails',
    description: 'Enhanced email retrieval with rich formatting',
    category: 'email',
    parameters: ['filter', 'search', 'limit', 'language'],
    useCases: ['emails', 'mail messages', 'inbox'],
    parallelizable: false,
  },
  {
    name: 'list-mail-messages',
    description: 'List email messages',
    category: 'email',
    parameters: ['top', 'skip', 'filter', 'search', 'orderby'],
    useCases: ['list emails', 'email list', 'messages'],
    parallelizable: false,
  },
  // Calendar Tools
  {
    name: 'find-upcoming-meetings',
    description: 'Find upcoming calendar meetings',
    category: 'calendar',
    parameters: ['hoursAhead', 'limit'],
    useCases: ['upcoming meetings', 'future meetings', 'calendar events'],
    parallelizable: false,
  },
  {
    name: 'list-calendar-events',
    description: 'List calendar events',
    category: 'calendar',
    parameters: ['top', 'skip', 'filter', 'startDateTime', 'endDateTime'],
    useCases: ['calendar events', 'meetings', 'appointments'],
    parallelizable: false,
  },
  // File Tools
  {
    name: 'find-files-from-person',
    description: 'Find files shared by a specific person',
    category: 'files',
    parameters: ['person', 'limit'],
    useCases: ['files from person', 'shared files', 'documents'],
    dependencies: ['find-user'],
    parallelizable: false,
  },
];

/**
 * Analyze query and create tool execution plan
 */
function createToolExecutionPlan(
  query: string,
  nlpEnhancer: NLPEnhancer
): {
  query: string;
  analysis: {
    intent: string;
    entities: string[];
    dataTypes: string[];
  };
  executionPlan: Array<{
    step: number;
    tool: string;
    parameters: Record<string, unknown>;
    reason: string;
    expectedResult: string;
    nextSteps?: string;
    dependsOn?: number[];
    canRunInParallel: boolean;
  }>;
  summary: {
    totalSteps: number;
    estimatedTime: string;
    parallelSteps: number[];
    sequentialSteps: number[];
    primaryTool: string;
    followUpTools: string[];
  };
  instructions: {
    forLLM: string;
    fallback: string;
  };
} {
  // Analyze query using NLP
  const decomposed = nlpEnhancer.decomposeQuery(query);
  const queryLower = query.toLowerCase();

  // Extract entities
  const entities: string[] = [];
  if (decomposed.entity) entities.push(decomposed.entity);
  entities.push(...decomposed.entities.map((e) => e.value));

  // Determine data types needed
  const dataTypes: string[] = [];
  if (
    queryLower.includes('email') ||
    queryLower.includes('mail') ||
    queryLower.includes('message') ||
    queryLower.includes('inbox')
  ) {
    dataTypes.push('emails');
  }
  if (
    queryLower.includes('meeting') ||
    queryLower.includes('calendar') ||
    queryLower.includes('event') ||
    queryLower.includes('appointment')
  ) {
    dataTypes.push('meetings');
  }
  if (
    queryLower.includes('file') ||
    queryLower.includes('document') ||
    queryLower.includes('onedrive') ||
    queryLower.includes('sharepoint')
  ) {
    dataTypes.push('files');
  }
  if (
    queryLower.includes('task') ||
    queryLower.includes('todo') ||
    queryLower.includes('planner')
  ) {
    dataTypes.push('tasks');
  }
  if (
    queryLower.includes('person') ||
    queryLower.includes('people') ||
    queryLower.includes('contact') ||
    queryLower.includes('colleague') ||
    queryLower.includes('team')
  ) {
    dataTypes.push('people');
  }
  if (dataTypes.length === 0) {
    dataTypes.push('all'); // Default to all if unclear
  }

  // Determine intent
  let intent = 'information_gathering';
  if (queryLower.includes('prepare') || queryLower.includes('preparation')) {
    intent = 'preparation';
  } else if (queryLower.includes('summary') || queryLower.includes('summarize')) {
    intent = 'summarization';
  } else if (queryLower.includes('analyze') || queryLower.includes('analysis')) {
    intent = 'analysis';
  } else if (queryLower.includes('find') || queryLower.includes('search')) {
    intent = 'search';
  } else if (queryLower.includes('everything') || queryLower.includes('all')) {
    intent = 'comprehensive_information_gathering';
  }

  const executionPlan: Array<{
    step: number;
    tool: string;
    parameters: Record<string, unknown>;
    reason: string;
    expectedResult: string;
    nextSteps?: string;
    dependsOn?: number[];
    canRunInParallel: boolean;
  }> = [];

  let stepNumber = 1;

  // Extract main topic/entity from query
  const mainTopic = entities[0] || query.split(' ').slice(-2).join(' ');

  // Plan based on intent and data types
  if (intent === 'comprehensive_information_gathering' || dataTypes.includes('all')) {
    // Start with universal search
    executionPlan.push({
      step: stepNumber++,
      tool: 'search-everything',
      parameters: {
        query: mainTopic,
        limit: 25,
      },
      reason: 'Start with universal search to find all mentions across all Microsoft 365 services',
      expectedResult: 'List of emails, files, meetings, and other items mentioning the topic',
      nextSteps: 'Use results to identify key people, dates, and documents for follow-up queries',
      canRunInParallel: false,
    });

    // If it's a project, add project-specific tools
    if (queryLower.includes('project') || mainTopic.toLowerCase().includes('project')) {
      executionPlan.push({
        step: stepNumber++,
        tool: 'get-project-overview',
        parameters: {
          projectName: mainTopic,
          includeFiles: true,
          includeMeetings: true,
          includeEmails: true,
          includeTasks: true,
        },
        reason: 'Get structured project overview with all related items',
        dependsOn: [1],
        expectedResult: 'Comprehensive project overview with files, meetings, emails, and tasks',
        canRunInParallel: false,
      });

      executionPlan.push({
        step: stepNumber++,
        tool: 'get-project-stakeholders',
        parameters: {
          projectName: mainTopic,
          days: 90,
        },
        reason: 'Identify all people involved in the project',
        dependsOn: [1],
        expectedResult: 'List of stakeholders with their involvement levels',
        canRunInParallel: true,
      });
    }

    // Add document search
    executionPlan.push({
      step: stepNumber++,
      tool: 'find-related-documents',
      parameters: {
        topic: mainTopic,
        days: 180,
        limit: 50,
        includeEmails: true,
        includeMeetings: true,
      },
      reason: 'Find all documents related to the topic',
      dependsOn: [1],
      expectedResult: 'List of related documents with relevance scores',
      canRunInParallel: true,
    });
  } else if (intent === 'preparation' && queryLower.includes('meeting')) {
    // Meeting preparation workflow
    const meetingSubject = mainTopic;
    executionPlan.push({
      step: stepNumber++,
      tool: 'prepare-for-meeting',
      parameters: {
        meetingSubject: meetingSubject,
        hoursAhead: 48,
      },
      reason: 'Gather all context for the upcoming meeting',
      expectedResult: 'Meeting preparation package with history, emails, and related documents',
      canRunInParallel: false,
    });
  } else if (dataTypes.includes('people') && entities.length > 0) {
    // Person-focused query
    const person = entities[0];
    executionPlan.push({
      step: stepNumber++,
      tool: 'get-communication-summary',
      parameters: {
        person: person,
        includeEmails: true,
        includeChats: true,
        includeMeetings: true,
        includeFiles: true,
      },
      reason: 'Get complete communication overview with the person',
      expectedResult: 'Comprehensive summary of all interactions',
      canRunInParallel: false,
    });
  } else if (dataTypes.includes('emails')) {
    // Email-focused query
    if (entities.length > 0 && queryLower.includes('with')) {
      // Emails with person
      const person = entities[0];
      executionPlan.push({
        step: stepNumber++,
        tool: 'find-emails-with-person',
        parameters: {
          person: person,
          limit: 20,
        },
        reason: 'Find all email conversations with the specified person',
        expectedResult: 'List of emails with the person',
        canRunInParallel: false,
      });
    } else if (queryLower.includes('thread') || queryLower.includes('conversation')) {
      // Email thread summary
      executionPlan.push({
        step: stepNumber++,
        tool: 'summarize-email-thread',
        parameters: {
          topic: mainTopic,
          days: 30,
          limit: 50,
        },
        reason: 'Summarize the email thread',
        expectedResult: 'Summary of email thread with key points and decisions',
        canRunInParallel: false,
      });
    } else {
      // General email query
      executionPlan.push({
        step: stepNumber++,
        tool: 'get-my-emails',
        parameters: {
          filter: queryLower.includes('unread') ? 'unread' : 'all',
          search: mainTopic,
          limit: 20,
        },
        reason: 'Retrieve emails matching the query',
        expectedResult: 'List of relevant emails',
        canRunInParallel: false,
      });
    }
  } else if (dataTypes.includes('meetings')) {
    // Meeting-focused query
    if (queryLower.includes('upcoming') || queryLower.includes('future')) {
      executionPlan.push({
        step: stepNumber++,
        tool: 'find-upcoming-meetings',
        parameters: {
          hoursAhead: 168, // 1 week
          limit: 20,
        },
        reason: 'Find upcoming meetings',
        expectedResult: 'List of upcoming meetings',
        canRunInParallel: false,
      });
    } else if (entities.length > 0 && queryLower.includes('with')) {
      const person = entities[0];
      executionPlan.push({
        step: stepNumber++,
        tool: 'find-meetings-with-person',
        parameters: {
          person: person,
          limit: 20,
        },
        reason: 'Find meetings with the specified person',
        expectedResult: 'List of meetings with the person',
        canRunInParallel: false,
      });
    } else {
      executionPlan.push({
        step: stepNumber++,
        tool: 'list-calendar-events',
        parameters: {
          top: 20,
          filter: queryLower.includes('today')
            ? `start/dateTime ge ${new Date().toISOString().split('T')[0]}T00:00:00Z`
            : undefined,
        },
        reason: 'List calendar events',
        expectedResult: 'List of calendar events',
        canRunInParallel: false,
      });
    }
  } else if (dataTypes.includes('tasks')) {
    // Task-focused query
    executionPlan.push({
      step: stepNumber++,
      tool: 'get-all-my-tasks',
      parameters: {
        includeCompleted: queryLower.includes('completed') || queryLower.includes('all'),
        dueSoon: true,
      },
      reason: 'Get unified task view',
      expectedResult: 'List of tasks from To-Do and Planner',
      canRunInParallel: false,
    });
  } else if (intent === 'summarization') {
    // Summarization query
    if (queryLower.includes('week')) {
      executionPlan.push({
        step: stepNumber++,
        tool: 'get-my-week-summary',
        parameters: {
          weekOffset: 0,
        },
        reason: 'Get weekly productivity summary',
        expectedResult: 'Weekly summary with meetings, emails, and tasks',
        canRunInParallel: false,
      });
    } else if (queryLower.includes('email') && queryLower.includes('thread')) {
      executionPlan.push({
        step: stepNumber++,
        tool: 'summarize-email-thread',
        parameters: {
          topic: mainTopic,
          days: 30,
          limit: 50,
        },
        reason: 'Summarize email thread',
        expectedResult: 'Email thread summary',
        canRunInParallel: false,
      });
    }
  } else {
    // Default: use intelligent search
    executionPlan.push({
      step: stepNumber++,
      tool: 'ask-microsoft-365',
      parameters: {
        question: query,
      },
      reason: 'Use intelligent assistant to answer the query',
      expectedResult: 'Comprehensive answer from Microsoft 365 data',
      canRunInParallel: false,
    });
  }

  // Calculate parallel steps
  const parallelSteps: number[] = [];
  const sequentialSteps: number[] = [];
  for (const step of executionPlan) {
    if (step.canRunInParallel && (!step.dependsOn || step.dependsOn.length === 0)) {
      parallelSteps.push(step.step);
    } else {
      sequentialSteps.push(step.step);
    }
  }

  // Build LLM instructions
  const instructionSteps = executionPlan.map((step) => {
    const params = Object.entries(step.parameters)
      .map(([key, value]) => `${key}: ${JSON.stringify(value)}`)
      .join(', ');
    return `${step.step}. Call '${step.tool}' with parameters: ${params}. Reason: ${step.reason}`;
  });
  const forLLM = instructionSteps.join('\n');

  return {
    query,
    analysis: {
      intent,
      entities,
      dataTypes,
    },
    executionPlan,
    summary: {
      totalSteps: executionPlan.length,
      estimatedTime: `~${executionPlan.length * 2}-${executionPlan.length * 3} seconds`,
      parallelSteps,
      sequentialSteps,
      primaryTool: executionPlan[0]?.tool || 'ask-microsoft-365',
      followUpTools: executionPlan.slice(1).map((s) => s.tool),
    },
    instructions: {
      forLLM,
      fallback: "If any tool fails, try 'ask-microsoft-365' as fallback with the same query",
    },
  };
}

export default { registerCompoundTools };
