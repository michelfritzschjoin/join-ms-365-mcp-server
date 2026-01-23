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

    const response = await graphClient.makeRequest('/me/calendarView', {
      method: 'GET',
      queryParams: {
        startDateTime: pastDate.toISOString(),
        endDateTime: futureDate.toISOString(),
        $top: '100',
        $select: 'id,subject,bodyPreview,start,end,attendees,organizer,location,webLink',
        $orderby: 'start/dateTime desc',
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

    const events = response.value as GraphEvent[];
    const matchingEvents: GraphEvent[] = [];

    const emailLower = userEmail.toLowerCase();
    const nameLower = userDisplayName.toLowerCase();

    for (const event of events) {
      // Check organizer
      const organizerEmail = event.organizer?.emailAddress?.address?.toLowerCase();
      const organizerName = event.organizer?.emailAddress?.name?.toLowerCase();

      if (organizerEmail === emailLower || (organizerName && organizerName.includes(nameLower))) {
        matchingEvents.push(event);
        continue;
      }

      // Check attendees
      if (event.attendees && Array.isArray(event.attendees)) {
        const hasAttendee = event.attendees.some((attendee) => {
          const attendeeEmail = attendee.emailAddress?.address?.toLowerCase();
          const attendeeName = attendee.emailAddress?.name?.toLowerCase();
          return attendeeEmail === emailLower || (attendeeName && attendeeName.includes(nameLower));
        });

        if (hasAttendee) {
          matchingEvents.push(event);
        }
      }
    }

    return matchingEvents.slice(0, limit);
  } catch (error) {
    logger.error(`Error finding meetings: ${error}`);
    return [];
  }
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

  // 2. Search OneDrive/SharePoint for files mentioning the person
  try {
    const searchResponse = await graphClient.makeRequest('/search/query', {
      method: 'POST',
      body: JSON.stringify({
        requests: [
          {
            entityTypes: ['driveItem'],
            query: {
              queryString: `author:"${userDisplayName}" OR createdBy:"${userDisplayName}"`,
            },
            from: 0,
            size: limit,
          },
        ],
      }),
    });

    if (
      searchResponse &&
      typeof searchResponse === 'object' &&
      'value' in searchResponse &&
      Array.isArray(searchResponse.value)
    ) {
      const searchValues = searchResponse.value as Array<{
        hitsContainers?: Array<{
          hits?: Array<{
            resource?: GraphDriveItem;
          }>;
        }>;
      }>;

      for (const container of searchValues) {
        if (container.hitsContainers) {
          for (const hitsContainer of container.hitsContainers) {
            if (hitsContainer.hits) {
              for (const hit of hitsContainer.hits) {
                if (hit.resource) {
                  allFiles.push(hit.resource);
                }
              }
            }
          }
        }
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
 * Register all compound tools
 */
export function registerCompoundTools(
  server: McpServer,
  graphClient: GraphClient,
  readOnly: boolean = false
): number {
  let registeredCount = 0;

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
        content: msg.body?.content?.replace(/<[^>]*>/g, '').substring(0, 500), // Strip HTML
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
      const pastMeetings = meetings.filter((m) => new Date(m.start.dateTime) < now);
      const upcomingMeetings = meetings.filter((m) => new Date(m.start.dateTime) >= now);

      const formatMeeting = (event: GraphEvent) => ({
        id: event.id,
        subject: event.subject,
        start: event.start.dateTime,
        end: event.end.dateTime,
        location: event.location?.displayName,
        organizer: event.organizer?.emailAddress?.name,
        attendeeCount: event.attendees?.length || 0,
        webLink: event.webLink,
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

      const formattedFiles = files.map((file) => ({
        id: file.id,
        name: file.name,
        webUrl: file.webUrl,
        size: file.size,
        type: file.file?.mimeType || (file.folder ? 'folder' : 'unknown'),
        createdDateTime: file.createdDateTime,
        lastModifiedDateTime: file.lastModifiedDateTime,
        sharedBy: file.shared?.sharedBy?.user?.displayName,
        sharedDate: file.shared?.sharedDateTime,
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
                filesFound: formattedFiles.length,
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
                  content: m.body?.content?.replace(/<[^>]*>/g, '').substring(0, 200),
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
            summary.sharedFiles = {
              count: files.length,
              files: files.slice(0, 5).map((f) => ({
                name: f.name,
                webUrl: f.webUrl,
                sharedDate: f.shared?.sharedDateTime,
              })),
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
  // 6. SEARCH EVERYTHING ABOUT TOPIC - Cross-product search
  // ==========================================================================
  server.tool(
    'search-everything',
    `Search across ALL Microsoft 365 products for a topic. Combines results from:
- Emails (subject, body)
- Teams messages
- Files (OneDrive, SharePoint)
- Calendar events

Use this for broad searches like "Find everything about Project X" or "What do we have about the budget meeting?".`,
    {
      query: z.string().describe('Search query to find across all Microsoft 365'),
      limit: z.number().optional().describe('Maximum results per category (default: 10)'),
    },
    {
      title: 'Search Everything',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ query, limit = 10 }) => {
      logger.info(`Searching everything for: ${query}`);

      const results: Record<string, unknown> = {
        query,
        searchedAt: new Date().toISOString(),
      };

      // Search emails
      try {
        const emailResponse = await graphClient.makeRequest('/me/messages', {
          method: 'GET',
          queryParams: {
            $search: `"${query}"`,
            $top: String(limit),
            $select: 'id,subject,bodyPreview,receivedDateTime,from',
          },
        });

        if (
          emailResponse &&
          typeof emailResponse === 'object' &&
          'value' in emailResponse &&
          Array.isArray(emailResponse.value)
        ) {
          results.emails = {
            count: emailResponse.value.length,
            items: (emailResponse.value as GraphEmail[]).map((e) => ({
              subject: e.subject,
              from: e.from?.emailAddress?.address,
              date: e.receivedDateTime,
              preview: e.bodyPreview?.substring(0, 100),
            })),
          };
        }
      } catch (error) {
        results.emails = { error: `Search failed: ${error}` };
      }

      // Search files using Microsoft Search API
      try {
        const searchResponse = await graphClient.makeRequest('/search/query', {
          method: 'POST',
          body: JSON.stringify({
            requests: [
              {
                entityTypes: ['driveItem', 'listItem', 'site'],
                query: { queryString: query },
                from: 0,
                size: limit,
              },
            ],
          }),
        });

        if (searchResponse && typeof searchResponse === 'object' && 'value' in searchResponse) {
          const items: Array<{ name?: string; webUrl?: string; type?: string }> = [];
          const searchValues = searchResponse.value as Array<{
            hitsContainers?: Array<{
              hits?: Array<{
                resource?: { name?: string; webUrl?: string; '@odata.type'?: string };
              }>;
            }>;
          }>;

          for (const container of searchValues) {
            if (container.hitsContainers) {
              for (const hitsContainer of container.hitsContainers) {
                if (hitsContainer.hits) {
                  for (const hit of hitsContainer.hits) {
                    if (hit.resource) {
                      items.push({
                        name: hit.resource.name,
                        webUrl: hit.resource.webUrl,
                        type: hit.resource['@odata.type'],
                      });
                    }
                  }
                }
              }
            }
          }

          results.files = {
            count: items.length,
            items: items.slice(0, limit),
          };
        }
      } catch (error) {
        results.files = { error: `Search failed: ${error}` };
      }

      // Search calendar events
      try {
        const now = new Date();
        const pastDate = new Date(now.getTime() - 180 * 24 * 60 * 60 * 1000);
        const futureDate = new Date(now.getTime() + 90 * 24 * 60 * 60 * 1000);

        const calendarResponse = await graphClient.makeRequest('/me/calendarView', {
          method: 'GET',
          queryParams: {
            startDateTime: pastDate.toISOString(),
            endDateTime: futureDate.toISOString(),
            $filter: `contains(subject, '${query}')`,
            $top: String(limit),
            $select: 'id,subject,start,end,location',
          },
        });

        if (
          calendarResponse &&
          typeof calendarResponse === 'object' &&
          'value' in calendarResponse &&
          Array.isArray(calendarResponse.value)
        ) {
          results.events = {
            count: calendarResponse.value.length,
            items: (calendarResponse.value as GraphEvent[]).map((e) => ({
              subject: e.subject,
              start: e.start?.dateTime,
              location: e.location?.displayName,
            })),
          };
        }
      } catch (error) {
        results.events = { error: `Search failed: ${error}` };
      }

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(results, null, 2),
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
        const calendarResponse = await graphClient.makeRequest('/me/calendarView', {
          method: 'GET',
          queryParams: {
            startDateTime: now.toISOString(),
            endDateTime: futureDate.toISOString(),
            $top: '20',
            $select: 'id,subject,start,end,attendees,organizer,location,bodyPreview,webLink',
            $orderby: 'start/dateTime',
          },
        });

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
            const pastMeetingsResponse = await graphClient.makeRequest('/me/calendarView', {
              method: 'GET',
              queryParams: {
                startDateTime: pastDate.toISOString(),
                endDateTime: now.toISOString(),
                $top: '50',
                $select: 'id,subject,start,attendees',
                $orderby: 'start/dateTime desc',
              },
            });

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
            const meetingsResponse = await graphClient.makeRequest('/me/calendarView', {
              method: 'GET',
              queryParams: {
                startDateTime: weekStart.toISOString(),
                endDateTime: weekEnd.toISOString(),
                $top: '100',
                $select: 'id,subject,start,end,attendees,isOnlineMeeting',
              },
            });

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

      // Files
      if (includeFiles) {
        promises.push(
          (async () => {
            try {
              const searchResponse = await graphClient.makeRequest('/search/query', {
                method: 'POST',
                body: JSON.stringify({
                  requests: [
                    {
                      entityTypes: ['driveItem', 'listItem'],
                      query: { queryString: projectName },
                      from: 0,
                      size: 15,
                    },
                  ],
                }),
              });

              if (
                searchResponse &&
                typeof searchResponse === 'object' &&
                'value' in searchResponse
              ) {
                const items: Array<{ name?: string; webUrl?: string; lastModified?: string }> = [];
                const searchValues = searchResponse.value as Array<{
                  hitsContainers?: Array<{
                    hits?: Array<{
                      resource?: {
                        name?: string;
                        webUrl?: string;
                        lastModifiedDateTime?: string;
                      };
                    }>;
                  }>;
                }>;

                for (const container of searchValues) {
                  for (const hc of container.hitsContainers || []) {
                    for (const hit of hc.hits || []) {
                      if (hit.resource) {
                        items.push({
                          name: hit.resource.name,
                          webUrl: hit.resource.webUrl,
                          lastModified: hit.resource.lastModifiedDateTime,
                        });
                      }
                    }
                  }
                }

                result.files = { count: items.length, items };
              }
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

              const meetingsResponse = await graphClient.makeRequest('/me/calendarView', {
                method: 'GET',
                queryParams: {
                  startDateTime: pastDate.toISOString(),
                  endDateTime: futureDate.toISOString(),
                  $filter: `contains(subject, '${projectName}')`,
                  $top: '20',
                  $select: 'id,subject,start,end,organizer,attendees',
                  $orderby: 'start/dateTime desc',
                },
              });

              if (
                meetingsResponse &&
                typeof meetingsResponse === 'object' &&
                'value' in meetingsResponse &&
                Array.isArray(meetingsResponse.value)
              ) {
                const events = meetingsResponse.value as GraphEvent[];
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

              const meetingsResponse = await graphClient.makeRequest('/me/calendarView', {
                method: 'GET',
                queryParams: {
                  startDateTime: now.toISOString(),
                  endDateTime: futureDate.toISOString(),
                  $top: '50',
                  $select: 'id,subject,start,organizer,responseStatus,webLink',
                },
              });

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
        const calendarResponse = await graphClient.makeRequest('/me/calendarView', {
          method: 'GET',
          queryParams: {
            startDateTime: startDate.toISOString(),
            endDateTime: endDate.toISOString(),
            $filter: 'isOnlineMeeting eq true',
            $select: 'id,subject,start,end,attendees,isOnlineMeeting,onlineMeeting,organizer',
            $orderby: 'start/dateTime desc',
            $top: '50',
          },
        });

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

export default { registerCompoundTools };
