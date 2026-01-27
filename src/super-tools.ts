/**
 * Super-Tools: Consolidated tool interface for Microsoft 365 MCP Server
 *
 * Instead of 126+ individual tools, we provide 10 "Super-Tools" that group
 * related functionality together. Each tool accepts an `action` parameter
 * to specify the operation.
 *
 * Benefits:
 * - Easier for LLMs to choose the right tool
 * - Cleaner UI in MCP clients
 * - Reduced cognitive load
 * - Same underlying functionality
 */

import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import GraphClient from './graph-client.js';
import logger from './logger.js';
import { addThinkingToResponse, isThinkingEnabled } from './thinking-process.js';

// Common schemas
const paginationSchema = {
  top: z.number().optional().describe('Maximum number of items to return (default: 25)'),
  skip: z.number().optional().describe('Number of items to skip for pagination'),
};

const filterSchema = {
  filter: z.string().optional().describe('OData filter expression'),
  search: z.string().optional().describe('Search query string'),
  orderby: z.string().optional().describe('OData orderby expression'),
};

// ============================================================================
// 1. EMAIL SUPER-TOOL
// ============================================================================
const emailActions = z.enum([
  'list', // List messages from inbox or folder
  'get', // Get a specific message
  'folders', // List mail folders
  'attachments', // List/get attachments
  'search', // Search messages
]);

const emailSchema = z.object({
  action: emailActions.describe('The email operation to perform'),
  // Identifiers
  messageId: z.string().optional().describe('Message ID (required for get, attachments)'),
  folderId: z.string().optional().describe('Folder ID to list messages from'),
  attachmentId: z.string().optional().describe('Attachment ID (for getting specific attachment)'),
  // Filters
  ...filterSchema,
  ...paginationSchema,
});

type EmailInput = z.infer<typeof emailSchema>;

async function handleEmail(input: EmailInput, graphClient: GraphClient): Promise<string> {
  const thinking: string[] = [];

  switch (input.action) {
    case 'list': {
      thinking.push(
        `Listing emails${input.folderId ? ` from folder ${input.folderId}` : ' from inbox'}`
      );
      const endpoint = input.folderId
        ? `/me/mailFolders/${input.folderId}/messages`
        : '/me/messages';
      const params: Record<string, string> = { $top: String(input.top || 25) };
      if (input.filter) params.$filter = input.filter;
      if (input.search) params.$search = `"${input.search}"`;
      if (input.orderby) params.$orderby = input.orderby;
      if (input.skip) params.$skip = String(input.skip);

      const result = await graphClient.callEndpoint('GET', endpoint, params);
      return addThinkingToResponse(result, thinking);
    }

    case 'get': {
      if (!input.messageId) throw new Error('messageId is required for action "get"');
      thinking.push(`Getting email with ID: ${input.messageId}`);
      const result = await graphClient.callEndpoint('GET', `/me/messages/${input.messageId}`);
      return addThinkingToResponse(result, thinking);
    }

    case 'folders': {
      thinking.push('Listing mail folders');
      const params: Record<string, string> = { $top: String(input.top || 50) };
      const result = await graphClient.callEndpoint('GET', '/me/mailFolders', params);
      return addThinkingToResponse(result, thinking);
    }

    case 'attachments': {
      if (!input.messageId) throw new Error('messageId is required for action "attachments"');
      thinking.push(`Getting attachments for message: ${input.messageId}`);
      const endpoint = input.attachmentId
        ? `/me/messages/${input.messageId}/attachments/${input.attachmentId}`
        : `/me/messages/${input.messageId}/attachments`;
      const result = await graphClient.callEndpoint('GET', endpoint);
      return addThinkingToResponse(result, thinking);
    }

    case 'search': {
      if (!input.search) throw new Error('search query is required for action "search"');
      thinking.push(`Searching emails for: ${input.search}`);
      const params: Record<string, string> = {
        $search: `"${input.search}"`,
        $top: String(input.top || 25),
      };
      const result = await graphClient.callEndpoint('GET', '/me/messages', params);
      return addThinkingToResponse(result, thinking);
    }

    default:
      throw new Error(`Unknown email action: ${input.action}`);
  }
}

// ============================================================================
// 2. CALENDAR SUPER-TOOL
// ============================================================================
const calendarActions = z.enum([
  'list', // List events from primary calendar
  'get', // Get specific event
  'view', // Get calendar view (date range)
  'calendars', // List all calendars
  'specific-calendar', // List events from specific calendar
]);

const calendarSchema = z.object({
  action: calendarActions.describe('The calendar operation to perform'),
  // Identifiers
  eventId: z.string().optional().describe('Event ID (required for get)'),
  calendarId: z.string().optional().describe('Calendar ID (for specific-calendar action)'),
  // Date range for view
  startDateTime: z.string().optional().describe('Start date/time for calendar view (ISO format)'),
  endDateTime: z.string().optional().describe('End date/time for calendar view (ISO format)'),
  // Timezone
  timezone: z.string().optional().describe('Timezone for date/time values (e.g., "Europe/Berlin")'),
  // Filters
  ...filterSchema,
  ...paginationSchema,
});

type CalendarInput = z.infer<typeof calendarSchema>;

async function handleCalendar(input: CalendarInput, graphClient: GraphClient): Promise<string> {
  const thinking: string[] = [];
  const headers: Record<string, string> = {};
  if (input.timezone) {
    headers['Prefer'] = `outlook.timezone="${input.timezone}"`;
  }

  switch (input.action) {
    case 'list': {
      thinking.push('Listing calendar events');
      const params: Record<string, string> = { $top: String(input.top || 25) };
      if (input.filter) params.$filter = input.filter;
      if (input.orderby) params.$orderby = input.orderby;
      const result = await graphClient.callEndpoint(
        'GET',
        '/me/events',
        params,
        undefined,
        headers
      );
      return addThinkingToResponse(result, thinking);
    }

    case 'get': {
      if (!input.eventId) throw new Error('eventId is required for action "get"');
      thinking.push(`Getting event: ${input.eventId}`);
      const result = await graphClient.callEndpoint(
        'GET',
        `/me/events/${input.eventId}`,
        undefined,
        undefined,
        headers
      );
      return addThinkingToResponse(result, thinking);
    }

    case 'view': {
      if (!input.startDateTime || !input.endDateTime) {
        throw new Error('startDateTime and endDateTime are required for action "view"');
      }
      thinking.push(`Getting calendar view from ${input.startDateTime} to ${input.endDateTime}`);
      const params: Record<string, string> = {
        startDateTime: input.startDateTime,
        endDateTime: input.endDateTime,
        $top: String(input.top || 50),
      };
      if (input.orderby) params.$orderby = input.orderby;
      const result = await graphClient.callEndpoint(
        'GET',
        '/me/calendarView',
        params,
        undefined,
        headers
      );
      return addThinkingToResponse(result, thinking);
    }

    case 'calendars': {
      thinking.push('Listing all calendars');
      const result = await graphClient.callEndpoint('GET', '/me/calendars');
      return addThinkingToResponse(result, thinking);
    }

    case 'specific-calendar': {
      if (!input.calendarId)
        throw new Error('calendarId is required for action "specific-calendar"');
      thinking.push(`Listing events from calendar: ${input.calendarId}`);
      const params: Record<string, string> = { $top: String(input.top || 25) };
      if (input.filter) params.$filter = input.filter;
      const result = await graphClient.callEndpoint(
        'GET',
        `/me/calendars/${input.calendarId}/events`,
        params,
        undefined,
        headers
      );
      return addThinkingToResponse(result, thinking);
    }

    default:
      throw new Error(`Unknown calendar action: ${input.action}`);
  }
}

// ============================================================================
// 3. TEAMS SUPER-TOOL
// ============================================================================
const teamsActions = z.enum([
  'list-teams', // List joined teams
  'get-team', // Get team details
  'channels', // List team channels
  'channel-messages', // List channel messages
  'chats', // List chats
  'chat-messages', // List chat messages
]);

const teamsSchema = z.object({
  action: teamsActions.describe('The Teams operation to perform'),
  // Identifiers
  teamId: z.string().optional().describe('Team ID'),
  channelId: z.string().optional().describe('Channel ID'),
  chatId: z.string().optional().describe('Chat ID'),
  messageId: z.string().optional().describe('Message ID'),
  // Filters
  ...filterSchema,
  ...paginationSchema,
});

type TeamsInput = z.infer<typeof teamsSchema>;

async function handleTeams(input: TeamsInput, graphClient: GraphClient): Promise<string> {
  const thinking: string[] = [];

  switch (input.action) {
    case 'list-teams': {
      thinking.push('Listing joined teams');
      const result = await graphClient.callEndpoint('GET', '/me/joinedTeams');
      return addThinkingToResponse(result, thinking);
    }

    case 'get-team': {
      if (!input.teamId) throw new Error('teamId is required');
      thinking.push(`Getting team: ${input.teamId}`);
      const result = await graphClient.callEndpoint('GET', `/teams/${input.teamId}`);
      return addThinkingToResponse(result, thinking);
    }

    case 'channels': {
      if (!input.teamId) throw new Error('teamId is required for channels');
      thinking.push(`Listing channels for team: ${input.teamId}`);
      const result = await graphClient.callEndpoint('GET', `/teams/${input.teamId}/channels`);
      return addThinkingToResponse(result, thinking);
    }

    case 'channel-messages': {
      if (!input.teamId || !input.channelId) {
        throw new Error('teamId and channelId are required for channel-messages');
      }
      thinking.push(`Listing messages in channel: ${input.channelId}`);
      const params: Record<string, string> = { $top: String(input.top || 25) };
      const result = await graphClient.callEndpoint(
        'GET',
        `/teams/${input.teamId}/channels/${input.channelId}/messages`,
        params
      );
      return addThinkingToResponse(result, thinking);
    }

    case 'chats': {
      thinking.push('Listing chats');
      const params: Record<string, string> = { $top: String(input.top || 25) };
      const result = await graphClient.callEndpoint('GET', '/me/chats', params);
      return addThinkingToResponse(result, thinking);
    }

    case 'chat-messages': {
      if (!input.chatId) throw new Error('chatId is required for chat-messages');
      thinking.push(`Listing messages in chat: ${input.chatId}`);
      const params: Record<string, string> = { $top: String(input.top || 25) };
      const result = await graphClient.callEndpoint(
        'GET',
        `/me/chats/${input.chatId}/messages`,
        params
      );
      return addThinkingToResponse(result, thinking);
    }

    default:
      throw new Error(`Unknown teams action: ${input.action}`);
  }
}

// ============================================================================
// 4. FILES SUPER-TOOL
// ============================================================================
const filesActions = z.enum([
  'drives', // List drives
  'list', // List files in folder
  'get', // Get file metadata
  'download', // Download file content
  'root', // Get drive root
  'search', // Search files
]);

const filesSchema = z.object({
  action: filesActions.describe('The files operation to perform'),
  // Identifiers
  driveId: z.string().optional().describe('Drive ID'),
  itemId: z.string().optional().describe('Item (file/folder) ID'),
  path: z.string().optional().describe('Path to file/folder'),
  // Filters
  ...filterSchema,
  ...paginationSchema,
});

type FilesInput = z.infer<typeof filesSchema>;

async function handleFiles(input: FilesInput, graphClient: GraphClient): Promise<string> {
  const thinking: string[] = [];

  switch (input.action) {
    case 'drives': {
      thinking.push('Listing drives');
      const result = await graphClient.callEndpoint('GET', '/me/drives');
      return addThinkingToResponse(result, thinking);
    }

    case 'list': {
      const driveId = input.driveId || 'me';
      const itemId = input.itemId || 'root';
      thinking.push(`Listing files in ${driveId}/${itemId}`);
      const endpoint =
        driveId === 'me'
          ? `/me/drive/items/${itemId}/children`
          : `/drives/${driveId}/items/${itemId}/children`;
      const params: Record<string, string> = { $top: String(input.top || 50) };
      const result = await graphClient.callEndpoint('GET', endpoint, params);
      return addThinkingToResponse(result, thinking);
    }

    case 'get': {
      if (!input.itemId) throw new Error('itemId is required for get');
      const driveId = input.driveId || 'me';
      thinking.push(`Getting file metadata: ${input.itemId}`);
      const endpoint =
        driveId === 'me'
          ? `/me/drive/items/${input.itemId}`
          : `/drives/${driveId}/items/${input.itemId}`;
      const result = await graphClient.callEndpoint('GET', endpoint);
      return addThinkingToResponse(result, thinking);
    }

    case 'download': {
      if (!input.itemId) throw new Error('itemId is required for download');
      const driveId = input.driveId || 'me';
      thinking.push(`Downloading file: ${input.itemId}`);
      const endpoint =
        driveId === 'me'
          ? `/me/drive/items/${input.itemId}/content`
          : `/drives/${driveId}/items/${input.itemId}/content`;
      const result = await graphClient.callEndpoint('GET', endpoint);
      return addThinkingToResponse(result, thinking);
    }

    case 'root': {
      const driveId = input.driveId || 'me';
      thinking.push('Getting drive root');
      const endpoint = driveId === 'me' ? '/me/drive/root' : `/drives/${driveId}/root`;
      const result = await graphClient.callEndpoint('GET', endpoint);
      return addThinkingToResponse(result, thinking);
    }

    case 'search': {
      if (!input.search) throw new Error('search query is required');
      thinking.push(`Searching files for: ${input.search}`);
      const result = await graphClient.callEndpoint(
        'GET',
        `/me/drive/root/search(q='${input.search}')`
      );
      return addThinkingToResponse(result, thinking);
    }

    default:
      throw new Error(`Unknown files action: ${input.action}`);
  }
}

// ============================================================================
// 5. TASKS SUPER-TOOL
// ============================================================================
const tasksActions = z.enum([
  'todo-lists', // List To-Do task lists
  'todo-tasks', // List tasks in a To-Do list
  'todo-get', // Get specific To-Do task
  'planner-tasks', // List Planner tasks assigned to me
  'planner-plans', // Get Planner plan details
  'plan-tasks', // List tasks in a Planner plan
]);

const tasksSchema = z.object({
  action: tasksActions.describe('The tasks operation to perform'),
  // Identifiers
  taskListId: z.string().optional().describe('To-Do task list ID'),
  taskId: z.string().optional().describe('Task ID'),
  planId: z.string().optional().describe('Planner plan ID'),
  // Filters
  ...filterSchema,
  ...paginationSchema,
});

type TasksInput = z.infer<typeof tasksSchema>;

async function handleTasks(input: TasksInput, graphClient: GraphClient): Promise<string> {
  const thinking: string[] = [];

  switch (input.action) {
    case 'todo-lists': {
      thinking.push('Listing To-Do task lists');
      const result = await graphClient.callEndpoint('GET', '/me/todo/lists');
      return addThinkingToResponse(result, thinking);
    }

    case 'todo-tasks': {
      if (!input.taskListId) throw new Error('taskListId is required');
      thinking.push(`Listing tasks in list: ${input.taskListId}`);
      const params: Record<string, string> = { $top: String(input.top || 50) };
      if (input.filter) params.$filter = input.filter;
      const result = await graphClient.callEndpoint(
        'GET',
        `/me/todo/lists/${input.taskListId}/tasks`,
        params
      );
      return addThinkingToResponse(result, thinking);
    }

    case 'todo-get': {
      if (!input.taskListId || !input.taskId) {
        throw new Error('taskListId and taskId are required');
      }
      thinking.push(`Getting task: ${input.taskId}`);
      const result = await graphClient.callEndpoint(
        'GET',
        `/me/todo/lists/${input.taskListId}/tasks/${input.taskId}`
      );
      return addThinkingToResponse(result, thinking);
    }

    case 'planner-tasks': {
      thinking.push('Listing Planner tasks assigned to me');
      const result = await graphClient.callEndpoint('GET', '/me/planner/tasks');
      return addThinkingToResponse(result, thinking);
    }

    case 'planner-plans': {
      if (!input.planId) throw new Error('planId is required');
      thinking.push(`Getting Planner plan: ${input.planId}`);
      const result = await graphClient.callEndpoint('GET', `/planner/plans/${input.planId}`);
      return addThinkingToResponse(result, thinking);
    }

    case 'plan-tasks': {
      if (!input.planId) throw new Error('planId is required');
      thinking.push(`Listing tasks in plan: ${input.planId}`);
      const result = await graphClient.callEndpoint('GET', `/planner/plans/${input.planId}/tasks`);
      return addThinkingToResponse(result, thinking);
    }

    default:
      throw new Error(`Unknown tasks action: ${input.action}`);
  }
}

// ============================================================================
// 6. CONTACTS SUPER-TOOL
// ============================================================================
const contactsActions = z.enum([
  'list', // List contacts
  'get', // Get specific contact
  'users', // List organization users
  'current-user', // Get current user info
  'search', // Search contacts/users
]);

const contactsSchema = z.object({
  action: contactsActions.describe('The contacts operation to perform'),
  // Identifiers
  contactId: z.string().optional().describe('Contact ID'),
  userId: z.string().optional().describe('User ID'),
  // Filters
  ...filterSchema,
  ...paginationSchema,
});

type ContactsInput = z.infer<typeof contactsSchema>;

async function handleContacts(input: ContactsInput, graphClient: GraphClient): Promise<string> {
  const thinking: string[] = [];

  switch (input.action) {
    case 'list': {
      thinking.push('Listing contacts');
      const params: Record<string, string> = { $top: String(input.top || 50) };
      if (input.filter) params.$filter = input.filter;
      if (input.search) params.$search = `"${input.search}"`;
      const result = await graphClient.callEndpoint('GET', '/me/contacts', params);
      return addThinkingToResponse(result, thinking);
    }

    case 'get': {
      if (!input.contactId) throw new Error('contactId is required');
      thinking.push(`Getting contact: ${input.contactId}`);
      const result = await graphClient.callEndpoint('GET', `/me/contacts/${input.contactId}`);
      return addThinkingToResponse(result, thinking);
    }

    case 'users': {
      thinking.push('Listing organization users');
      const params: Record<string, string> = { $top: String(input.top || 50) };
      if (input.filter) params.$filter = input.filter;
      if (input.search) params.$search = `"${input.search}"`;
      const result = await graphClient.callEndpoint('GET', '/users', params, undefined, {
        ConsistencyLevel: 'eventual',
      });
      return addThinkingToResponse(result, thinking);
    }

    case 'current-user': {
      thinking.push('Getting current user info');
      const result = await graphClient.callEndpoint('GET', '/me');
      return addThinkingToResponse(result, thinking);
    }

    case 'search': {
      if (!input.search) throw new Error('search query is required');
      thinking.push(`Searching for: ${input.search}`);
      const params: Record<string, string> = {
        $search: `"${input.search}"`,
        $top: String(input.top || 25),
      };
      const result = await graphClient.callEndpoint('GET', '/users', params, undefined, {
        ConsistencyLevel: 'eventual',
      });
      return addThinkingToResponse(result, thinking);
    }

    default:
      throw new Error(`Unknown contacts action: ${input.action}`);
  }
}

// ============================================================================
// 7. MEETINGS SUPER-TOOL
// ============================================================================
const meetingsActions = z.enum([
  'list', // List online meetings
  'get', // Get meeting details
  'recordings', // List/get recordings
  'transcripts', // List/get transcripts
  'transcript-content', // Get transcript content
]);

const meetingsSchema = z.object({
  action: meetingsActions.describe('The meetings operation to perform'),
  // Identifiers
  meetingId: z.string().optional().describe('Online meeting ID'),
  recordingId: z.string().optional().describe('Recording ID'),
  transcriptId: z.string().optional().describe('Transcript ID'),
  // Filters
  ...filterSchema,
  ...paginationSchema,
});

type MeetingsInput = z.infer<typeof meetingsSchema>;

async function handleMeetings(input: MeetingsInput, graphClient: GraphClient): Promise<string> {
  const thinking: string[] = [];

  switch (input.action) {
    case 'list': {
      thinking.push('Listing online meetings');
      const params: Record<string, string> = { $top: String(input.top || 25) };
      if (input.filter) params.$filter = input.filter;
      const result = await graphClient.callEndpoint('GET', '/me/onlineMeetings', params);
      return addThinkingToResponse(result, thinking);
    }

    case 'get': {
      if (!input.meetingId) throw new Error('meetingId is required');
      thinking.push(`Getting meeting: ${input.meetingId}`);
      const result = await graphClient.callEndpoint('GET', `/me/onlineMeetings/${input.meetingId}`);
      return addThinkingToResponse(result, thinking);
    }

    case 'recordings': {
      if (!input.meetingId) throw new Error('meetingId is required');
      thinking.push(`Getting recordings for meeting: ${input.meetingId}`);
      const endpoint = input.recordingId
        ? `/me/onlineMeetings/${input.meetingId}/recordings/${input.recordingId}`
        : `/me/onlineMeetings/${input.meetingId}/recordings`;
      const result = await graphClient.callEndpoint('GET', endpoint);
      return addThinkingToResponse(result, thinking);
    }

    case 'transcripts': {
      if (!input.meetingId) throw new Error('meetingId is required');
      thinking.push(`Getting transcripts for meeting: ${input.meetingId}`);
      const endpoint = input.transcriptId
        ? `/me/onlineMeetings/${input.meetingId}/transcripts/${input.transcriptId}`
        : `/me/onlineMeetings/${input.meetingId}/transcripts`;
      const result = await graphClient.callEndpoint('GET', endpoint);
      return addThinkingToResponse(result, thinking);
    }

    case 'transcript-content': {
      if (!input.meetingId || !input.transcriptId) {
        throw new Error('meetingId and transcriptId are required');
      }
      thinking.push(`Getting transcript content: ${input.transcriptId}`);
      const result = await graphClient.callEndpoint(
        'GET',
        `/me/onlineMeetings/${input.meetingId}/transcripts/${input.transcriptId}/content`
      );
      return addThinkingToResponse(result, thinking);
    }

    default:
      throw new Error(`Unknown meetings action: ${input.action}`);
  }
}

// ============================================================================
// 8. SHAREPOINT SUPER-TOOL
// ============================================================================
const sharepointActions = z.enum([
  'search-sites', // Search SharePoint sites
  'get-site', // Get site details
  'site-drives', // List site drives
  'site-lists', // List site lists
  'list-items', // List items in a list
  'site-items', // List items in a site
]);

const sharepointSchema = z.object({
  action: sharepointActions.describe('The SharePoint operation to perform'),
  // Identifiers
  siteId: z.string().optional().describe('Site ID'),
  driveId: z.string().optional().describe('Drive ID'),
  listId: z.string().optional().describe('List ID'),
  itemId: z.string().optional().describe('Item ID'),
  // Filters
  ...filterSchema,
  ...paginationSchema,
});

type SharePointInput = z.infer<typeof sharepointSchema>;

async function handleSharePoint(input: SharePointInput, graphClient: GraphClient): Promise<string> {
  const thinking: string[] = [];

  switch (input.action) {
    case 'search-sites': {
      thinking.push('Searching SharePoint sites');
      const params: Record<string, string> = { $top: String(input.top || 25) };
      if (input.search) params.search = input.search;
      const result = await graphClient.callEndpoint('GET', '/sites', params);
      return addThinkingToResponse(result, thinking);
    }

    case 'get-site': {
      if (!input.siteId) throw new Error('siteId is required');
      thinking.push(`Getting site: ${input.siteId}`);
      const result = await graphClient.callEndpoint('GET', `/sites/${input.siteId}`);
      return addThinkingToResponse(result, thinking);
    }

    case 'site-drives': {
      if (!input.siteId) throw new Error('siteId is required');
      thinking.push(`Listing drives for site: ${input.siteId}`);
      const result = await graphClient.callEndpoint('GET', `/sites/${input.siteId}/drives`);
      return addThinkingToResponse(result, thinking);
    }

    case 'site-lists': {
      if (!input.siteId) throw new Error('siteId is required');
      thinking.push(`Listing lists for site: ${input.siteId}`);
      const result = await graphClient.callEndpoint('GET', `/sites/${input.siteId}/lists`);
      return addThinkingToResponse(result, thinking);
    }

    case 'list-items': {
      if (!input.siteId || !input.listId) {
        throw new Error('siteId and listId are required');
      }
      thinking.push(`Listing items in list: ${input.listId}`);
      const params: Record<string, string> = { $top: String(input.top || 50) };
      if (input.filter) params.$filter = input.filter;
      const result = await graphClient.callEndpoint(
        'GET',
        `/sites/${input.siteId}/lists/${input.listId}/items`,
        params
      );
      return addThinkingToResponse(result, thinking);
    }

    case 'site-items': {
      if (!input.siteId) throw new Error('siteId is required');
      thinking.push(`Listing items in site: ${input.siteId}`);
      const params: Record<string, string> = { $top: String(input.top || 50) };
      const result = await graphClient.callEndpoint('GET', `/sites/${input.siteId}/items`, params);
      return addThinkingToResponse(result, thinking);
    }

    default:
      throw new Error(`Unknown sharepoint action: ${input.action}`);
  }
}

// ============================================================================
// 9. NOTES SUPER-TOOL (OneNote)
// ============================================================================
const notesActions = z.enum([
  'notebooks', // List notebooks
  'sections', // List sections in notebook
  'pages', // List pages in section
  'page-content', // Get page content
]);

const notesSchema = z.object({
  action: notesActions.describe('The OneNote operation to perform'),
  // Identifiers
  notebookId: z.string().optional().describe('Notebook ID'),
  sectionId: z.string().optional().describe('Section ID'),
  pageId: z.string().optional().describe('Page ID'),
  // Filters
  ...filterSchema,
  ...paginationSchema,
});

type NotesInput = z.infer<typeof notesSchema>;

async function handleNotes(input: NotesInput, graphClient: GraphClient): Promise<string> {
  const thinking: string[] = [];

  switch (input.action) {
    case 'notebooks': {
      thinking.push('Listing OneNote notebooks');
      const result = await graphClient.callEndpoint('GET', '/me/onenote/notebooks');
      return addThinkingToResponse(result, thinking);
    }

    case 'sections': {
      if (!input.notebookId) throw new Error('notebookId is required');
      thinking.push(`Listing sections in notebook: ${input.notebookId}`);
      const result = await graphClient.callEndpoint(
        'GET',
        `/me/onenote/notebooks/${input.notebookId}/sections`
      );
      return addThinkingToResponse(result, thinking);
    }

    case 'pages': {
      if (!input.sectionId) throw new Error('sectionId is required');
      thinking.push(`Listing pages in section: ${input.sectionId}`);
      const params: Record<string, string> = { $top: String(input.top || 50) };
      const result = await graphClient.callEndpoint(
        'GET',
        `/me/onenote/sections/${input.sectionId}/pages`,
        params
      );
      return addThinkingToResponse(result, thinking);
    }

    case 'page-content': {
      if (!input.pageId) throw new Error('pageId is required');
      thinking.push(`Getting page content: ${input.pageId}`);
      const result = await graphClient.callEndpoint(
        'GET',
        `/me/onenote/pages/${input.pageId}/content`
      );
      return addThinkingToResponse(result, thinking);
    }

    default:
      throw new Error(`Unknown notes action: ${input.action}`);
  }
}

// ============================================================================
// 10. ASSISTANT SUPER-TOOL (Smart/Compound Operations)
// ============================================================================
const assistantActions = z.enum([
  'ask', // Natural language question about M365 data
  'search', // Search across all M365 data
  'my-day', // Get today's summary
  'my-week', // Get week summary
  'person-info', // Get all info about a person
  'project-overview', // Get project overview
  'follow-ups', // Get pending follow-up items
  'meeting-prep', // Prepare for upcoming meeting
]);

const assistantSchema = z.object({
  action: assistantActions.describe('The assistant operation to perform'),
  // Query
  query: z.string().optional().describe('Natural language query or search term'),
  // Person context
  person: z.string().optional().describe('Person name or email'),
  // Project/topic context
  topic: z.string().optional().describe('Topic or project name'),
  // Time context
  days: z.number().optional().describe('Number of days to look back (default: 7)'),
  // Limits
  limit: z.number().optional().describe('Maximum results to return (default: 25)'),
});

type AssistantInput = z.infer<typeof assistantSchema>;

async function handleAssistant(input: AssistantInput, graphClient: GraphClient): Promise<string> {
  const thinking: string[] = [];
  const results: Record<string, unknown> = {};
  const limit = input.limit || 25;
  const days = input.days || 7;

  switch (input.action) {
    case 'ask': {
      if (!input.query) throw new Error('query is required for ask action');
      thinking.push(`Processing question: ${input.query}`);
      thinking.push('Searching across emails, calendar, files, and chats...');

      // Search emails
      const emailResult = await graphClient.callEndpoint('GET', '/me/messages', {
        $search: `"${input.query}"`,
        $top: String(Math.min(limit, 10)),
      });
      results.emails = JSON.parse(emailResult);

      // Search files
      const filesResult = await graphClient.callEndpoint(
        'GET',
        `/me/drive/root/search(q='${input.query}')`
      );
      results.files = JSON.parse(filesResult);

      return addThinkingToResponse(JSON.stringify(results, null, 2), thinking);
    }

    case 'search': {
      if (!input.query) throw new Error('query is required for search action');
      thinking.push(`Searching everything for: ${input.query}`);

      const emailResult = await graphClient.callEndpoint('GET', '/me/messages', {
        $search: `"${input.query}"`,
        $top: String(limit),
      });
      results.emails = JSON.parse(emailResult);

      const filesResult = await graphClient.callEndpoint(
        'GET',
        `/me/drive/root/search(q='${input.query}')`
      );
      results.files = JSON.parse(filesResult);

      return addThinkingToResponse(JSON.stringify(results, null, 2), thinking);
    }

    case 'my-day': {
      thinking.push("Getting today's summary");
      const today = new Date();
      const tomorrow = new Date(today);
      tomorrow.setDate(tomorrow.getDate() + 1);

      const eventsResult = await graphClient.callEndpoint('GET', '/me/calendarView', {
        startDateTime: today.toISOString(),
        endDateTime: tomorrow.toISOString(),
      });
      results.todayEvents = JSON.parse(eventsResult);

      const emailsResult = await graphClient.callEndpoint('GET', '/me/messages', {
        $top: '10',
        $orderby: 'receivedDateTime desc',
      });
      results.recentEmails = JSON.parse(emailsResult);

      return addThinkingToResponse(JSON.stringify(results, null, 2), thinking);
    }

    case 'my-week': {
      thinking.push('Getting week summary');
      const today = new Date();
      const weekEnd = new Date(today);
      weekEnd.setDate(weekEnd.getDate() + 7);

      const eventsResult = await graphClient.callEndpoint('GET', '/me/calendarView', {
        startDateTime: today.toISOString(),
        endDateTime: weekEnd.toISOString(),
      });
      results.weekEvents = JSON.parse(eventsResult);

      const tasksResult = await graphClient.callEndpoint('GET', '/me/planner/tasks');
      results.tasks = JSON.parse(tasksResult);

      return addThinkingToResponse(JSON.stringify(results, null, 2), thinking);
    }

    case 'person-info': {
      if (!input.person) throw new Error('person is required');
      thinking.push(`Getting all info about: ${input.person}`);

      const emailsResult = await graphClient.callEndpoint('GET', '/me/messages', {
        $filter: `from/emailAddress/address eq '${input.person}' or contains(from/emailAddress/name, '${input.person}')`,
        $top: String(limit),
      });
      results.emails = JSON.parse(emailsResult);

      return addThinkingToResponse(JSON.stringify(results, null, 2), thinking);
    }

    case 'project-overview': {
      if (!input.topic) throw new Error('topic is required');
      thinking.push(`Getting project overview for: ${input.topic}`);

      const emailsResult = await graphClient.callEndpoint('GET', '/me/messages', {
        $search: `"${input.topic}"`,
        $top: String(limit),
      });
      results.emails = JSON.parse(emailsResult);

      const filesResult = await graphClient.callEndpoint(
        'GET',
        `/me/drive/root/search(q='${input.topic}')`
      );
      results.files = JSON.parse(filesResult);

      return addThinkingToResponse(JSON.stringify(results, null, 2), thinking);
    }

    case 'follow-ups': {
      thinking.push('Getting pending follow-up items');

      const flaggedEmails = await graphClient.callEndpoint('GET', '/me/messages', {
        $filter: "flag/flagStatus eq 'flagged'",
        $top: String(limit),
      });
      results.flaggedEmails = JSON.parse(flaggedEmails);

      const tasks = await graphClient.callEndpoint('GET', '/me/planner/tasks');
      results.tasks = JSON.parse(tasks);

      return addThinkingToResponse(JSON.stringify(results, null, 2), thinking);
    }

    case 'meeting-prep': {
      thinking.push('Preparing for upcoming meetings');
      const today = new Date();
      const tomorrow = new Date(today);
      tomorrow.setDate(tomorrow.getDate() + 1);

      const eventsResult = await graphClient.callEndpoint('GET', '/me/calendarView', {
        startDateTime: today.toISOString(),
        endDateTime: tomorrow.toISOString(),
      });
      results.upcomingMeetings = JSON.parse(eventsResult);

      return addThinkingToResponse(JSON.stringify(results, null, 2), thinking);
    }

    default:
      throw new Error(`Unknown assistant action: ${input.action}`);
  }
}

// ============================================================================
// REGISTRATION FUNCTION
// ============================================================================
export function registerSuperTools(server: McpServer, graphClient: GraphClient): void {
  logger.info('Registering Super-Tools (consolidated interface)');

  // 1. Email
  server.tool(
    'email',
    'Unified email operations: list messages, get message, folders, attachments, search',
    emailSchema.shape,
    async (input: EmailInput) => {
      try {
        const result = await handleEmail(input, graphClient);
        return { content: [{ type: 'text' as const, text: result }] };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // 2. Calendar
  server.tool(
    'calendar',
    'Unified calendar operations: list events, get event, calendar view, list calendars',
    calendarSchema.shape,
    async (input: CalendarInput) => {
      try {
        const result = await handleCalendar(input, graphClient);
        return { content: [{ type: 'text' as const, text: result }] };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // 3. Teams
  server.tool(
    'teams',
    'Unified Teams operations: teams, channels, chats, messages',
    teamsSchema.shape,
    async (input: TeamsInput) => {
      try {
        const result = await handleTeams(input, graphClient);
        return { content: [{ type: 'text' as const, text: result }] };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // 4. Files
  server.tool(
    'files',
    'Unified file operations: drives, list files, get file, download, search',
    filesSchema.shape,
    async (input: FilesInput) => {
      try {
        const result = await handleFiles(input, graphClient);
        return { content: [{ type: 'text' as const, text: result }] };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // 5. Tasks
  server.tool(
    'tasks',
    'Unified task operations: To-Do lists/tasks, Planner plans/tasks',
    tasksSchema.shape,
    async (input: TasksInput) => {
      try {
        const result = await handleTasks(input, graphClient);
        return { content: [{ type: 'text' as const, text: result }] };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // 6. Contacts
  server.tool(
    'contacts',
    'Unified contact operations: contacts, users, current user, search',
    contactsSchema.shape,
    async (input: ContactsInput) => {
      try {
        const result = await handleContacts(input, graphClient);
        return { content: [{ type: 'text' as const, text: result }] };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // 7. Meetings
  server.tool(
    'meetings',
    'Unified meeting operations: online meetings, recordings, transcripts',
    meetingsSchema.shape,
    async (input: MeetingsInput) => {
      try {
        const result = await handleMeetings(input, graphClient);
        return { content: [{ type: 'text' as const, text: result }] };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // 8. SharePoint
  server.tool(
    'sharepoint',
    'Unified SharePoint operations: sites, drives, lists, items',
    sharepointSchema.shape,
    async (input: SharePointInput) => {
      try {
        const result = await handleSharePoint(input, graphClient);
        return { content: [{ type: 'text' as const, text: result }] };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // 9. Notes (OneNote)
  server.tool(
    'notes',
    'Unified OneNote operations: notebooks, sections, pages, content',
    notesSchema.shape,
    async (input: NotesInput) => {
      try {
        const result = await handleNotes(input, graphClient);
        return { content: [{ type: 'text' as const, text: result }] };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  // 10. Assistant
  server.tool(
    'assistant',
    'Smart assistant: natural language queries, search everything, daily/weekly summaries, person info, project overview, follow-ups, meeting prep',
    assistantSchema.shape,
    async (input: AssistantInput) => {
      try {
        const result = await handleAssistant(input, graphClient);
        return { content: [{ type: 'text' as const, text: result }] };
      } catch (error) {
        return {
          content: [
            {
              type: 'text' as const,
              text: `Error: ${error instanceof Error ? error.message : String(error)}`,
            },
          ],
          isError: true,
        };
      }
    }
  );

  logger.info('Registered 10 Super-Tools');
}
