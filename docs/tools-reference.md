# MCP Tools Reference

> **Last Updated:** 2026-01-23  
> **Repository:** https://github.com/michelfritzschjoin/join-ms-365-mcp-server

## Overview

The Join Microsoft 365 MCP Server provides a comprehensive set of tools for interacting with Microsoft 365 services. All tools follow the Model Context Protocol (MCP) specification.

## Tool Categories

### Authentication Tools

| Tool             | Description                                        | Mode      |
| ---------------- | -------------------------------------------------- | --------- |
| `login`          | Authenticate with Microsoft using device code flow | HTTP only |
| `logout`         | Log out from Microsoft account                     | HTTP only |
| `verify-login`   | Check current authentication status                | All       |
| `list-accounts`  | List all cached Microsoft accounts                 | All       |
| `select-account` | Select a specific account to use                   | All       |
| `remove-account` | Remove an account from cache                       | All       |

### Capability Tools

| Tool                    | Description                                                                 | Mode |
| ----------------------- | --------------------------------------------------------------------------- | ---- |
| `get-example-questions` | Returns example questions the server can answer 100% (for "Was kannst du?") | All  |

Use `get-example-questions` when the user asks "What can you do?", "Was kannst du?", or "Welche Fragen kann ich dir stellen?". Optional parameter `language`: `"de"`, `"en"`, or `"both"` (default). The same data is exposed in `GET /capabilities` as `exampleQuestions`.

### User Tools

| Tool               | Description              | Mode     |
| ------------------ | ------------------------ | -------- |
| `get-current-user` | Get current user profile | All      |
| `list-users`       | List users in directory  | Org Mode |

### Mail Tools

| Tool                        | Description                | Mode |
| --------------------------- | -------------------------- | ---- |
| `list-mail-messages`        | List email messages        | All  |
| `get-mail-message`          | Get specific email details | All  |
| `send-mail`                 | Send an email              | All  |
| `create-draft-email`        | Create email draft         | All  |
| `delete-mail-message`       | Delete an email            | All  |
| `move-mail-message`         | Move email to folder       | All  |
| `list-mail-folders`         | List mail folders          | All  |
| `list-mail-folder-messages` | List messages in folder    | All  |
| `list-mail-attachments`     | List email attachments     | All  |
| `get-mail-attachment`       | Download attachment        | All  |
| `add-mail-attachment`       | Add attachment to draft    | All  |
| `delete-mail-attachment`    | Remove attachment          | All  |

### Calendar Tools

| Tool                             | Description                         | Mode |
| -------------------------------- | ----------------------------------- | ---- |
| `list-calendars`                 | List user calendars                 | All  |
| `list-calendar-events`           | List events                         | All  |
| `get-calendar-event`             | Get event details                   | All  |
| `create-calendar-event`          | Create new event                    | All  |
| `update-calendar-event`          | Update event                        | All  |
| `delete-calendar-event`          | Delete event                        | All  |
| `get-calendar-view`              | Get calendar view                   | All  |
| `list-specific-calendar-events`  | List events from specific calendar  | All  |
| `get-specific-calendar-event`    | Get event from specific calendar    | All  |
| `create-specific-calendar-event` | Create event in specific calendar   | All  |
| `update-specific-calendar-event` | Update event in specific calendar   | All  |
| `delete-specific-calendar-event` | Delete event from specific calendar | All  |
| `find-meeting-times`             | Find available meeting times        | All  |

### Contact Tools

| Tool                     | Description         | Mode |
| ------------------------ | ------------------- | ---- |
| `list-outlook-contacts`  | List contacts       | All  |
| `get-outlook-contact`    | Get contact details | All  |
| `create-outlook-contact` | Create new contact  | All  |
| `update-outlook-contact` | Update contact      | All  |
| `delete-outlook-contact` | Delete contact      | All  |

### File Tools (OneDrive)

| Tool                             | Description           | Mode |
| -------------------------------- | --------------------- | ---- |
| `list-drives`                    | List available drives | All  |
| `get-drive-root-item`            | Get drive root folder | All  |
| `list-folder-files`              | List files in folder  | All  |
| `download-onedrive-file-content` | Download file content | All  |
| `upload-file-content`            | Upload file           | All  |
| `delete-onedrive-file`           | Delete file           | All  |

When using **Super-Tools** (`files` tool with `action`): Use `get-content` to get readable text from Word, Excel, and PowerPoint files (after locating the file via search or list). Use `fetchAllPages: true` (default) to load all pages for large folders or SharePoint lists.

### Excel Tools

| Tool                    | Description         | Mode |
| ----------------------- | ------------------- | ---- |
| `list-excel-worksheets` | List worksheets     | All  |
| `get-excel-range`       | Get cell range data | All  |
| `format-excel-range`    | Format cells        | All  |
| `sort-excel-range`      | Sort data           | All  |
| `create-excel-chart`    | Create chart        | All  |

### Task Tools (To-Do)

| Tool                   | Description      | Mode |
| ---------------------- | ---------------- | ---- |
| `list-todo-task-lists` | List task lists  | All  |
| `list-todo-tasks`      | List tasks       | All  |
| `get-todo-task`        | Get task details | All  |
| `create-todo-task`     | Create task      | All  |
| `update-todo-task`     | Update task      | All  |
| `delete-todo-task`     | Delete task      | All  |

### Planner Tools

| Tool                          | Description         | Mode     |
| ----------------------------- | ------------------- | -------- |
| `list-planner-tasks`          | List Planner tasks  | All      |
| `get-planner-plan`            | Get plan details    | Org Mode |
| `list-plan-tasks`             | List tasks in plan  | Org Mode |
| `create-planner-task`         | Create Planner task | Org Mode |
| `get-planner-task`            | Get task details    | Org Mode |
| `update-planner-task`         | Update task         | Org Mode |
| `update-planner-task-details` | Update task details | Org Mode |

### OneNote Tools

| Tool                             | Description      | Mode |
| -------------------------------- | ---------------- | ---- |
| `list-onenote-notebooks`         | List notebooks   | All  |
| `list-onenote-notebook-sections` | List sections    | All  |
| `list-onenote-section-pages`     | List pages       | All  |
| `get-onenote-page-content`       | Get page content | All  |
| `create-onenote-page`            | Create page      | All  |

### Teams Tools

| Tool                        | Description          | Mode     |
| --------------------------- | -------------------- | -------- |
| `list-joined-teams`         | List joined teams    | Org Mode |
| `get-team`                  | Get team details     | Org Mode |
| `list-chats`                | List chats           | Org Mode |
| `get-chat`                  | Get chat details     | Org Mode |
| `list-chat-messages`        | List chat messages   | Org Mode |
| `get-chat-message`          | Get message details  | Org Mode |
| `send-chat-message`         | Send chat message    | Org Mode |
| `list-chat-message-replies` | List message replies | Org Mode |
| `reply-to-chat-message`     | Reply to message     | Org Mode |

### SharePoint Tools

| Tool                              | Description      | Mode     |
| --------------------------------- | ---------------- | -------- |
| `search-sharepoint-sites`         | Search sites     | Org Mode |
| `get-sharepoint-site`             | Get site details | Org Mode |
| `get-sharepoint-site-by-path`     | Get site by path | Org Mode |
| `list-sharepoint-site-drives`     | List site drives | Org Mode |
| `get-sharepoint-site-drive-by-id` | Get drive by ID  | Org Mode |
| `list-sharepoint-site-items`      | List site items  | Org Mode |
| `get-sharepoint-site-item`        | Get item details | Org Mode |
| `list-sharepoint-site-lists`      | List site lists  | Org Mode |
| `get-sharepoint-site-list`        | Get list details | Org Mode |
| `list-sharepoint-site-list-items` | List list items  | Org Mode |
| `get-sharepoint-site-list-item`   | Get list item    | Org Mode |
| `get-sharepoint-sites-delta`      | Get site changes | Org Mode |

When using **Super-Tools** (`sharepoint` tool): For large lists, list-items and site-items automatically follow `@odata.nextLink` and merge pages (up to `MS365_MCP_MAX_PAGES`). Use `fetchAllPages: false` to get only the first page.

### Search Tools

| Tool           | Description                  | Mode |
| -------------- | ---------------------------- | ---- |
| `search-query` | Unified Microsoft 365 search | All  |

### Compound Tools (Intelligent Multi-Step)

These tools automatically chain multiple API calls to answer complex contextual questions. They are designed for natural language queries and eliminate the need for the AI to manually orchestrate multiple tool calls.

#### Person-Focused Tools

| Tool                        | Description                                                       | Mode |
| --------------------------- | ----------------------------------------------------------------- | ---- |
| `find-messages-with-person` | Find all Teams chat messages with a specific person               | All  |
| `find-emails-with-person`   | Find all email conversations with a specific person               | All  |
| `find-meetings-with-person` | Find past and upcoming meetings with a specific person            | All  |
| `find-upcoming-meetings`    | Find upcoming calendar meetings/events using rolling time windows | All  |
| `find-files-from-person`    | Find files shared by or from a specific person                    | All  |
| `get-communication-summary` | Get complete communication overview with a person                 | All  |

#### Business & Productivity Tools

| Tool                   | Description                                                    | Mode |
| ---------------------- | -------------------------------------------------------------- | ---- |
| `search-everything`    | Search across all M365 products (mail, files, calendar, Teams) | All  |
| `prepare-for-meeting`  | Gather all context for an upcoming meeting (history, emails)   | All  |
| `get-my-week-summary`  | Weekly productivity digest (meetings, emails, tasks)           | All  |
| `get-all-my-tasks`     | Unified task view from To-Do and Planner                       | All  |
| `get-project-overview` | Complete project overview (files, meetings, emails, tasks)     | All  |
| `get-company-contacts` | Find all contacts and interactions with a company              | All  |
| `get-follow-up-items`  | Items needing attention (flagged emails, overdue tasks, etc.)  | All  |

#### Content Intelligence & Extraction Tools

| Tool                     | Description                                   | Mode |
| ------------------------ | --------------------------------------------- | ---- |
| `extract-action-items`   | Extract action items from emails and meetings | All  |
| `summarize-email-thread` | Summarize long email threads                  | All  |
| `extract-decisions`      | Extract decisions from communications         | All  |

#### Relationship Intelligence Tools

| Tool                            | Description                                 | Mode |
| ------------------------------- | ------------------------------------------- | ---- |
| `analyze-relationship-strength` | Analyze relationship strength with contacts | All  |
| `find-mutual-connections`       | Find mutual connections between people      | All  |
| `get-communication-frequency`   | Analyze communication frequency             | All  |

#### Document Intelligence Tools

| Tool                     | Description                            | Mode |
| ------------------------ | -------------------------------------- | ---- |
| `find-related-documents` | Find related documents across services | All  |
| `build-knowledge-graph`  | Build knowledge graph from data        | All  |

#### Tool Orchestration & Intelligence Tools

| Tool                       | Description                                                 | Mode |
| -------------------------- | ----------------------------------------------------------- | ---- |
| `plan-tool-execution`      | **THE ORCHESTRATOR** - Creates execution plan for ANY query | All  |
| `suggest-tool-sequence`    | Suggests optimal tool sequences for scenarios               | All  |
| `get-tool-recommendations` | AI-powered tool recommendations                             | All  |

#### Business Workflow Tools

| Tool                        | Description                           | Mode |
| --------------------------- | ------------------------------------- | ---- |
| `execute-business-workflow` | Execute predefined business workflows | All  |
| `create-custom-workflow`    | Create and save custom workflows      | All  |

#### Advanced Analytics & Insights Tools

| Tool                        | Description                                   | Mode |
| --------------------------- | --------------------------------------------- | ---- |
| `analyze-business-metrics`  | Analyze business metrics across Microsoft 365 | All  |
| `get-business-intelligence` | Comprehensive business intelligence dashboard | All  |
| `analyze-team-performance`  | Team performance analytics                    | All  |

#### Smart Automation Tools

| Tool                    | Description                                   | Mode |
| ----------------------- | --------------------------------------------- | ---- |
| `auto-categorize-items` | Automatically categorize emails, files, tasks | All  |
| `smart-reminder-system` | Intelligent reminder system                   | All  |
| `auto-summarize-period` | Automatically summarize a time period         | All  |

#### Advanced Search & Discovery Tools

| Tool                        | Description                                                                     | Mode      |
| --------------------------- | ------------------------------------------------------------------------------- | --------- |
| `intelligent-query-builder` | Build optimized queries automatically                                           | All       |
| `discover-related-topics`   | Discover related topics and connections                                         | All       |
| `get-query-recommendation`  | Get intent-based tool and entity-type recommendation for a question (Discovery) | Discovery |

When **Discovery Tools** are enabled (`MS365_MCP_ENABLE_DISCOVERY_TOOLS=true`), `get-query-recommendation` returns suggested tools, recommended entity types for search, optimized query, and query analysis markdown. Use it before calling search or other tools to align with NLP intent.

#### Collaboration Intelligence Tools

| Tool                                 | Description                           | Mode |
| ------------------------------------ | ------------------------------------- | ---- |
| `analyze-collaboration-patterns`     | Analyze how teams collaborate         | All  |
| `suggest-collaboration-improvements` | Suggest improvements to collaboration | All  |

#### Usage Examples for Compound Tools

**Find Messages with Person:**

Perfect for queries like "What were my last messages with Ricardo Rohland?"

```json
{
  "tool": "find-messages-with-person",
  "arguments": {
    "person": "Ricardo Rohland",
    "limit": 20
  }
}
```

**Get Communication Summary:**

Perfect for queries like "Tell me everything about my interactions with John Smith"

```json
{
  "tool": "get-communication-summary",
  "arguments": {
    "person": "John Smith",
    "includeEmails": true,
    "includeChats": true,
    "includeMeetings": true,
    "includeFiles": true
  }
}
```

**Search Everything:**

Perfect for queries like "Find everything about Project Apollo"

```json
{
  "tool": "search-everything",
  "arguments": {
    "query": "Project Apollo",
    "limit": 15
  }
}
```

**Prepare for Meeting:**

Perfect for "Prepare me for my meeting with the marketing team"

```json
{
  "tool": "prepare-for-meeting",
  "arguments": {
    "meetingSubject": "Marketing",
    "hoursAhead": 48
  }
}
```

**Get My Week Summary:**

Perfect for "What did I accomplish this week?" or "Summarize my productivity"

```json
{
  "tool": "get-my-week-summary",
  "arguments": {
    "weekOffset": 0
  }
}
```

**Get All My Tasks:**

Perfect for "What do I need to work on?" or "Show me all my tasks"

```json
{
  "tool": "get-all-my-tasks",
  "arguments": {
    "includeCompleted": false,
    "dueSoon": true
  }
}
```

**Get Project Overview:**

Perfect for "What's the status of Project Apollo?" or "Summarize the Q4 Budget project"

```json
{
  "tool": "get-project-overview",
  "arguments": {
    "projectName": "Q4 Budget",
    "includeFiles": true,
    "includeMeetings": true,
    "includeEmails": true,
    "includeTasks": true
  }
}
```

**Get Company Contacts:**

Perfect for "Who do we know at Microsoft?" or "Find all contacts from Acme Corp"

```json
{
  "tool": "get-company-contacts",
  "arguments": {
    "companyName": "Microsoft"
  }
}
```

**Get Follow-Up Items:**

Perfect for "What needs my attention?" or "Show me urgent items"

```json
{
  "tool": "get-follow-up-items",
  "arguments": {
    "includeEmails": true,
    "includeTasks": true,
    "includeMeetings": true
  }
}
```

**Extract Action Items:**

Perfect for "What action items do I have from recent emails?" or "Extract tasks from yesterday's meeting"

```json
{
  "tool": "extract-action-items",
  "arguments": {
    "source": "both",
    "days": 7,
    "limit": 50
  }
}
```

**Summarize Email Thread:**

Perfect for "Summarize the email thread about [topic]" or "What was decided in this email chain?"

```json
{
  "tool": "summarize-email-thread",
  "arguments": {
    "topic": "Project Apollo",
    "days": 30,
    "limit": 50
  }
}
```

**Extract Decisions:**

Perfect for "What decisions were made about [topic]?" or "Extract all decisions from last week"

```json
{
  "tool": "extract-decisions",
  "arguments": {
    "topic": "Budget planning",
    "days": 90,
    "source": "both",
    "limit": 50
  }
}
```

**Analyze Relationship Strength:**

Perfect for "How strong is my relationship with [person]?" or "Who do I communicate with most?"

```json
{
  "tool": "analyze-relationship-strength",
  "arguments": {
    "person": "John Smith",
    "days": 90,
    "limit": 20
  }
}
```

**Find Mutual Connections:**

Perfect for "Who do I know in common with [person]?" or "Find mutual connections for [person]"

```json
{
  "tool": "find-mutual-connections",
  "arguments": {
    "person": "Jane Doe",
    "days": 180,
    "limit": 20
  }
}
```

**Get Communication Frequency:**

Perfect for "Who do I email most often?" or "Show my communication frequency"

```json
{
  "tool": "get-communication-frequency",
  "arguments": {
    "days": 90,
    "limit": 30,
    "includeMeetings": true
  }
}
```

**Find Related Documents:**

Perfect for "Find documents related to [topic]" or "Show files related to this meeting"

```json
{
  "tool": "find-related-documents",
  "arguments": {
    "topic": "Project Apollo",
    "days": 180,
    "limit": 50,
    "includeEmails": true,
    "includeMeetings": true
  }
}
```

**Build Knowledge Graph:**

Perfect for "Build a knowledge graph for [topic]" or "Show connections for [project]"

```json
{
  "tool": "build-knowledge-graph",
  "arguments": {
    "topic": "Project Apollo",
    "days": 180,
    "maxNodes": 50
  }
}
```

**Plan Tool Execution (THE ORCHESTRATOR):**

**CRITICAL**: This is the most important tool for LLMs. Call this FIRST for any user query to get a detailed execution plan.

Perfect for ANY query: "What do I know about Project Apollo?" → Get step-by-step plan

```json
{
  "tool": "plan-tool-execution",
  "arguments": {
    "query": "Find everything about Project Apollo"
  }
}
```

**Suggest Tool Sequence:**

Perfect for "What's the best way to prepare for a meeting?" or "How do I get a complete customer overview?"

```json
{
  "tool": "suggest-tool-sequence",
  "arguments": {
    "scenario": "meeting_preparation",
    "context": "Marketing team meeting"
  }
}
```

**Get Tool Recommendations:**

Perfect for "What tools should I use for this query?" or "Recommend tools for analyzing [topic]"

```json
{
  "tool": "get-tool-recommendations",
  "arguments": {
    "query": "Find all emails about budget",
    "limit": 5
  }
}
```

**Execute Business Workflow:**

Perfect for "Execute customer onboarding workflow for [company]" or "Run meeting preparation workflow"

```json
{
  "tool": "execute-business-workflow",
  "arguments": {
    "workflow": "customer_onboarding",
    "context": "Acme Corp"
  }
}
```

**Analyze Business Metrics:**

Perfect for "Analyze my business metrics this quarter" or "Show team collaboration patterns"

```json
{
  "tool": "analyze-business-metrics",
  "arguments": {
    "period": "quarter",
    "includeCommunication": true,
    "includeProjects": true,
    "includeCollaboration": true
  }
}
```

**Get Business Intelligence:**

Perfect for "Show my business intelligence dashboard" or "What are the key trends this month?"

```json
{
  "tool": "get-business-intelligence",
  "arguments": {
    "period": "month",
    "compareWithPrevious": true
  }
}
```

**Auto Categorize Items:**

Perfect for "Categorize my recent emails" or "Auto-organize files by project"

```json
{
  "tool": "auto-categorize-items",
  "arguments": {
    "source": "emails",
    "days": 7,
    "limit": 50
  }
}
```

**Smart Reminder System:**

Perfect for "What should I follow up on?" or "Show me upcoming deadlines"

```json
{
  "tool": "smart-reminder-system",
  "arguments": {
    "action": "list",
    "days": 7
  }
}
```

**Auto Summarize Period:**

Perfect for "Summarize my week" or "What happened this month?"

```json
{
  "tool": "auto-summarize-period",
  "arguments": {
    "period": "week",
    "includeEmails": true,
    "includeMeetings": true,
    "includeTasks": true
  }
}
```

**Intelligent Query Builder:**

Perfect for "Build a query for [topic]" or "Optimize this search query"

```json
{
  "tool": "intelligent-query-builder",
  "arguments": {
    "query": "project budget",
    "expandTerms": true,
    "addFilters": true
  }
}
```

**Discover Related Topics:**

Perfect for "What topics are related to [topic]?" or "Discover connections for [project]"

```json
{
  "tool": "discover-related-topics",
  "arguments": {
    "topic": "Project Apollo",
    "days": 180,
    "limit": 20
  }
}
```

**Analyze Collaboration Patterns:**

Perfect for "How does our team collaborate?" or "Identify collaboration bottlenecks"

```json
{
  "tool": "analyze-collaboration-patterns",
  "arguments": {
    "days": 90,
    "includeMeetings": true,
    "includeEmails": true
  }
}
```

**Suggest Collaboration Improvements:**

Perfect for "How can we collaborate better?" or "Suggest improvements to our workflow"

```json
{
  "tool": "suggest-collaboration-improvements",
  "arguments": {
    "focusArea": "all",
    "days": 90
  }
}
```

## Common Parameters

### Pagination Parameters

| Parameter       | Type    | Description                   |
| --------------- | ------- | ----------------------------- |
| `top`           | number  | Maximum items to return       |
| `skip`          | number  | Items to skip                 |
| `fetchAllPages` | boolean | Automatically fetch all pages |

### Filter Parameters

| Parameter | Type     | Description                 |
| --------- | -------- | --------------------------- |
| `filter`  | string   | OData filter expression     |
| `search`  | string   | Search query                |
| `select`  | string[] | Properties to return        |
| `expand`  | string[] | Related entities to include |
| `orderby` | string[] | Sort order                  |

### Control Parameters

| Parameter         | Type    | Description                 |
| ----------------- | ------- | --------------------------- |
| `includeHeaders`  | boolean | Include response headers    |
| `excludeResponse` | boolean | Return only success/failure |
| `timezone`        | string  | IANA timezone for dates     |

## Tool Response Format

All tools return MCP-compliant responses:

```typescript
interface ToolResponse {
  content: Array<{
    type: 'text';
    text: string; // JSON-stringified result
  }>;
  isError?: boolean;
}
```

## Error Handling

Errors include descriptive messages for AI understanding:

```json
{
  "content": [
    {
      "type": "text",
      "text": "Authentication required. Please use the 'login' tool first."
    }
  ],
  "isError": true
}
```

## Usage Examples

### List Recent Emails

```json
{
  "tool": "list-mail-messages",
  "arguments": {
    "top": 10,
    "orderby": ["receivedDateTime desc"]
  }
}
```

### Search Emails

```json
{
  "tool": "list-mail-messages",
  "arguments": {
    "search": "\"from:john@example.com subject:meeting\""
  }
}
```

### Create Calendar Event

```json
{
  "tool": "create-calendar-event",
  "arguments": {
    "body": {
      "subject": "Team Meeting",
      "start": {
        "dateTime": "2026-01-24T14:00:00",
        "timeZone": "Europe/Berlin"
      },
      "end": {
        "dateTime": "2026-01-24T15:00:00",
        "timeZone": "Europe/Berlin"
      }
    }
  }
}
```

---

_For detailed parameter schemas, see the MCP tool definitions in the server._
