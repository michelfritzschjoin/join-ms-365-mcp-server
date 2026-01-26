<p align="center">
  <img src="https://img.shields.io/badge/Microsoft%20365-0078D4?style=for-the-badge&logo=microsoft&logoColor=white" alt="Microsoft 365">
  <img src="https://img.shields.io/badge/MCP%20Protocol-00A9CE?style=for-the-badge&logo=data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCI+PHBhdGggZmlsbD0id2hpdGUiIGQ9Ik0xMiAyTDIgN2wxMCA1IDEwLTV6Ii8+PC9zdmc+" alt="MCP">
  <img src="https://img.shields.io/badge/TypeScript-3178C6?style=for-the-badge&logo=typescript&logoColor=white" alt="TypeScript">
</p>

<h1 align="center">🚀 Join Microsoft 365 MCP Server</h1>

<p align="center">
  <strong>The Ultimate AI-Powered Gateway to Microsoft 365</strong>
</p>

<p align="center">
  A powerful Model Context Protocol (MCP) server enabling AI assistants to seamlessly interact with Microsoft 365 services through the Graph API.
</p>

<p align="center">
  <a href="#-quick-start">Quick Start</a> •
  <a href="#-features">Features</a> •
  <a href="#-tool-categories">Tools</a> •
  <a href="#-intelligent-discovery">Discovery</a> •
  <a href="#-configuration">Configuration</a> •
  <a href="#-security">Security</a>
</p>

---

## 📋 Table of Contents

- [Overview](#-overview)
- [Quick Start](#-quick-start)
- [Features](#-features)
- [Tool Categories](#-tool-categories)
  - [🔐 Authentication Tools](#-authentication-tools)
  - [📧 Email & Communication Tools](#-email--communication-tools)
  - [📅 Calendar Tools](#-calendar-tools)
  - [📁 File & Drive Tools](#-file--drive-tools)
  - [📊 Excel Tools](#-excel-tools)
  - [✅ Task & Planner Tools](#-task--planner-tools)
  - [📔 OneNote Tools](#-onenote-tools)
  - [💬 Microsoft Teams Tools](#-microsoft-teams-tools)
  - [🌐 SharePoint Tools](#-sharepoint-tools)
  - [👥 Contact Tools](#-contact-tools)
  - [🔍 Search & Discovery Tools](#-search--discovery-tools)
  - [🧠 Intelligent Compound Tools](#-intelligent-compound-tools)
  - [🎯 Business Intelligence Tools](#-business-intelligence-tools)
  - [🤖 AI Learning System Tools](#-ai-learning-system-tools)
- [Intelligent Discovery System](#-intelligent-discovery-system)
- [Configuration](#-configuration)
- [Authentication Methods](#-authentication-methods)
- [Cloud Environments](#-cloud-environments)
- [Security & Compliance](#-security--compliance)
- [API Reference](#-api-reference)

---

## 🌟 Overview

The **Join Microsoft 365 MCP Server** transforms how AI assistants interact with Microsoft 365. Instead of simple API wrappers, it provides an **intelligent layer** that understands context, learns from usage patterns, and executes complex multi-step operations seamlessly.

### Why Choose This Server?

| Feature               | Traditional APIs        | Join MS365 MCP Server           |
| --------------------- | ----------------------- | ------------------------------- |
| Context Understanding | ❌ None                 | ✅ Deep semantic understanding  |
| Multi-step Operations | ❌ Manual orchestration | ✅ Automatic chaining           |
| Learning System       | ❌ Static               | ✅ Adaptive learning from usage |
| Natural Language      | ❌ Not supported        | ✅ Ask questions naturally      |
| Token Optimization    | ❌ Standard JSON        | ✅ TOON format (30-60% savings) |

---

## ⚡ Quick Start

### Prerequisites

- **Node.js** >= 18 (recommended: 20+)
- A Microsoft 365 account (personal, work, or school)

### Installation & First Run

```bash
# Install and run with npx
npx @softeria/ms-365-mcp-server

# Or install globally
npm install -g @softeria/ms-365-mcp-server
ms-365-mcp-server
```

### Claude Desktop Integration

Add to your Claude Desktop configuration (`Settings > Developer`):

```json
{
  "mcpServers": {
    "ms365": {
      "command": "npx",
      "args": ["-y", "@softeria/ms-365-mcp-server"]
    }
  }
}
```

### First Authentication

Simply ask Claude: _"Log me into Microsoft 365"_ - the server will guide you through device code authentication.

---

## 🎯 Features

### Core Capabilities

| Capability             | Description                                      |
| ---------------------- | ------------------------------------------------ |
| **90+ MCP Tools**      | Comprehensive coverage of Microsoft 365 services |
| **Intelligent Search** | Cross-product search with semantic understanding |
| **Deep Research**      | Multi-step reasoning for complex questions       |
| **Learning System**    | Improves over time based on usage patterns       |
| **Download Links**     | Generate direct download links for files         |
| **Read-Only Mode**     | Safe exploration without write operations        |
| **Preset Filtering**   | Load only the tools you need                     |

### Supported Microsoft 365 Services

<table>
<tr>
<td width="25%">

**📧 Outlook**

- Email management
- Folder organization
- Attachments
- Drafts

</td>
<td width="25%">

**📅 Calendar**

- Events & meetings
- Scheduling
- Recurring events
- Meeting times

</td>
<td width="25%">

**📁 OneDrive**

- File operations
- Folder management
- Sharing
- Download/Upload

</td>
<td width="25%">

**💬 Teams**

- Chats & messages
- Channels
- Team management
- Transcripts

</td>
</tr>
<tr>
<td>

**🌐 SharePoint**

- Sites & lists
- Document libraries
- Site search
- Permissions

</td>
<td>

**📊 Excel**

- Worksheet operations
- Range manipulation
- Charts
- Formatting

</td>
<td>

**✅ Tasks**

- To-Do lists
- Planner tasks
- Task assignment
- Due dates

</td>
<td>

**📔 OneNote**

- Notebooks
- Sections & pages
- Content creation
- Search

</td>
</tr>
</table>

---

## 🛠 Tool Categories

### 🔐 Authentication Tools

Secure authentication with multi-account support.

| Tool             | Description                       | Notes                             |
| ---------------- | --------------------------------- | --------------------------------- |
| `login`          | Authenticate via device code flow | Required before using other tools |
| `logout`         | Log out from Microsoft account    | Clears cached tokens              |
| `verify-login`   | Check authentication status       | Non-interactive verification      |
| `list-accounts`  | List cached Microsoft accounts    | Multi-account support             |
| `select-account` | Switch between accounts           | Seamless account switching        |
| `remove-account` | Remove account from cache         | Clean up stored credentials       |

---

### 📧 Email & Communication Tools

Complete email management and shared mailbox support.

#### Personal Email

| Tool                        | Description                | Parameters                             |
| --------------------------- | -------------------------- | -------------------------------------- |
| `list-mail-messages`        | List emails with filtering | `top`, `filter`, `search`, `orderby`   |
| `get-mail-message`          | Get full email content     | `messageId`                            |
| `send-mail`                 | Send new email             | `to`, `subject`, `body`, `attachments` |
| `create-draft-email`        | Create email draft         | `to`, `subject`, `body`                |
| `delete-mail-message`       | Delete email               | `messageId`                            |
| `move-mail-message`         | Move email to folder       | `messageId`, `folderId`                |
| `list-mail-folders`         | List all mail folders      | -                                      |
| `list-mail-folder-messages` | List messages in folder    | `folderId`                             |

#### Attachments

| Tool                     | Description             | Parameters                  |
| ------------------------ | ----------------------- | --------------------------- |
| `list-mail-attachments`  | List email attachments  | `messageId`                 |
| `get-mail-attachment`    | Download attachment     | `messageId`, `attachmentId` |
| `add-mail-attachment`    | Add attachment to draft | `messageId`, `file`         |
| `delete-mail-attachment` | Remove attachment       | `messageId`, `attachmentId` |

#### Shared Mailboxes (Org Mode)

| Tool                                  | Description                | Parameters                         |
| ------------------------------------- | -------------------------- | ---------------------------------- |
| `list-shared-mailbox-messages`        | List shared mailbox emails | `user-id`                          |
| `list-shared-mailbox-folder-messages` | List folder messages       | `user-id`, `folderId`              |
| `get-shared-mailbox-message`          | Get shared email content   | `user-id`, `messageId`             |
| `send-shared-mailbox-mail`            | Send from shared mailbox   | `user-id`, `to`, `subject`, `body` |

---

### 📅 Calendar Tools

Full calendar management with meeting scheduling.

| Tool                    | Description          | Parameters                             |
| ----------------------- | -------------------- | -------------------------------------- |
| `list-calendars`        | List all calendars   | -                                      |
| `list-calendar-events`  | List events          | `top`, `filter`, `orderby`             |
| `get-calendar-event`    | Get event details    | `eventId`                              |
| `create-calendar-event` | Create new event     | `subject`, `start`, `end`, `attendees` |
| `update-calendar-event` | Update event         | `eventId`, `updates`                   |
| `delete-calendar-event` | Delete event         | `eventId`                              |
| `get-calendar-view`     | Get calendar view    | `startDateTime`, `endDateTime`         |
| `find-meeting-times`    | Find available slots | `attendees`, `duration`                |

#### Specific Calendar Operations

For working with calendars other than the default:

| Tool                             | Description                         |
| -------------------------------- | ----------------------------------- |
| `list-specific-calendar-events`  | List events from specific calendar  |
| `get-specific-calendar-event`    | Get event from specific calendar    |
| `create-specific-calendar-event` | Create event in specific calendar   |
| `update-specific-calendar-event` | Update event in specific calendar   |
| `delete-specific-calendar-event` | Delete event from specific calendar |

---

### 📁 File & Drive Tools

OneDrive file management with upload/download capabilities.

| Tool                             | Description           | Parameters                        |
| -------------------------------- | --------------------- | --------------------------------- |
| `list-drives`                    | List available drives | -                                 |
| `get-drive-root-item`            | Get drive root folder | `driveId`                         |
| `list-folder-files`              | List files in folder  | `folderId`, `top`                 |
| `download-onedrive-file-content` | Download file content | `itemId`                          |
| `upload-file-content`            | Update file content   | `itemId`, `content`               |
| `upload-new-file`                | Upload new file       | `folderId`, `fileName`, `content` |
| `delete-onedrive-file`           | Delete file           | `itemId`                          |

---

### 📊 Excel Tools

Spreadsheet operations directly through AI.

| Tool                    | Description                 | Parameters                                        |
| ----------------------- | --------------------------- | ------------------------------------------------- |
| `list-excel-worksheets` | List worksheets in workbook | `workbookId`                                      |
| `get-excel-range`       | Get cell range data         | `workbookId`, `worksheetId`, `range`              |
| `format-excel-range`    | Apply formatting            | `workbookId`, `worksheetId`, `range`, `format`    |
| `sort-excel-range`      | Sort data in range          | `workbookId`, `worksheetId`, `range`, `sortBy`    |
| `create-excel-chart`    | Create chart from data      | `workbookId`, `worksheetId`, `chartType`, `range` |

---

### ✅ Task & Planner Tools

Unified task management across To-Do and Planner.

#### Microsoft To-Do

| Tool                   | Description          | Parameters                        |
| ---------------------- | -------------------- | --------------------------------- |
| `list-todo-task-lists` | List all task lists  | -                                 |
| `list-todo-tasks`      | List tasks in a list | `taskListId`                      |
| `get-todo-task`        | Get task details     | `taskListId`, `taskId`            |
| `create-todo-task`     | Create new task      | `taskListId`, `title`, `dueDate`  |
| `update-todo-task`     | Update task          | `taskListId`, `taskId`, `updates` |
| `delete-todo-task`     | Delete task          | `taskListId`, `taskId`            |

#### Microsoft Planner (Org Mode)

| Tool                          | Description             | Parameters                    |
| ----------------------------- | ----------------------- | ----------------------------- |
| `list-planner-tasks`          | List your Planner tasks | -                             |
| `get-planner-plan`            | Get plan details        | `planId`                      |
| `list-plan-tasks`             | List tasks in plan      | `planId`                      |
| `create-planner-task`         | Create Planner task     | `planId`, `title`, `bucketId` |
| `get-planner-task`            | Get task details        | `taskId`                      |
| `update-planner-task`         | Update task             | `taskId`, `updates`           |
| `update-planner-task-details` | Update task details     | `taskId`, `details`           |

---

### 📔 OneNote Tools

Notebook, section, and page management.

| Tool                             | Description            | Parameters                      |
| -------------------------------- | ---------------------- | ------------------------------- |
| `list-onenote-notebooks`         | List all notebooks     | -                               |
| `list-onenote-notebook-sections` | List notebook sections | `notebookId`                    |
| `list-onenote-section-pages`     | List section pages     | `sectionId`                     |
| `get-onenote-page-content`       | Get page content       | `pageId`                        |
| `create-onenote-page`            | Create new page        | `sectionId`, `title`, `content` |

---

### 💬 Microsoft Teams Tools

> **Note:** Requires `--org-mode` flag (work/school accounts only)

#### Chats

| Tool                        | Description           | Parameters                       |
| --------------------------- | --------------------- | -------------------------------- |
| `list-chats`                | List all chats        | `top`                            |
| `get-chat`                  | Get chat details      | `chatId`                         |
| `list-chat-messages`        | List messages in chat | `chatId`, `top`                  |
| `get-chat-message`          | Get message details   | `chatId`, `messageId`            |
| `send-chat-message`         | Send chat message     | `chatId`, `content`              |
| `list-chat-message-replies` | List message replies  | `chatId`, `messageId`            |
| `reply-to-chat-message`     | Reply to message      | `chatId`, `messageId`, `content` |

#### Teams & Channels

| Tool                    | Description           | Parameters                         |
| ----------------------- | --------------------- | ---------------------------------- |
| `list-joined-teams`     | List teams you're in  | -                                  |
| `get-team`              | Get team details      | `teamId`                           |
| `list-team-channels`    | List team channels    | `teamId`                           |
| `get-team-channel`      | Get channel details   | `teamId`, `channelId`              |
| `list-channel-messages` | List channel messages | `teamId`, `channelId`              |
| `get-channel-message`   | Get channel message   | `teamId`, `channelId`, `messageId` |
| `send-channel-message`  | Send channel message  | `teamId`, `channelId`, `content`   |
| `list-team-members`     | List team members     | `teamId`                           |

---

### 🌐 SharePoint Tools

> **Note:** Requires `--org-mode` flag (work/school accounts only)

| Tool                              | Description          | Parameters                    |
| --------------------------------- | -------------------- | ----------------------------- |
| `search-sharepoint-sites`         | Search for sites     | `query`                       |
| `get-sharepoint-site`             | Get site by ID       | `siteId`                      |
| `get-sharepoint-site-by-path`     | Get site by URL path | `hostname`, `path`            |
| `list-sharepoint-site-drives`     | List site drives     | `siteId`                      |
| `get-sharepoint-site-drive-by-id` | Get specific drive   | `siteId`, `driveId`           |
| `list-sharepoint-site-items`      | List site items      | `siteId`, `driveId`           |
| `get-sharepoint-site-item`        | Get item details     | `siteId`, `driveId`, `itemId` |
| `list-sharepoint-site-lists`      | List site lists      | `siteId`                      |
| `get-sharepoint-site-list`        | Get list details     | `siteId`, `listId`            |
| `list-sharepoint-site-list-items` | List items in list   | `siteId`, `listId`            |
| `get-sharepoint-site-list-item`   | Get list item        | `siteId`, `listId`, `itemId`  |
| `get-sharepoint-sites-delta`      | Track site changes   | -                             |

---

### 👥 Contact Tools

Outlook contacts management.

| Tool                     | Description         | Parameters                      |
| ------------------------ | ------------------- | ------------------------------- |
| `list-outlook-contacts`  | List all contacts   | `top`, `filter`                 |
| `get-outlook-contact`    | Get contact details | `contactId`                     |
| `create-outlook-contact` | Create new contact  | `displayName`, `email`, `phone` |
| `update-outlook-contact` | Update contact      | `contactId`, `updates`          |
| `delete-outlook-contact` | Delete contact      | `contactId`                     |

---

### 🔍 Search & Discovery Tools

Powerful cross-product search capabilities.

| Tool                | Description                                        | Use Case                         |
| ------------------- | -------------------------------------------------- | -------------------------------- |
| `ms365-search`      | **Primary search tool** - searches across all M365 | "Find info about Project Alpha"  |
| `search-query`      | Unified Microsoft 365 search                       | Technical search with filters    |
| `search-everything` | Universal fallback search                          | Broad search across all products |
| `search-tools`      | Find available MCP tools                           | "Which tool can send emails?"    |
| `execute-tool`      | Execute tool by name                               | Dynamic tool execution           |

---

### 🧠 Intelligent Compound Tools

These **intelligent tools** automatically chain multiple API calls to answer complex contextual questions. They eliminate the need for manual orchestration.

#### Person-Focused Discovery

| Tool                        | What It Does                         | Example Query                    |
| --------------------------- | ------------------------------------ | -------------------------------- |
| `find-messages-with-person` | Find all Teams chats with a person   | "What did I discuss with John?"  |
| `find-emails-with-person`   | Find all email threads with a person | "Show emails from Sarah"         |
| `find-meetings-with-person` | Find past & future meetings          | "When did I meet with Mike?"     |
| `find-files-from-person`    | Find files shared by person          | "Files John sent me"             |
| `get-communication-summary` | Complete interaction overview        | "Summary of all comms with Lisa" |

#### Entity Discovery

| Tool                | What It Does                       | Example Query                     |
| ------------------- | ---------------------------------- | --------------------------------- |
| `discover-project`  | Find all project-related content   | "Everything about Project Apollo" |
| `discover-person`   | Comprehensive person profile       | "Who is John Smith?"              |
| `discover-meeting`  | Meeting details, notes, recordings | "Details about quarterly review"  |
| `discover-document` | Document info, versions, sharing   | "Find the budget spreadsheet"     |
| `discover-team`     | Team info, members, activity       | "About the Marketing team"        |
| `discover-customer` | Customer contact & history         | "Info about Acme Corp"            |
| `discover-contract` | Contract details & communications  | "Contract with Vendor XYZ"        |
| `discover-decision` | Decision history & rationale       | "Why did we choose Option B?"     |

#### Productivity Tools

| Tool                   | What It Does                        | Example Query                      |
| ---------------------- | ----------------------------------- | ---------------------------------- |
| `get-my-emails`        | Enhanced email view with formatting | "Show my recent emails"            |
| `get-my-week-summary`  | Weekly productivity digest          | "What did I accomplish this week?" |
| `get-all-my-tasks`     | Unified To-Do + Planner tasks       | "What's on my task list?"          |
| `prepare-for-meeting`  | Gather all meeting context          | "Prepare me for the sales meeting" |
| `get-project-overview` | Complete project status             | "Status of Q4 initiative"          |
| `get-company-contacts` | All contacts from a company         | "Who do we know at Microsoft?"     |
| `get-follow-up-items`  | Items needing attention             | "What needs my attention?"         |

---

### 🎯 Business Intelligence Tools

Advanced analytics and insights tools.

| Tool                         | What It Does                    | Example Query                      |
| ---------------------------- | ------------------------------- | ---------------------------------- |
| `analyze-team-collaboration` | Collaboration patterns analysis | "How does the team communicate?"   |
| `get-customer-360`           | 360° customer view              | "Complete picture of Client X"     |
| `analyze-meeting-load`       | Meeting overload analysis       | "Am I in too many meetings?"       |
| `get-deadline-overview`      | All upcoming deadlines          | "What's due this month?"           |
| `find-decision-context`      | Decision history & context      | "History of the platform decision" |
| `get-project-stakeholders`   | Project participant mapping     | "Who's involved in Project Z?"     |
| `find-unresponded-requests`  | Pending response items          | "What's waiting for my reply?"     |
| `get-collaboration-network`  | Professional network mapping    | "Who do I work with most?"         |

#### Meeting Intelligence

| Tool                             | What It Does                       | Example Query                    |
| -------------------------------- | ---------------------------------- | -------------------------------- |
| `get-meeting-transcript-summary` | Meeting transcript with key points | "Summary of yesterday's standup" |
| `list-meetings-with-transcripts` | Find recorded meetings             | "Meetings with transcripts"      |
| `search-across-transcripts`      | Search all transcripts             | "When did we discuss budget?"    |

---

### 🤖 AI Learning System Tools

The server includes an **adaptive learning system** that improves over time.

| Tool                    | What It Does                  | Purpose                                   |
| ----------------------- | ----------------------------- | ----------------------------------------- |
| `deep-research`         | Multi-step reasoning research | Complex questions requiring deep analysis |
| `ask-microsoft-365`     | **Primary AI assistant**      | Natural language questions                |
| `what-can-i-ask`        | Discover available questions  | Learn system capabilities                 |
| `provide-feedback`      | Submit search result feedback | Help improve accuracy                     |
| `export-knowledge-base` | Export learned patterns       | Backup or share knowledge                 |
| `import-knowledge-base` | Import knowledge base         | Restore or merge patterns                 |
| `get-learning-insights` | View learning statistics      | Understand system performance             |

---

## 🔬 Intelligent Discovery System

The server features a sophisticated **Search-First Strategy** with multiple intelligent components:

```
┌─────────────────────────────────────────────────────────────────┐
│                    User Question                                 │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                    NLP Enhancer                                  │
│  • Entity extraction  • Intent classification                   │
│  • Synonym expansion  • Query refinement                        │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                Search-First Strategy                             │
│  • Microsoft Search API (emails, files, chats, events)          │
│  • Learning-informed entity type selection                      │
│  • Automatic query refinement if no results                     │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                 Entity Extractor                                 │
│  • Identifies sites, teams, users, files                        │
│  • Extracts relevant keywords                                   │
│  • Maps to specific product queries                             │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                 Data Aggregator                                  │
│  • Deduplication  • Relevance sorting                           │
│  • LLM-optimized formatting  • Source tracking                  │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                 Learning System                                  │
│  • Records successful patterns  • Updates confidence            │
│  • Learns entity type preferences  • User feedback              │
└─────────────────────────────────────────────────────────────────┘
```

### Deep Research Engine

For complex questions requiring multiple investigation steps:

```typescript
// Example: Deep research on a topic
{
  "tool": "deep-research",
  "arguments": {
    "question": "What's the status of the Q4 marketing initiative?",
    "maxDepth": 3,
    "maxIterations": 5,
    "includeDownloadLinks": true
  }
}
```

Features:

- **Multi-step reasoning** - iteratively refines queries
- **Cross-product correlation** - connects emails, files, meetings
- **Confidence scoring** - indicates result reliability
- **Download link generation** - direct file access

---

## ⚙ Configuration

### CLI Options

| Option                      | Description                                  | Example                            |
| --------------------------- | -------------------------------------------- | ---------------------------------- |
| `--org-mode`                | Enable organization mode (Teams, SharePoint) | `--org-mode`                       |
| `--read-only`               | Disable write operations                     | `--read-only`                      |
| `--http [port]`             | Start HTTP server (default: 3000)            | `--http 8080`                      |
| `--preset <name>`           | Load specific tool presets                   | `--preset mail,calendar`           |
| `--enabled-tools <pattern>` | Filter tools by regex                        | `--enabled-tools "excel\|contact"` |
| `--toon`                    | Enable TOON format (30-60% token savings)    | `--toon`                           |
| `--discovery`               | Start with discovery tools only              | `--discovery`                      |
| `--cloud <type>`            | Cloud environment (global/china)             | `--cloud china`                    |
| `-v`                        | Enable verbose logging                       | `-v`                               |

### Tool Presets

Reduce initial load by using presets:

```bash
npx @softeria/ms-365-mcp-server --preset mail,calendar
npx @softeria/ms-365-mcp-server --list-presets  # See all presets
```

| Preset     | Description            | Tools Included                         |
| ---------- | ---------------------- | -------------------------------------- |
| `mail`     | Email operations       | Mail, folders, attachments             |
| `calendar` | Calendar management    | Events, scheduling                     |
| `files`    | OneDrive operations    | Files, folders, drives                 |
| `personal` | Personal productivity  | Mail, calendar, files, contacts, tasks |
| `work`     | Organization tools     | Teams, SharePoint, shared mailboxes    |
| `excel`    | Spreadsheet operations | Worksheets, ranges, charts             |
| `contacts` | Contact management     | Outlook contacts                       |
| `tasks`    | Task management        | To-Do, Planner                         |
| `onenote`  | Notes                  | Notebooks, sections, pages             |
| `search`   | Search capabilities    | Search queries                         |
| `users`    | Directory              | User listings                          |
| `all`      | Everything             | All 90+ tools                          |

### Environment Variables

| Variable                  | Description                       | Default      |
| ------------------------- | --------------------------------- | ------------ |
| `MS365_MCP_CLIENT_ID`     | Custom Azure app client ID        | Built-in app |
| `MS365_MCP_TENANT_ID`     | Custom tenant ID                  | `common`     |
| `MS365_MCP_CLIENT_SECRET` | Client secret (confidential apps) | -            |
| `MS365_MCP_OAUTH_TOKEN`   | Pre-existing OAuth token (BYOT)   | -            |
| `MS365_MCP_ORG_MODE`      | Enable organization mode          | `false`      |
| `MS365_MCP_OUTPUT_FORMAT` | Output format (`json`/`toon`)     | `json`       |
| `MS365_MCP_CLOUD_TYPE`    | Cloud environment                 | `global`     |
| `MS365_MCP_KEYVAULT_URL`  | Azure Key Vault URL               | -            |
| `MS365_MCP_MAX_RESULTS`   | Maximum search results            | `500`        |
| `READ_ONLY`               | Enable read-only mode             | `false`      |
| `LOG_LEVEL`               | Logging level                     | `info`       |
| `SILENT`                  | Disable console output            | `false`      |

---

## 🔑 Authentication Methods

### 1. Device Code Flow (Default)

Interactive authentication for desktop/CLI applications:

```bash
npx @softeria/ms-365-mcp-server --login
```

Or via MCP tool: Call `login` → Visit URL → Enter code → Call `verify-login`

### 2. OAuth Authorization Code Flow (HTTP Mode)

For web applications and remote servers:

```bash
npx @softeria/ms-365-mcp-server --http 3000
```

- Exposes OAuth endpoints at `/auth/*`
- Requires `Authorization: Bearer <token>` for MCP requests
- Supports MCP OAuth 2.1 with Dynamic Client Registration

### 3. Bring Your Own Token (BYOT)

For integration with existing OAuth systems:

```bash
MS365_MCP_OAUTH_TOKEN=your_token npx @softeria/ms-365-mcp-server
```

### Azure Key Vault Integration

For production deployments:

```bash
MS365_MCP_KEYVAULT_URL=https://your-vault.vault.azure.net npx @softeria/ms-365-mcp-server
```

Store secrets:

- `ms365-mcp-client-id`
- `ms365-mcp-tenant-id`
- `ms365-mcp-client-secret`

---

## ☁️ Cloud Environments

| Cloud      | Description                        | Flag                       |
| ---------- | ---------------------------------- | -------------------------- |
| **Global** | International Microsoft 365        | `--cloud global` (default) |
| **China**  | Microsoft 365 operated by 21Vianet | `--cloud china`            |

```json
{
  "mcpServers": {
    "ms365-china": {
      "command": "npx",
      "args": ["-y", "@softeria/ms-365-mcp-server", "--org-mode", "--cloud", "china"]
    }
  }
}
```

---

## 🔒 Security & Compliance

### Security Features

- ✅ **OAuth 2.1 / PKCE** - Secure token handling
- ✅ **Token validation** - Verified against Microsoft Graph
- ✅ **Secure storage** - OS credential store (keytar) with file fallback
- ✅ **Read-only mode** - Safe exploration without modifications
- ✅ **Input validation** - Zod schema validation on all inputs
- ✅ **Rate limiting** - Configurable request limits

### Compliance

- **ISO 27001** - Information security management
- **GDPR/DSGVO** - Data protection by design
- **OWASP** - Security best practices

### Required Permissions (Scopes)

#### Personal Account

```
User.Read Mail.Read Mail.Send Calendars.ReadWrite
Files.ReadWrite Contacts.ReadWrite Tasks.ReadWrite
Notes.ReadWrite
```

#### Organization Account (--org-mode)

```
+ Team.ReadBasic.All Chat.ReadWrite Channel.ReadBasic.All
+ ChannelMessage.Read.All Sites.Read.All Sites.ReadWrite.All
+ Mail.Read.Shared Mail.Send.Shared User.Read.All
```

---

## 📚 API Reference

### MCP Client Configuration Examples

#### Claude Desktop (Personal)

```json
{
  "mcpServers": {
    "ms365": {
      "command": "npx",
      "args": ["-y", "@softeria/ms-365-mcp-server"]
    }
  }
}
```

#### Claude Desktop (Organization)

```json
{
  "mcpServers": {
    "ms365": {
      "command": "npx",
      "args": ["-y", "@softeria/ms-365-mcp-server", "--org-mode"]
    }
  }
}
```

#### Claude Code CLI

```bash
# Personal
claude mcp add ms365 -- npx -y @softeria/ms-365-mcp-server

# Organization (macOS/Linux)
claude mcp add ms365 -- npx -y @softeria/ms-365-mcp-server --org-mode

# Organization (Windows)
claude mcp add ms365 -s user -- cmd /c "npx -y @softeria/ms-365-mcp-server --org-mode"
```

#### HTTP Mode (Remote/OpenWebUI)

```json
{
  "mcpServers": {
    "ms365": {
      "url": "https://your-server.com/mcp",
      "transportType": "streamable-http"
    }
  }
}
```

### Tool Response Format

All tools return MCP-compliant responses:

```typescript
interface McpToolResult {
  content: Array<{
    type: 'text' | 'image' | 'resource';
    text?: string;
    data?: string;
    mimeType?: string;
    uri?: string;
  }>;
  isError?: boolean;
}
```

---

## 🐳 Docker Deployment

```bash
docker-compose up -d
```

Or build manually:

```bash
docker build -t ms365-mcp-server .
docker run -p 3000:3000 -e MS365_MCP_CLIENT_ID=... ms365-mcp-server
```

---

## 🤝 Contributing

1. Fork the repository
2. Run `npm install`
3. Generate client: `npm run generate`
4. Make changes
5. Run verification: `npm run verify`
6. Submit PR

---

## 📄 License

MIT © 2025 Softeria

---

## 📞 Support

- 📋 [Issues](https://github.com/softeria/ms-365-mcp-server/issues)
- 💬 [Discussions](https://github.com/softeria/ms-365-mcp-server/discussions)
- 📧 Email: eirikb@eirikb.no
- 💬 Discord: https://discord.gg/WvGVNScrAZ

---

<p align="center">
  <strong>Built with ❤️ for the AI community</strong>
</p>
