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

### Search Tools

| Tool           | Description                  | Mode |
| -------------- | ---------------------------- | ---- |
| `search-query` | Unified Microsoft 365 search | All  |

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
