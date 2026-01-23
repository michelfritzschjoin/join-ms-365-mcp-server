# Join Microsoft 365 MCP Server Documentation

> **Version:** 0.0.0-development  
> **Last Updated:** 2026-01-23  
> **Protocol:** Model Context Protocol (MCP)  
> **Repository:** https://github.com/michelfritzschjoin/join-ms-365-mcp-server

## Overview

The Join Microsoft 365 MCP Server is a Model Context Protocol (MCP) server that provides AI assistants with secure access to Microsoft 365 services through the Microsoft Graph API.

## Documentation Index

| Document                                | Description                            |
| --------------------------------------- | -------------------------------------- |
| [Getting Started](./getting-started.md) | Quick start guide and installation     |
| [Configuration](./configuration.md)     | Environment variables and settings     |
| [Authentication](./authentication.md)   | Azure AD setup and auth flows          |
| [Tools Reference](./tools-reference.md) | Complete MCP tools documentation       |
| [Security](./security.md)               | Security best practices and compliance |
| [API Reference](./api-reference.md)     | HTTP endpoints and protocols           |
| [Development](./development.md)         | Contributing and development guide     |
| [Changelog](./changelog.md)             | Version history and changes            |

## Quick Links

- **Repository:** https://github.com/michelfritzschjoin/join-ms-365-mcp-server
- **MCP Specification:** https://modelcontextprotocol.io
- **Microsoft Graph API:** https://learn.microsoft.com/en-us/graph/

## Architecture

```
┌─────────────────────┐         ┌─────────────────────┐         ┌─────────────────────┐
│    MCP CLIENT       │         │  MS365 MCP SERVER   │         │  Microsoft Graph    │
│  (Cursor, Claude,   │◄───────►│                     │◄───────►│       API           │
│   OpenWebUI)        │  MCP    │  - Auth Management  │  REST   │                     │
│                     │         │  - Tool Execution   │         │  - Mail             │
└─────────────────────┘         │  - Token Handling   │         │  - Calendar         │
                                └─────────────────────┘         │  - Files            │
                                                                │  - Teams            │
                                                                └─────────────────────┘
```

## Supported MCP Clients

| Client         | Support Level | Transport      |
| -------------- | ------------- | -------------- |
| Cursor IDE     | ✅ Full       | stdio, HTTP    |
| Claude Desktop | ✅ Full       | stdio          |
| OpenWebUI      | ✅ Full       | HTTP, SSE      |
| Continue.dev   | ✅ Full       | stdio          |
| Custom Clients | ✅ Full       | All transports |

## Features

### Core Capabilities

- 📧 **Email Management** - Read, send, search, and organize emails
- 📅 **Calendar Operations** - Events, meetings, and scheduling
- 📁 **File Operations** - OneDrive and SharePoint file access
- 👥 **Contacts** - Contact management and lookups
- ✅ **Tasks** - Microsoft To-Do and Planner integration
- 📝 **OneNote** - Notebook and page access
- 🔍 **Search** - Unified search across Microsoft 365
- 👔 **Teams** - Team and channel operations (Org Mode)
- 🌐 **SharePoint** - Site and list access (Org Mode)

### Security & Compliance

- 🔐 ISO 27001 compliant architecture
- 🛡️ DSGVO/GDPR compliant data handling
- 🔑 OAuth 2.1 with PKCE
- 🚫 Read-only mode support
- 📊 Comprehensive audit logging

## License

All Rights Reserved - See [LICENSE](../LICENSE) for details.

---

_This documentation is automatically kept in sync with the codebase._
