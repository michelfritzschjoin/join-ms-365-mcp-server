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

---

## 🚀 Enterprise AI Solutions by Join

This MCP Server is developed by **[Join GmbH](https://ki.join.de)** - your partner for secure, enterprise-ready AI solutions in Germany and the EU.

### 🏢 CompanyGPT - ChatGPT for Enterprises

<a href="https://ki.join.de/loesungen/company-gpt/" target="_blank">
  <img src="https://img.shields.io/badge/CompanyGPT-Enterprise_AI-0078D4?style=for-the-badge&logo=openai&logoColor=white" alt="CompanyGPT" />
</a>

**[CompanyGPT](https://ki.join.de/loesungen/company-gpt/)** is a self-hosted, GDPR-compliant AI platform tailored to your business processes:

- ✅ **Full Data Sovereignty** - Self-hosted or EU-hosted, no data leaves your control
- ✅ **DSGVO/GDPR & AI-Act Compliant** - Built for European regulatory requirements
- ✅ **Multi-LLM Support** - Use GPT, Gemini, Copilot, or other models flexibly
- ✅ **Enterprise Integration** - Seamless connection to Microsoft 365, CRM, ERP systems
- ✅ **80% Efficiency Gains** - Automate routine tasks and unlock workforce potential
- ✅ **Knowledge Management** - AI-powered access to your company knowledge base

**Use Cases:** Customer Service Automation • Marketing Content Creation • Internal Knowledge Search • Code Generation • Compliance Monitoring

👉 **[Request a free AI consultation](https://ki.join.de/loesungen/company-gpt/)**

---

### 🏛️ RathausGPT - AI for Public Administration

<a href="https://ki.join.de/loesungen/ki-in-der-verwaltung/" target="_blank">
  <img src="https://img.shields.io/badge/RathausGPT-AI_for_Government-107C10?style=for-the-badge&logo=microsoft&logoColor=white" alt="RathausGPT" />
</a>

**[RathausGPT](https://ki.join.de/loesungen/ki-in-der-verwaltung/)** brings secure AI capabilities to government and public sector organizations:

- 🏛️ **Built for Government** - Designed for the unique needs of public administration
- 🔒 **Highest Security Standards** - BSI-compliant, on-premises deployment options
- 📋 **Citizen Service Automation** - Faster responses, 24/7 availability
- 📄 **Document Processing** - AI-powered forms, applications, and archiving
- 🇩🇪 **Made in Germany** - Local support, German data centers, EU compliance

**Ideal for:** City Halls • District Offices • Government Agencies • Public Utilities

👉 **[Learn more about AI in public administration](https://ki.join.de/loesungen/ki-in-der-verwaltung/)**

---

### Why Choose Join?

| Feature                    | Join AI Solutions                           |
| -------------------------- | ------------------------------------------- |
| 🇪🇺 **EU Data Residency**   | All data stays in certified EU data centers |
| 🔐 **Enterprise Security** | ISO 27001, DSGVO, AI-Act compliance         |
| 🛠️ **Full Customization**  | Tailored workflows for your processes       |
| 📞 **German Support**      | Local experts, German-language assistance   |
| 🔄 **Flexible Models**     | Switch between GPT, Gemini, Copilot anytime |

**Contact:**  
📧 [info@join.de](mailto:info@join.de)  
📞 +49 3691 7090 00  
🌐 [ki.join.de](https://ki.join.de)

---

## License

All Rights Reserved - See [LICENSE](../LICENSE) for details.

---

_This documentation is automatically kept in sync with the codebase._
