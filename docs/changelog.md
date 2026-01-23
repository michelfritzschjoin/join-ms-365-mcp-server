# Changelog

> **Repository:** https://github.com/michelfritzschjoin/join-ms-365-mcp-server

All notable changes to the Join Microsoft 365 MCP Server are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- Comprehensive documentation in `/docs`
- Security guide with ISO 27001 and DSGVO compliance
- API reference documentation
- Development guide

### Changed

- Updated repository URL to https://github.com/michelfritzschjoin/join-ms-365-mcp-server

## [0.0.0-development] - 2026-01-23

### Added

- Initial MCP server implementation
- Microsoft Graph API integration
- OAuth 2.1 with PKCE authentication
- Device code flow for login
- Support for multiple Microsoft 365 services:
  - Mail (read, send, manage)
  - Calendar (events, meetings)
  - Contacts
  - OneDrive (files, folders)
  - To-Do tasks
  - OneNote
  - Teams (chats, channels)
  - SharePoint (sites, lists)
  - Planner
  - Excel operations
- HTTP server mode with Streamable HTTP
- SSE transport for legacy clients
- stdio mode for local MCP clients
- Read-only mode
- Tool presets for filtering
- Azure Key Vault integration
- Multi-cloud support (Global, China)
- Comprehensive error handling
- Rate limiting
- Security headers
- CORS support

### Security

- Secure token storage using system keychain
- Input validation with Zod schemas
- ISO 27001 compliance measures
- DSGVO/GDPR compliance measures

---

## Version History Format

### Added

New features.

### Changed

Changes in existing functionality.

### Deprecated

Soon-to-be removed features.

### Removed

Removed features.

### Fixed

Bug fixes.

### Security

Vulnerability fixes.

---

_For release details, see [GitHub Releases](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/releases)._
