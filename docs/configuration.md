# Configuration Guide

> **Last Updated:** 2026-01-23  
> **Repository:** https://github.com/michelfritzschjoin/join-ms-365-mcp-server

## Environment Variables

All configuration is done through environment variables. Copy `.example.env` to `.env` and customize as needed.

### Authentication & Azure Configuration

| Variable                  | Description                           | Default          | Required |
| ------------------------- | ------------------------------------- | ---------------- | -------- |
| `MS365_MCP_CLIENT_ID`     | Azure AD App Client ID                | Built-in default | No       |
| `MS365_MCP_TENANT_ID`     | Azure AD Tenant ID                    | `common`         | No       |
| `MS365_MCP_CLIENT_SECRET` | Client Secret (app-only auth)         | -                | No       |
| `MS365_MCP_CLOUD_TYPE`    | Cloud environment (`global`, `china`) | `global`         | No       |
| `MS365_MCP_OAUTH_TOKEN`   | Pre-existing OAuth token (BYOT)       | -                | No       |
| `MS365_MCP_KEYVAULT_URL`  | Azure Key Vault URL for secrets       | -                | No       |

### Server Mode & Behavior

| Variable                           | Description                            | Default   | Required |
| ---------------------------------- | -------------------------------------- | --------- | -------- |
| `READ_ONLY`                        | Disable write operations               | `false`   | No       |
| `MS365_MCP_ORG_MODE`               | Enable Teams, SharePoint features      | `false`   | No       |
| `ENABLED_TOOLS`                    | Regex pattern to filter tools          | All tools | No       |
| `MS365_MCP_OUTPUT_FORMAT`          | Output format (`toon` for compression) | Default   | No       |
| `MS365_MCP_ENABLE_DISCOVERY_TOOLS` | Enable deep research tools             | `false`   | No       |

### Performance & Limits

| Variable                            | Description           | Default | Required |
| ----------------------------------- | --------------------- | ------- | -------- |
| `MS365_MCP_MAX_RESULTS`             | Max results per query | `500`   | No       |
| `MS365_MCP_MAX_AGGREGATE_ITEMS`     | Max aggregate items   | `500`   | No       |
| `MS365_MCP_MAX_PAGES`               | Max pagination pages  | `500`   | No       |
| `MS365_MCP_MAX_CONCURRENT_TOOLS`    | Max concurrent tools  | `5`     | No       |
| `MS365_MCP_DEEP_RESEARCH_MAX_DEPTH` | Research depth        | `5`     | No       |
| `MS365_MCP_MAX_RESEARCH_ITERATIONS` | Research iterations   | `5`     | No       |
| `MS365_MCP_STOP_ON_ERROR`           | Stop on tool error    | `false` | No       |

### HTTP Server Configuration

| Variable                            | Description                            | Default                           | Required |
| ----------------------------------- | -------------------------------------- | --------------------------------- | -------- |
| `MS365_MCP_CORS_ORIGINS`            | Allowed CORS origins (comma-separated) | Disabled                          | No       |
| `MS365_MCP_CORS_METHODS`            | Allowed HTTP methods                   | `GET, POST, PUT, DELETE, OPTIONS` | No       |
| `MS365_MCP_CORS_HEADERS`            | Allowed headers                        | Standard MCP headers              | No       |
| `MS365_MCP_CORS_MAX_AGE`            | CORS preflight cache (seconds)         | `86400`                           | No       |
| `MS365_MCP_RATE_LIMIT_WINDOW_MS`    | Rate limit window (ms)                 | `900000`                          | No       |
| `MS365_MCP_RATE_LIMIT_MAX_REQUESTS` | Max requests per window                | `100`                             | No       |

### Security Headers

| Variable                       | Description               | Default                           | Required |
| ------------------------------ | ------------------------- | --------------------------------- | -------- |
| `MS365_MCP_X_FRAME_OPTIONS`    | X-Frame-Options header    | `DENY`                            | No       |
| `MS365_MCP_REFERRER_POLICY`    | Referrer-Policy header    | `strict-origin-when-cross-origin` | No       |
| `MS365_MCP_PERMISSIONS_POLICY` | Permissions-Policy header | Restrictive                       | No       |
| `MS365_MCP_CSP`                | Content-Security-Policy   | Restrictive                       | No       |
| `MS365_MCP_HSTS_MAX_AGE`       | HSTS max age (seconds)    | `31536000`                        | No       |

### Logging Configuration

| Variable     | Description                                  | Default | Required |
| ------------ | -------------------------------------------- | ------- | -------- |
| `LOG_LEVEL`  | Log level (`error`, `warn`, `info`, `debug`) | `info`  | No       |
| `LOG_FORMAT` | Log format (`text`, `json`)                  | `text`  | No       |
| `SILENT`     | Disable console output                       | `false` | No       |

### Learning & Knowledge Base

> **Important**: The Learning System requires Discovery Tools to be enabled!
> Set `MS365_MCP_ENABLE_DISCOVERY_TOOLS=true` along with `MS365_MCP_LEARNING_ENABLED=true`.

| Variable                            | Description                         | Default                       | Required |
| ----------------------------------- | ----------------------------------- | ----------------------------- | -------- |
| `MS365_MCP_ENABLE_DISCOVERY_TOOLS`  | Enable discovery tools (required!)  | `false`                       | Yes*     |
| `MS365_MCP_LEARNING_ENABLED`        | Enable learning system              | `true`                        | No       |
| `MS365_MCP_KNOWLEDGE_BASE_PATH`     | Knowledge base file path            | `./data/knowledge-base.json`  | No       |
| `MS365_MCP_LEARNING_DECAY_DAYS`     | Days before pattern decay starts    | `90`                          | No       |
| `MS365_MCP_LEARNING_DECAY_FACTOR`   | Decay rate per month                | `0.1`                         | No       |
| `MS365_MCP_LEARNING_CLUSTER_ENABLED`| Enable pattern clustering           | `true`                        | No       |
| `MS365_MCP_LEARNING_NLP_ENABLED`    | Enable NLP enhancements             | `true`                        | No       |

*Required if you want to use the Learning System

#### Learning System Tools

When properly configured, the following tools are available:

- `learning-status` - Check if the Learning System is active and functioning
- `get-learning-insights` - Get detailed analytics and performance metrics
- `provide-feedback` - Provide explicit feedback to improve learning
- `export-knowledge-base` - Export learned patterns for backup
- `import-knowledge-base` - Import learned patterns from backup

## Tool Presets

Use presets to enable groups of related tools:

```bash
# Enable only mail tools
ms-365-mcp-server --preset mail

# Enable multiple presets
ms-365-mcp-server --preset mail,calendar,contacts
```

### Available Presets

| Preset       | Description         | Org Mode Required |
| ------------ | ------------------- | ----------------- |
| `mail`       | Email operations    | No                |
| `calendar`   | Calendar and events | No                |
| `files`      | OneDrive files      | No                |
| `contacts`   | Contact management  | No                |
| `tasks`      | To-Do and Planner   | No                |
| `onenote`    | OneNote access      | No                |
| `personal`   | All personal tools  | No                |
| `teams`      | Teams operations    | Yes               |
| `sharepoint` | SharePoint access   | Yes               |
| `users`      | User directory      | Yes               |
| `work`       | All work tools      | Yes               |
| `search`     | Search capabilities | No                |
| `excel`      | Excel operations    | No                |
| `all`        | All available tools | Partial           |

## Configuration Examples

### Minimal Setup (Personal Account)

```env
# No configuration needed - uses defaults
```

### Organization Setup

```env
MS365_MCP_ORG_MODE=true
```

### Read-Only Mode

```env
READ_ONLY=true
```

### Production HTTP Server

```env
# Server
MS365_MCP_CORS_ORIGINS=https://your-app.com
MS365_MCP_RATE_LIMIT_MAX_REQUESTS=100

# Security
MS365_MCP_X_FRAME_OPTIONS=DENY
MS365_MCP_CSP=default-src 'self'

# Logging
LOG_LEVEL=warn
LOG_FORMAT=json
NODE_ENV=production
```

### Custom Azure AD App

```env
MS365_MCP_CLIENT_ID=your-client-id
MS365_MCP_TENANT_ID=your-tenant-id
```

### Azure Key Vault Integration

```env
MS365_MCP_KEYVAULT_URL=https://your-vault.vault.azure.net/
```

---

_See [Authentication](./authentication.md) for Azure AD setup details._
