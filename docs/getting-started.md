# Getting Started

> **Last Updated:** 2026-01-23  
> **Repository:** https://github.com/michelfritzschjoin/join-ms-365-mcp-server

## Prerequisites

- **Node.js:** Version 18 or higher
- **npm:** Latest stable version
- **Microsoft 365 Account:** Personal or Work/School account

## Installation

### From npm (Recommended)

```bash
npm install -g join-ms-365-mcp-server
```

### From Source

```bash
# Clone the repository
git clone https://github.com/michelfritzschjoin/join-ms-365-mcp-server.git
cd join-ms-365-mcp-server

# Install dependencies
npm install

# Build the project
npm run build
```

## Quick Start

### 1. Initial Setup

```bash
# Create environment file
cp .example.env .env

# Edit configuration (optional)
# The server works with default settings for most use cases
```

### 2. Authentication

```bash
# Login with your Microsoft account
ms-365-mcp-server --login

# Follow the device code flow instructions
# Open the URL and enter the code shown
```

### 3. Verify Connection

```bash
# Test the connection
ms-365-mcp-server --verify-login
```

### 4. Start the Server

#### stdio Mode (for Cursor, Claude Desktop)

```bash
ms-365-mcp-server
```

#### HTTP Mode (for OpenWebUI, web clients)

```bash
ms-365-mcp-server --http 3000
```

## MCP Client Configuration

### Cursor IDE

Add to `.cursor/mcp.json`:

```json
{
  "mcpServers": {
    "ms365": {
      "command": "ms-365-mcp-server",
      "args": []
    }
  }
}
```

### Cursor IDE (HTTP Mode)

```json
{
  "mcpServers": {
    "ms365": {
      "url": "http://localhost:3000/mcp",
      "transportType": "streamable-http"
    }
  }
}
```

### Claude Desktop

Add to Claude Desktop configuration:

```json
{
  "mcpServers": {
    "ms365": {
      "command": "npx",
      "args": ["-y", "join-ms-365-mcp-server"]
    }
  }
}
```

### OpenWebUI

Configure as MCP server with HTTP endpoint:

```
Server URL: http://your-server:3000/mcp
Transport: Streamable HTTP
```

## First Steps

Once connected, try these commands in your AI assistant:

1. **Check your inbox:**

   > "Show me my latest emails"

2. **View calendar:**

   > "What meetings do I have today?"

3. **Search files:**
   > "Find documents about the project proposal"

## Command Line Options

| Option             | Description                             |
| ------------------ | --------------------------------------- |
| `--login`          | Login using device code flow            |
| `--logout`         | Clear saved credentials                 |
| `--verify-login`   | Test connection without starting server |
| `--http [port]`    | Use HTTP transport (default: 3000)      |
| `--read-only`      | Disable write operations                |
| `--org-mode`       | Enable Teams, SharePoint features       |
| `--preset <names>` | Use tool presets (mail, calendar, etc.) |
| `-v`               | Enable verbose logging                  |

## Docker Deployment

### Docker Compose with Traefik

For production deployments with automatic HTTPS via Traefik:

1. **Create the stack environment file:**

```bash
cp stack.env.example stack.env
# Edit stack.env with your Azure AD credentials
```

2. **Configure Traefik labels:**

Edit `docker-compose.yml` and update the host:

```yaml
- 'traefik.http.routers.ms365-mcp.rule=Host(`ms365-mcp.yourdomain.com`)'
```

3. **Create the external network (if not exists):**

```bash
docker network create web
```

4. **Start the service:**

```bash
docker compose up -d
```

### Standalone Docker (Development)

For development or testing without Traefik:

```bash
docker compose --profile standalone up -d
```

This exposes port 3000 directly.

### Docker Run (Simple)

```bash
docker run -d \
  --name ms365-mcp \
  -p 3000:3000 \
  -v ./data:/app/data \
  --env-file stack.env \
  aijoin/join-ms-365-mcp-server:latest \
  --http 3000 -v
```

### Example docker-compose.yml

```yaml
version: '3.9'

services:
  ms365-mcp-server:
    image: aijoin/join-ms-365-mcp-server:latest
    container_name: ms365-mcp
    restart: unless-stopped
    env_file:
      - stack.env
    command: ['--http', '3000', '-v']
    volumes:
      - ./data:/app/data
    labels:
      - 'traefik.enable=true'
      - 'traefik.http.routers.ms365-mcp.rule=Host(`ms365-mcp.yourdomain.com`)'
      - 'traefik.http.routers.ms365-mcp.entrypoints=websecure'
      - 'traefik.http.routers.ms365-mcp.tls.certresolver=myresolver'
      - 'traefik.http.services.ms365-mcp.loadbalancer.server.port=3000'
    networks:
      - web

networks:
  web:
    external: true
```

### Essential Environment Variables

| Variable                 | Description              | Required |
| ------------------------ | ------------------------ | -------- |
| `MS365_MCP_CLIENT_ID`    | Azure AD App Client ID   | Yes\*    |
| `MS365_MCP_TENANT_ID`    | Azure AD Tenant ID       | No       |
| `MS365_MCP_ORG_MODE`     | Enable Teams/SharePoint  | No       |
| `READ_ONLY`              | Disable write operations | No       |
| `MS365_MCP_CORS_ORIGINS` | Allowed CORS origins     | No       |

\*Required for production. Leave empty to use built-in default app for testing.

See [Configuration Guide](./configuration.md) for all environment variables.

## Next Steps

- [Configuration Guide](./configuration.md) - Customize your setup
- [Authentication](./authentication.md) - Azure AD configuration
- [Tools Reference](./tools-reference.md) - Available MCP tools

---

## 🚀 Need Enterprise AI Solutions?

This MCP Server is developed by **[Join GmbH](https://ki.join.de)** - your partner for secure, GDPR-compliant AI in Germany and the EU.

| Solution          | Description                                                    | Link                                                               |
| ----------------- | -------------------------------------------------------------- | ------------------------------------------------------------------ |
| 🏢 **CompanyGPT** | Self-hosted ChatGPT for enterprises with full data sovereignty | [Learn more →](https://ki.join.de/loesungen/company-gpt/)          |
| 🏛️ **RathausGPT** | AI for public administration and government agencies           | [Learn more →](https://ki.join.de/loesungen/ki-in-der-verwaltung/) |

📧 Contact: [info@join.de](mailto:info@join.de) | 📞 +49 3691 7090 00

---

_For issues and support, visit the [GitHub repository](https://github.com/michelfritzschjoin/join-ms-365-mcp-server)._
