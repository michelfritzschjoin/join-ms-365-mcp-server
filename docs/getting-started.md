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

## Next Steps

- [Configuration Guide](./configuration.md) - Customize your setup
- [Authentication](./authentication.md) - Azure AD configuration
- [Tools Reference](./tools-reference.md) - Available MCP tools

---

_For issues and support, visit the [GitHub repository](https://github.com/michelfritzschjoin/join-ms-365-mcp-server)._
