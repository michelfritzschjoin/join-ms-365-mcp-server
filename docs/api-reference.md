# API Reference

> **Last Updated:** 2026-01-23  
> **Repository:** https://github.com/michelfritzschjoin/join-ms-365-mcp-server

## Overview

The Join Microsoft 365 MCP Server supports multiple transport protocols for different deployment scenarios.

## Transport Modes

### stdio Mode (Default)

Standard input/output for local MCP clients:

```bash
# Start in stdio mode
ms-365-mcp-server
```

**Protocol:** JSON-RPC 2.0 over stdio

### HTTP Mode

RESTful HTTP server with SSE support:

```bash
# Start in HTTP mode
ms-365-mcp-server --http 0.0.0.0:3000
```

**Base URL:** `http://localhost:3000`

## HTTP Endpoints

### Discovery

#### OAuth Authorization Server Metadata

```http
GET /.well-known/oauth-authorization-server
```

Response:

```json
{
  "issuer": "https://your-server.com",
  "authorization_endpoint": "/authorize",
  "token_endpoint": "/token",
  "registration_endpoint": "/register",
  "response_types_supported": ["code"],
  "grant_types_supported": ["authorization_code", "refresh_token"],
  "code_challenge_methods_supported": ["S256"]
}
```

#### Protected Resource Metadata

```http
GET /.well-known/oauth-protected-resource
```

Response:

```json
{
  "resource": "https://your-server.com/mcp",
  "authorization_servers": ["https://your-server.com"],
  "scopes_supported": ["User.Read", "Mail.Read", "..."]
}
```

### MCP Endpoints

#### Streamable HTTP

```http
POST /mcp
Content-Type: application/json

{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "tools/list",
  "params": {}
}
```

Response:

```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "result": {
    "tools": [...]
  }
}
```

#### Server-Sent Events (SSE)

```http
GET /sse
Accept: text/event-stream
```

Events:

```
event: endpoint
data: /mcp

event: message
data: {"jsonrpc":"2.0",...}
```

### Health Check

```http
GET /health
```

Response:

```json
{
  "status": "ok",
  "version": "0.0.0-development",
  "timestamp": "2026-01-23T12:00:00Z"
}
```

## JSON-RPC Methods

### initialize

Initialize the MCP connection:

```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "initialize",
  "params": {
    "protocolVersion": "2024-11-05",
    "clientInfo": {
      "name": "OpenWebUI",
      "version": "1.0.0"
    },
    "capabilities": {}
  }
}
```

Response:

```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "result": {
    "protocolVersion": "2024-11-05",
    "serverInfo": {
      "name": "join-ms-365-mcp-server",
      "version": "0.0.0-development"
    },
    "capabilities": {
      "tools": { "listChanged": true },
      "resources": { "listChanged": true }
    }
  }
}
```

### tools/list

List available tools:

```json
{
  "jsonrpc": "2.0",
  "id": 2,
  "method": "tools/list",
  "params": {}
}
```

### tools/call

Execute a tool:

```json
{
  "jsonrpc": "2.0",
  "id": 3,
  "method": "tools/call",
  "params": {
    "name": "list-mail-messages",
    "arguments": {
      "top": 10
    }
  }
}
```

Response:

```json
{
  "jsonrpc": "2.0",
  "id": 3,
  "result": {
    "content": [
      {
        "type": "text",
        "text": "{\"value\": [...]}"
      }
    ],
    "isError": false
  }
}
```

### resources/list

List available resources:

```json
{
  "jsonrpc": "2.0",
  "id": 4,
  "method": "resources/list",
  "params": {}
}
```

### resources/read

Read a resource:

```json
{
  "jsonrpc": "2.0",
  "id": 5,
  "method": "resources/read",
  "params": {
    "uri": "ms365://mail/message-id"
  }
}
```

## Error Codes

Standard JSON-RPC error codes:

| Code   | Name             | Description            |
| ------ | ---------------- | ---------------------- |
| -32700 | Parse error      | Invalid JSON           |
| -32600 | Invalid request  | Invalid request object |
| -32601 | Method not found | Unknown method         |
| -32602 | Invalid params   | Invalid parameters     |
| -32603 | Internal error   | Server error           |

Application error codes:

| Code   | Name                    | Description              |
| ------ | ----------------------- | ------------------------ |
| -32001 | Authentication required | User must login          |
| -32002 | Permission denied       | Insufficient permissions |
| -32003 | Resource not found      | Requested item not found |
| -32004 | Rate limited            | Too many requests        |

## Authentication Headers

For HTTP mode with pre-authenticated tokens:

```http
Authorization: Bearer <access-token>
```

Or via environment:

```env
MS365_MCP_OAUTH_TOKEN=<access-token>
```

## CORS Configuration

For browser-based clients, CORS is enabled:

```typescript
const CORS_CONFIG = {
  origin: '*', // Configure as needed
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
  credentials: true,
};
```

## Rate Limiting

HTTP mode includes rate limiting:

| Tier           | Limit | Window   |
| -------------- | ----- | -------- |
| Default        | 100   | 1 minute |
| Auth endpoints | 10    | 1 minute |
| High-volume    | 1000  | 1 minute |

Rate limit headers:

```http
X-RateLimit-Limit: 100
X-RateLimit-Remaining: 95
X-RateLimit-Reset: 1706011260
```

## Client Examples

### cURL

```bash
# Initialize
curl -X POST http://localhost:3000/mcp \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","clientInfo":{"name":"curl","version":"1.0"},"capabilities":{}}}'

# List tools
curl -X POST http://localhost:3000/mcp \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}}'
```

### JavaScript/TypeScript

```typescript
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StreamableHTTPClientTransport } from '@modelcontextprotocol/sdk/client/streamableHttp.js';

const client = new Client(
  {
    name: 'my-client',
    version: '1.0.0',
  },
  {
    capabilities: {},
  }
);

const transport = new StreamableHTTPClientTransport(new URL('http://localhost:3000/mcp'));

await client.connect(transport);

// List tools
const tools = await client.listTools();

// Call a tool
const result = await client.callTool({
  name: 'list-mail-messages',
  arguments: { top: 10 },
});
```

### Python

```python
import httpx

async with httpx.AsyncClient() as client:
    # Initialize
    response = await client.post(
        'http://localhost:3000/mcp',
        json={
            'jsonrpc': '2.0',
            'id': 1,
            'method': 'initialize',
            'params': {
                'protocolVersion': '2024-11-05',
                'clientInfo': {'name': 'python', 'version': '1.0'},
                'capabilities': {}
            }
        }
    )

    # Call tool
    response = await client.post(
        'http://localhost:3000/mcp',
        json={
            'jsonrpc': '2.0',
            'id': 2,
            'method': 'tools/call',
            'params': {
                'name': 'list-mail-messages',
                'arguments': {'top': 10}
            }
        }
    )
```

---

_For MCP specification details, see [modelcontextprotocol.io](https://modelcontextprotocol.io)._
