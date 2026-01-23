# Development Guide

> **Last Updated:** 2026-01-23  
> **Repository:** https://github.com/michelfritzschjoin/join-ms-365-mcp-server

## Prerequisites

- **Node.js:** 18 or higher
- **npm:** Latest stable version
- **Git:** For version control
- **TypeScript:** Included as dev dependency

## Setup Development Environment

### 1. Clone Repository

```bash
git clone https://github.com/michelfritzschjoin/join-ms-365-mcp-server.git
cd join-ms-365-mcp-server
```

### 2. Install Dependencies

```bash
npm install
```

### 3. Generate Graph Client

```bash
npm run generate
```

### 4. Build Project

```bash
npm run build
```

### 5. Run Tests

```bash
npm test
```

## Project Structure

```
join-ms-365-mcp-server/
├── src/
│   ├── index.ts              # Main entry point
│   ├── server.ts             # Express HTTP server
│   ├── cli.ts                # Command-line interface
│   ├── auth.ts               # Authentication module
│   ├── auth-tools.ts         # Auth MCP tools
│   ├── graph-tools.ts        # Microsoft Graph MCP tools
│   ├── graph-client.ts       # Graph API client
│   ├── discovery-tools.ts    # Discovery MCP tools
│   ├── generated/            # Auto-generated code
│   │   ├── client.ts         # Generated Graph client
│   │   └── endpoint-types.ts # Generated types
│   ├── lib/
│   │   └── microsoft-auth.ts # Microsoft auth helpers
│   └── middleware/           # Express middleware
│       ├── cors.ts
│       ├── rate-limit.ts
│       ├── request-logger.ts
│       └── security-headers.ts
├── bin/
│   ├── generate-graph-client.mjs
│   └── modules/              # Generator modules
├── docs/                     # Documentation
├── test/                     # Test files
├── openapi/                  # OpenAPI specs
├── package.json
├── tsconfig.json
└── tsup.config.ts
```

## Development Workflow

### Daily Development

```bash
# Start development server with watch mode
npm run dev:http

# Run in stdio mode for testing
npm run dev
```

### Before Committing

```bash
# Update dependencies
ncu -u
npm install

# Build
npm run build

# Test
npm test

# Lint
npm run lint

# Format
npm run format
```

### Complete Verification

```bash
npm run verify
```

This runs: generate → lint → format:check → build → test

## NPM Scripts

| Script                 | Description                        |
| ---------------------- | ---------------------------------- |
| `npm run generate`     | Generate Graph client from OpenAPI |
| `npm run build`        | Build TypeScript to JavaScript     |
| `npm run dev`          | Run in development mode (stdio)    |
| `npm run dev:http`     | Run HTTP server with watch mode    |
| `npm test`             | Run tests                          |
| `npm run test:watch`   | Run tests in watch mode            |
| `npm run lint`         | Run ESLint                         |
| `npm run lint:fix`     | Fix ESLint issues                  |
| `npm run format`       | Format code with Prettier          |
| `npm run format:check` | Check code formatting              |
| `npm run verify`       | Full verification pipeline         |
| `npm run inspector`    | Run MCP Inspector                  |

## Code Generation

### Microsoft Graph Client

The Graph client is generated from OpenAPI specifications:

```bash
npm run generate
```

This generates:

- `src/generated/client.ts` - Graph API client methods
- `src/generated/endpoint-types.ts` - TypeScript types

### Generator Modules

Located in `bin/modules/`:

- `download-openapi.mjs` - Download OpenAPI specs
- `simplified-openapi.mjs` - Simplify complex schemas
- `extract-descriptions.mjs` - Extract API descriptions
- `generate-mcp-tools.mjs` - Generate MCP tool definitions

## Adding New Tools

### 1. Define Tool Schema

```typescript
// In src/graph-tools.ts

const newToolSchema = z.object({
  param1: z.string().describe('Parameter description'),
  param2: z.number().optional().describe('Optional parameter'),
});

type NewToolParams = z.infer<typeof newToolSchema>;
```

### 2. Create Tool Definition

```typescript
const newTool: Tool = {
  name: 'new-tool-name',
  description: 'Clear description for AI assistants',
  inputSchema: zodToJsonSchema(newToolSchema),
};
```

### 3. Implement Handler

```typescript
async function handleNewTool(params: NewToolParams): Promise<string> {
  // Validate params
  const validated = newToolSchema.parse(params);

  // Call Graph API
  const result = await graphClient.someOperation(validated);

  // Return JSON string
  return JSON.stringify(result);
}
```

### 4. Register Tool

```typescript
// In tool registration
tools.push(newTool);
handlers.set('new-tool-name', handleNewTool);
```

## Testing

### Unit Tests

```bash
npm test
```

### Test with MCP Inspector

```bash
npm run inspector
```

Opens interactive tool for testing MCP tools.

### Integration Tests

```bash
# Test against real Microsoft Graph API
npm run test -- --run test/graph-api.test.ts
```

## Debugging

### Enable Verbose Logging

```bash
ms-365-mcp-server -v
```

### Debug in VS Code

`.vscode/launch.json`:

```json
{
  "version": "0.2.0",
  "configurations": [
    {
      "type": "node",
      "request": "launch",
      "name": "Debug Server",
      "runtimeExecutable": "npx",
      "runtimeArgs": ["tsx", "src/index.ts"],
      "cwd": "${workspaceFolder}",
      "console": "integratedTerminal"
    }
  ]
}
```

### Check Generated Code

```bash
node bin/check-generated.mjs
```

## Code Style

### TypeScript Guidelines

- Use explicit types
- Avoid `any`
- Use Zod for runtime validation
- Follow functional patterns where appropriate

### Naming Conventions

| Type      | Convention | Example              |
| --------- | ---------- | -------------------- |
| Files     | kebab-case | `graph-tools.ts`     |
| Classes   | PascalCase | `GraphClient`        |
| Functions | camelCase  | `listMailMessages`   |
| Constants | UPPER_CASE | `MAX_ITEMS`          |
| MCP Tools | kebab-case | `list-mail-messages` |

### Code Formatting

```bash
# Format all files
npm run format

# Check formatting
npm run format:check
```

## Dependencies

### Updating Dependencies

```bash
# Check for updates
ncu

# Update all packages
ncu -u
npm install

# Verify nothing is broken
npm run verify
```

### Security Audit

```bash
npm audit
npm audit fix
```

## Docker CI/CD

### GitHub Actions Setup

The project includes automated Docker builds via GitHub Actions. Images are automatically built and pushed to Docker Hub on:

- Push to `main` or `master` branch
- Tag creation (e.g., `v1.0.0`)
- Manual workflow dispatch

### Required Secrets

Configure the following secrets in GitHub repository settings (`Settings` → `Secrets and variables` → `Actions`):

| Secret Name           | Description              |
| --------------------- | ------------------------ |
| `DOCKER_HUB_USERNAME` | Your Docker Hub username |
| `DOCKER_HUB_TOKEN`    | Docker Hub access token  |

### Creating Docker Hub Token

1. Log in to [Docker Hub](https://hub.docker.com)
2. Go to `Account Settings` → `Security`
3. Click `New Access Token`
4. Name it (e.g., "GitHub Actions")
5. Copy the token and add it as `DOCKER_HUB_TOKEN` secret

### Image Tags

The workflow automatically creates multiple tags:

- `latest` - Latest build from main branch
- `main` or `master` - Branch-specific tag
- `v1.0.0` - Semantic version tag
- `v1.0` - Major.minor tag
- `v1` - Major version tag
- `main-<sha>` - Commit SHA tag

### Manual Build

You can manually trigger the workflow:

1. Go to `Actions` tab in GitHub
2. Select `Docker Build and Push` workflow
3. Click `Run workflow`
4. Select branch and click `Run workflow`

### Local Docker Build

```bash
# Build image
docker build -t join-ms-365-mcp-server:latest .

# Run container
docker run -p 3000:3000 \
  -e CLIENT_ID=your-client-id \
  -e CLIENT_SECRET=your-client-secret \
  join-ms-365-mcp-server:latest
```

## Release Process

1. Update version in `package.json`
2. Update CHANGELOG.md
3. Run full verification: `npm run verify`
4. Commit changes: `git commit -m "chore: release vX.Y.Z"`
5. Tag release: `git tag vX.Y.Z`
6. Push: `git push && git push --tags`

Semantic Release handles npm publishing automatically. Docker images are built automatically via GitHub Actions.

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/my-feature`
3. Make changes following code style
4. Write tests for new features
5. Run verification: `npm run verify`
6. Commit: `git commit -m "feat: add my feature"`
7. Push: `git push origin feature/my-feature`
8. Create Pull Request

### Commit Message Format

```
type(scope): description

[optional body]

[optional footer]
```

Types: `feat`, `fix`, `docs`, `style`, `refactor`, `test`, `chore`

---

_For questions, open an issue on [GitHub](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues)._
