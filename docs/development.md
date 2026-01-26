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

### Automatic Versioning with Semantic Release

This project uses **semantic-release** for fully automated versioning based on commit messages. **No manual version updates are required.**

When code is pushed to `main`:

1. Semantic-release analyzes commit messages
2. Determines the next version (major/minor/patch)
3. Updates `package.json` version
4. Updates `docs/changelog.md` automatically
5. Creates a Git tag and GitHub Release
6. Triggers Docker image build

### How Versions Are Determined

| Commit Type        | Description             | Version Bump  |
| ------------------ | ----------------------- | ------------- |
| `feat:`            | New feature             | Minor (0.X.0) |
| `fix:`             | Bug fix                 | Patch (0.0.X) |
| `perf:`            | Performance improvement | Patch         |
| `refactor:`        | Code refactoring        | Patch         |
| `BREAKING CHANGE:` | Breaking changes        | Major (X.0.0) |

### Example Workflow

```bash
# Make changes
git add .

# Commit using conventional format
git commit -m "feat(mail): add attachment download support"

# Push to main
git push origin main

# Semantic release handles the rest automatically!
```

### Dry Run

Preview what will happen without making changes:

```bash
npm run release:dry-run
```

### NPM Scripts for Release

| Script                    | Description                        |
| ------------------------- | ---------------------------------- |
| `npm run release`         | Run semantic-release (CI only)     |
| `npm run release:dry-run` | Preview release without publishing |
| `npm run commitlint`      | Validate commit message format     |

Docker images are built automatically via GitHub Actions when new versions are tagged.

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/my-feature`
3. Make changes following code style
4. Write tests for new features
5. Run verification: `npm run verify`
6. Commit: `git commit -m "feat: add my feature"`
7. Push: `git push origin feature/my-feature`
8. Create Pull Request

### Commit Message Format (Conventional Commits)

This project enforces **Conventional Commits** for all commit messages using commitlint and husky hooks.

```
type(scope): description

[optional body]

[optional footer]
```

#### Commit Types

| Type         | Description                   | Triggers Release |
| ------------ | ----------------------------- | ---------------- |
| `feat`       | New feature                   | ✅ Minor         |
| `fix`        | Bug fix                       | ✅ Patch         |
| `docs`       | Documentation only            | ❌               |
| `style`      | Code style (formatting, etc.) | ❌               |
| `refactor`   | Code refactoring              | ✅ Patch         |
| `perf`       | Performance improvements      | ✅ Patch         |
| `test`       | Adding or fixing tests        | ❌               |
| `build`      | Build system changes          | ❌               |
| `ci`         | CI/CD configuration           | ❌               |
| `chore`      | Other changes                 | ❌               |
| `revert`     | Revert previous commit        | ✅ Patch         |
| `security`   | Security fixes                | ✅ Patch         |
| `compliance` | GDPR/ISO compliance           | ✅ Patch         |

#### Examples

```bash
# Feature
git commit -m "feat(calendar): add recurring event support"

# Bug fix
git commit -m "fix(auth): resolve token refresh on expired session"

# Breaking change (triggers major version)
git commit -m "feat(api)!: change response format for mail endpoints

BREAKING CHANGE: Mail response now returns array instead of object"

# Documentation
git commit -m "docs: update API reference with new endpoints"

# Security fix
git commit -m "security: sanitize user input in search queries"
```

#### Validation

Commit messages are validated automatically by husky hooks. Invalid commits will be rejected:

```bash
# ❌ Bad - no type
git commit -m "added new feature"

# ❌ Bad - wrong format
git commit -m "FEAT: Add feature"

# ✅ Good
git commit -m "feat: add new feature"
```

#### Git Hooks

The following hooks are automatically installed:

- **pre-commit**: Runs format check, lint, build, and tests
- **commit-msg**: Validates commit message format with commitlint

---

_For questions, open an issue on [GitHub](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues)._
