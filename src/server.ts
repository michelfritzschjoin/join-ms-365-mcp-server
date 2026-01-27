import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { mcpAuthRouter } from '@modelcontextprotocol/sdk/server/auth/router.js';
import express, { Request, Response } from 'express';
import logger, { enableConsoleLogging } from './logger.js';
import { registerAuthTools } from './auth-tools.js';
import { registerGraphTools, registerDiscoveryTools } from './graph-tools.js';
import { registerDiscoveryTools as registerIntelligentDiscoveryTools } from './discovery-tools.js';
import { registerCompoundTools } from './compound-tools.js';
import GraphClient from './graph-client.js';
import AuthManager, { buildScopesFromEndpoints } from './auth.js';
import KnowledgeBase from './knowledge-base.js';
import { MicrosoftOAuthProvider } from './oauth-provider.js';
import {
  exchangeCodeForToken,
  createMicrosoftBearerTokenAuthMiddleware,
  refreshAccessToken,
} from './lib/microsoft-auth.js';
import type { CommandOptions } from './cli.ts';
import { getSecrets, type AppSecrets } from './secrets.js';
import { getCloudEndpoints } from './cloud-config.js';
import { requestContext, createTokenHash } from './request-context.js';
import { randomUUID } from 'crypto';
import { createDashboardRouter, isDashboardEnabled } from './query-dashboard.js';
import { getQueryStore } from './query-store.js';
import { z } from 'zod';

/**
 * Extract chat ID from request headers
 * Checks multiple possible header names in priority order:
 * 1. X-OpenWebUI-Chat-ID (OpenWebUI specific)
 * 2. X-Chat-ID (generic)
 * 3. X-Conversation-ID (alternative)
 * 4. X-Session-ID (fallback)
 * 5. Generate UUID if none provided
 *
 * @param req - Express request object
 * @returns Chat ID string
 */
function extractChatId(req: Request): string {
  const chatId =
    req.get('X-OpenWebUI-Chat-ID') ||
    req.get('X-Chat-ID') ||
    req.get('X-Conversation-ID') ||
    req.get('X-Session-ID') ||
    req.get('x-openwebui-chat-id') ||
    req.get('x-chat-id') ||
    req.get('x-conversation-id') ||
    req.get('x-session-id');

  if (chatId) {
    return chatId;
  }

  // Generate a UUID for this request if no chat ID provided
  // This ensures each "anonymous" session still gets tracked
  return `anon-${randomUUID()}`;
}

/**
 * Extract user ID from Microsoft auth context or token
 * @param req - Express request with microsoftAuth
 * @returns User ID string or undefined
 */
function extractUserId(
  req: Request & { microsoftAuth?: { accessToken: string; refreshToken?: string } }
): string | undefined {
  if (!req.microsoftAuth?.accessToken) {
    return undefined;
  }

  // Try to extract user ID from JWT token (without full validation - just decoding)
  try {
    const token = req.microsoftAuth.accessToken;
    const parts = token.split('.');
    if (parts.length === 3) {
      const payload = JSON.parse(Buffer.from(parts[1], 'base64').toString('utf8'));
      // Microsoft tokens typically have 'oid' (object ID) or 'sub' (subject)
      return payload.oid || payload.sub || payload.unique_name;
    }
  } catch {
    // Token parsing failed, return undefined
  }

  return undefined;
}

/**
 * Parse HTTP option into host and port components.
 * Supports formats: "host:port", ":port", "port"
 * @param httpOption - The HTTP option value (string or boolean)
 * @returns Object with host (undefined if not specified) and port number
 */
function parseHttpOption(httpOption: string | boolean): { host: string | undefined; port: number } {
  if (typeof httpOption === 'boolean') {
    return { host: undefined, port: 3000 };
  }

  const httpString = httpOption.trim();

  // Check if it contains a colon (host:port format)
  if (httpString.includes(':')) {
    const [hostPart, portPart] = httpString.split(':');
    const host = hostPart || undefined; // Empty string becomes undefined
    const port = parseInt(portPart) || 3000;
    return { host, port };
  }

  // No colon, treat as port only
  const port = parseInt(httpString) || 3000;
  return { host: undefined, port };
}

/**
 * Sanitize tool parameters to remove sensitive information before logging
 * SECURITY: Never log passwords, tokens, or other secrets
 */
function sanitizeToolParams(params: Record<string, unknown>): Record<string, unknown> {
  const sensitiveKeys = [
    'password',
    'token',
    'secret',
    'key',
    'authorization',
    'bearer',
    'credential',
    'apikey',
    'api_key',
    'access_token',
    'refresh_token',
    'client_secret',
  ];

  const sanitized: Record<string, unknown> = {};

  for (const [key, value] of Object.entries(params)) {
    const lowerKey = key.toLowerCase();
    if (sensitiveKeys.some((sensitive) => lowerKey.includes(sensitive))) {
      sanitized[key] = '[REDACTED]';
    } else if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
      sanitized[key] = sanitizeToolParams(value as Record<string, unknown>);
    } else {
      sanitized[key] = value;
    }
  }

  return sanitized;
}

class MicrosoftGraphServer {
  private authManager: AuthManager;
  private options: CommandOptions;
  private graphClient: GraphClient | null;
  private server: McpServer | null;
  private secrets: AppSecrets | null;
  private version: string;

  constructor(authManager: AuthManager, options: CommandOptions = {}) {
    this.authManager = authManager;
    this.options = options;
    this.graphClient = null; // Initialized in start() after secrets are loaded
    this.server = null;
    this.secrets = null;
    this.version = '0.0.0-development';
  }

  async initialize(version: string): Promise<void> {
    this.version = version;
    // Load secrets first
    this.secrets = await getSecrets();

    // Initialize GraphClient with secrets
    const outputFormat = this.options.toon ? 'toon' : 'json';
    this.graphClient = new GraphClient(this.authManager, this.secrets, outputFormat);

    this.server = new McpServer({
      name: 'Microsoft365MCP',
      version,
    });

    const shouldRegisterAuthTools = !this.options.http || this.options.enableAuthTools;
    if (shouldRegisterAuthTools) {
      registerAuthTools(this.server, this.authManager);
    }

    if (this.options.discovery) {
      logger.info('Discovery mode enabled (experimental) - registering discovery tool only');
      registerDiscoveryTools(
        this.server,
        this.graphClient,
        this.options.readOnly,
        this.options.orgMode
      );
    } else {
      // Initialize knowledge base for tool usage learning
      const knowledgeBase = new KnowledgeBase();
      registerGraphTools(
        this.server,
        this.graphClient,
        this.options.readOnly,
        this.options.enabledTools,
        this.options.orgMode,
        knowledgeBase
      );
    }

    // Register intelligent discovery tools if enabled
    // NOTE: The Learning System requires discovery tools to be enabled!
    const discoveryToolsEnabled =
      process.env.MS365_MCP_ENABLE_DISCOVERY_TOOLS === 'true' || this.options.enableDiscoveryTools;

    if (discoveryToolsEnabled) {
      logger.info('Intelligent discovery tools enabled');
      if (!this.secrets) {
        throw new Error('Secrets not loaded');
      }
      registerIntelligentDiscoveryTools(this.server, this.graphClient, this.secrets);
    } else {
      // Warn if learning is enabled but discovery tools are not
      const learningEnabled =
        process.env.MS365_MCP_LEARNING_ENABLED !== 'false' &&
        process.env.MS365_MCP_LEARNING_ENABLED !== '0';

      if (learningEnabled && process.env.MS365_MCP_LEARNING_ENABLED) {
        logger.warn(
          'MS365_MCP_LEARNING_ENABLED is set but Learning System requires Discovery Tools! ' +
            'Set MS365_MCP_ENABLE_DISCOVERY_TOOLS=true to activate the Learning System.'
        );
      }
    }

    // Register compound tools (multi-step contextual tools)
    // These are always enabled as they provide essential functionality for natural language queries
    const compoundToolCount = registerCompoundTools(
      this.server,
      this.graphClient,
      this.options.readOnly
    );
    logger.info(`Registered ${compoundToolCount} compound tools (multi-step contextual tools)`);
  }

  async start(): Promise<void> {
    if (this.options.v) {
      enableConsoleLogging();
    }

    logger.info('Microsoft 365 MCP Server starting...');

    // Debug: Check if secrets are loaded
    logger.info('Secrets Check:', {
      CLIENT_ID: this.secrets?.clientId ? `${this.secrets.clientId.substring(0, 8)}...` : 'NOT SET',
      CLIENT_SECRET: this.secrets?.clientSecret ? 'SET' : 'NOT SET',
      TENANT_ID: this.secrets?.tenantId || 'NOT SET',
      NODE_ENV: process.env.NODE_ENV || 'NOT SET',
    });

    if (this.options.readOnly) {
      logger.info('Server running in READ-ONLY mode. Write operations are disabled.');
    }

    if (this.options.http) {
      const { host, port } = parseHttpOption(this.options.http);

      const app = express();
      // Configure trust proxy securely for rate limiting
      // Use number of proxies instead of 'true' to prevent IP spoofing
      // In production behind a reverse proxy (nginx/traefik), typically 1 proxy
      // Set via TRUST_PROXY_COUNT env var or default to 1
      const trustProxyCount = parseInt(process.env.TRUST_PROXY_COUNT || '1', 10);
      app.set('trust proxy', trustProxyCount);

      // Import middleware
      const { securityHeadersMiddleware } = await import('./middleware/security-headers.js');
      const { rateLimitMiddleware } = await import('./middleware/rate-limit.js');
      const { corsMiddleware } = await import('./middleware/cors.js');
      const { requestLoggerMiddleware } = await import('./middleware/request-logger.js');

      // Enable request logging FIRST if verbose mode is enabled (before other middleware)
      // This ensures we capture all requests including health checks
      if (this.options.v || process.env.DEBUG_REQUESTS === 'true') {
        logger.info('Request logging enabled - all HTTP requests will be logged');
        app.use(requestLoggerMiddleware);
      }

      app.use(express.json());
      app.use(express.urlencoded({ extended: true }));

      // Apply middleware in order
      app.use(securityHeadersMiddleware);
      app.use(rateLimitMiddleware);
      app.use(corsMiddleware);

      // Build available scopes from endpoints
      const availableScopes = buildScopesFromEndpoints(
        this.options.orgMode,
        this.options.enabledTools
      );

      // Initialize OAuth provider with configuration
      const oauthProvider = new MicrosoftOAuthProvider(this.authManager, this.secrets!, {
        port,
        scopes: availableScopes,
      });

      // Create authentication middleware that REQUIRES OAuth
      // This ensures all MCP endpoints require valid Bearer tokens
      const requireOAuthMiddleware = createMicrosoftBearerTokenAuthMiddleware({
        requireAuth: true,
      });

      // OAuth Authorization Server Discovery (RFC 8414)
      app.get('/.well-known/oauth-authorization-server', async (req, res) => {
        const protocol = req.secure ? 'https' : 'http';
        const baseUrl = `${protocol}://${req.get('host')}`;

        // Update provider base URL dynamically
        oauthProvider.setBaseUrl(baseUrl);

        res.json({
          ...oauthProvider.getAuthorizationServerMetadata(baseUrl),
          scopes_supported: availableScopes,
        });
      });

      // OAuth Protected Resource Discovery (RFC 9728)
      app.get('/.well-known/oauth-protected-resource', async (req, res) => {
        const protocol = req.secure ? 'https' : 'http';
        const baseUrl = `${protocol}://${req.get('host')}`;

        res.json({
          ...oauthProvider.getProtectedResourceMetadata(baseUrl),
          scopes_supported: availableScopes,
        });
      });

      // Resource-specific OAuth Authorization Server Discovery (for /mcp resource)
      // OpenWebUI and other MCP clients may request this endpoint
      app.get('/.well-known/oauth-authorization-server/mcp', async (req, res) => {
        const protocol = req.secure ? 'https' : 'http';
        const url = new URL(`${protocol}://${req.get('host')}`);

        const scopes = buildScopesFromEndpoints(this.options.orgMode, this.options.enabledTools);

        res.json({
          issuer: url.origin,
          authorization_endpoint: `${url.origin}/authorize`,
          token_endpoint: `${url.origin}/token`,
          response_types_supported: ['code'],
          response_modes_supported: ['query'],
          grant_types_supported: ['authorization_code', 'refresh_token'],
          token_endpoint_auth_methods_supported: ['none'],
          code_challenge_methods_supported: ['S256'],
          scopes_supported: scopes,
          resource: `${url.origin}/mcp`,
        });
      });

      // OpenID Connect Configuration Discovery (for /mcp resource)
      // Some MCP clients may request OpenID Connect discovery endpoints
      app.get('/.well-known/openid-configuration/mcp', async (req, res) => {
        const protocol = req.secure ? 'https' : 'http';
        const url = new URL(`${protocol}://${req.get('host')}`);

        const scopes = buildScopesFromEndpoints(this.options.orgMode, this.options.enabledTools);

        res.json({
          issuer: url.origin,
          authorization_endpoint: `${url.origin}/authorize`,
          token_endpoint: `${url.origin}/token`,
          response_types_supported: ['code'],
          response_modes_supported: ['query'],
          grant_types_supported: ['authorization_code', 'refresh_token'],
          token_endpoint_auth_methods_supported: ['none'],
          code_challenge_methods_supported: ['S256'],
          scopes_supported: scopes,
          resource: `${url.origin}/mcp`,
        });
      });

      // OpenID Connect Configuration Discovery (alternative path)
      app.get('/mcp/.well-known/openid-configuration', async (req, res) => {
        const protocol = req.secure ? 'https' : 'http';
        const url = new URL(`${protocol}://${req.get('host')}`);

        const scopes = buildScopesFromEndpoints(this.options.orgMode, this.options.enabledTools);

        res.json({
          issuer: url.origin,
          authorization_endpoint: `${url.origin}/authorize`,
          token_endpoint: `${url.origin}/token`,
          response_types_supported: ['code'],
          response_modes_supported: ['query'],
          grant_types_supported: ['authorization_code', 'refresh_token'],
          token_endpoint_auth_methods_supported: ['none'],
          code_challenge_methods_supported: ['S256'],
          scopes_supported: scopes,
          resource: `${url.origin}/mcp`,
        });
      });

      // Authorization endpoint - redirects to Microsoft
      app.get('/authorize', async (req, res) => {
        const url = new URL(req.url!, `${req.protocol}://${req.get('host')}`);
        const tenantId = this.secrets?.tenantId || 'common';
        const clientId = this.secrets!.clientId;
        const cloudEndpoints = getCloudEndpoints(this.secrets!.cloudType);
        const microsoftAuthUrl = new URL(
          `${cloudEndpoints.authority}/${tenantId}/oauth2/v2.0/authorize`
        );

        // #region agent log
        const requestHost = req.get('host');
        const requestProtocol = req.protocol;
        const requestUrl = `${requestProtocol}://${requestHost}${req.url}`;
        const redirectUriParam = url.searchParams.get('redirect_uri');
        fetch('http://127.0.0.1:7245/ingest/76c7865f-57f2-4bf0-8001-38b29d141bbc', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            location: 'server.ts:189',
            message: 'OAuth authorize request',
            data: {
              requestHost,
              requestProtocol,
              requestUrl,
              redirectUriParam,
              isNgrok: requestHost?.includes('ngrok') || false,
            },
            timestamp: Date.now(),
            sessionId: 'debug-session',
            runId: 'run1',
            hypothesisId: 'D',
          }),
        }).catch(() => {});
        // #endregion

        // Store the original redirect_uri from the MCP client in the state parameter
        // so we can redirect back to it after receiving the authorization code
        const originalRedirectUri = url.searchParams.get('redirect_uri');
        const originalState = url.searchParams.get('state');

        // Determine the actual client URL
        // If redirect_uri is our own callback endpoint, extract client URL from Referer header
        const protocol = req.secure ? 'https' : 'http';
        const ourCallbackUrl = `${protocol}://${req.get('host')}/callback`;
        let clientUrl = originalRedirectUri || null;

        // If redirect_uri is our own callback, try to get client URL from Referer or Origin
        if (originalRedirectUri === ourCallbackUrl || !originalRedirectUri) {
          const referer = req.get('Referer') || req.get('Referrer');
          const origin = req.get('Origin');

          // Prefer Referer over Origin, as it's more likely to be the actual client page
          if (referer) {
            try {
              const refererUrl = new URL(referer);
              // Use the referer's origin as the client URL
              clientUrl = refererUrl.origin;
              logger.info('Extracted client URL from Referer header', {
                referer,
                clientUrl,
              });
            } catch {
              // Invalid referer URL, ignore
            }
          } else if (origin) {
            clientUrl = origin;
            logger.info('Extracted client URL from Origin header', {
              origin,
              clientUrl,
            });
          }
        }

        // Encode the client URL and state in a new state parameter
        // IMPORTANT: Preserve originalState exactly, including empty strings, for CSRF protection
        let enhancedState = originalState ?? '';
        if (clientUrl) {
          const stateData = {
            client_url: clientUrl,
            redirect_uri: originalRedirectUri || ourCallbackUrl,
            // Use undefined check to preserve empty string values
            original_state: originalState !== undefined ? originalState : null,
          };
          // Encode as base64url (URL-safe base64)
          const base64 = Buffer.from(JSON.stringify(stateData)).toString('base64');
          const base64url = base64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
          enhancedState = base64url;

          logger.info('Encoded enhanced state for Microsoft OAuth', {
            hasClientUrl: true,
            hasRedirectUri: !!originalRedirectUri,
            originalStatePreserved: originalState !== undefined,
          });
        } else if (originalRedirectUri && originalRedirectUri !== ourCallbackUrl) {
          // Fallback: use original redirect_uri if it's not our callback
          const stateData = {
            redirect_uri: originalRedirectUri,
            original_state: originalState !== undefined ? originalState : null,
          };
          const base64 = Buffer.from(JSON.stringify(stateData)).toString('base64');
          const base64url = base64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
          enhancedState = base64url;

          logger.info('Encoded enhanced state for Microsoft OAuth (redirect_uri only)', {
            hasRedirectUri: true,
            originalStatePreserved: originalState !== undefined,
          });
        }

        // Only forward parameters that Microsoft OAuth 2.0 v2.0 supports
        const allowedParams = [
          'response_type',
          'scope',
          'response_mode',
          'code_challenge',
          'code_challenge_method',
          'prompt',
          'login_hint',
          'domain_hint',
        ];

        allowedParams.forEach((param) => {
          const value = url.searchParams.get(param);
          if (value) {
            microsoftAuthUrl.searchParams.set(param, value);
          }
        });

        // Use our Microsoft app's client_id
        microsoftAuthUrl.searchParams.set('client_id', clientId);

        // Set our own callback URL as redirect_uri (Microsoft will redirect here)
        // Note: protocol and ourCallbackUrl are already defined above
        microsoftAuthUrl.searchParams.set('redirect_uri', ourCallbackUrl);

        // Set the enhanced state parameter (contains original redirect_uri)
        if (enhancedState) {
          microsoftAuthUrl.searchParams.set('state', enhancedState);
        }

        // Ensure we have the minimal required scopes if none provided
        if (!microsoftAuthUrl.searchParams.get('scope')) {
          microsoftAuthUrl.searchParams.set('scope', 'User.Read Files.Read Mail.Read');
        }

        // Redirect to Microsoft's authorization page
        res.redirect(microsoftAuthUrl.toString());
      });

      // Mount MCP Auth Router under /auth to avoid conflicting with our explicit OAuth endpoints
      const authRouter = mcpAuthRouter({
        provider: oauthProvider,
        issuerUrl: new URL(`http://localhost:${port}`),
      });

      // #region agent log
      // Instrumentation: Track requests before router
      app.use((req, res, next) => {
        if (req.path === '/callback' || req.url?.startsWith('/callback')) {
          fetch('http://127.0.0.1:7245/ingest/76c7865f-57f2-4bf0-8001-38b29d141bbc', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              location: 'server.ts:238',
              message: 'Request to /callback before router',
              data: {
                path: req.path,
                url: req.url,
                method: req.method,
                headersSent: res.headersSent,
              },
              timestamp: Date.now(),
              sessionId: 'debug-session',
              runId: 'run1',
              hypothesisId: 'A',
            }),
          }).catch(() => {});
        }
        next();
      });
      // #endregion

      // Mount under /auth so it doesn't intercept /authorize, /token, /callback, /register
      app.use('/auth', authRouter);

      // #region agent log
      // Instrumentation: Track requests after router
      app.use((req, res, next) => {
        if (req.path === '/callback' || req.url?.startsWith('/callback')) {
          fetch('http://127.0.0.1:7245/ingest/76c7865f-57f2-4bf0-8001-38b29d141bbc', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              location: 'server.ts:252',
              message: 'Request to /callback after router',
              data: {
                path: req.path,
                url: req.url,
                method: req.method,
                headersSent: res.headersSent,
              },
              timestamp: Date.now(),
              sessionId: 'debug-session',
              runId: 'run1',
              hypothesisId: 'A',
            }),
          }).catch(() => {});
        }
        next();
      });
      // #endregion

      // OAuth callback endpoint - receives authorization code from Microsoft
      // The mcpAuthRouter is mounted before this handler, so it will handle the callback first if it matches
      // If the router doesn't handle /callback (it likely handles /auth/callback instead), our handler will execute
      app.get('/callback', async (req, res) => {
        // #region agent log
        const requestHost = req.get('host');
        const requestProtocol = req.protocol;
        const requestUrl = `${requestProtocol}://${requestHost}${req.url}`;
        const isNgrok = requestHost?.includes('ngrok') || false;
        fetch('http://127.0.0.1:7245/ingest/76c7865f-57f2-4bf0-8001-38b29d141bbc', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            location: 'server.ts:243',
            message: 'Custom /callback handler executing',
            data: {
              headersSent: res.headersSent,
              hasCode: !!req.query.code,
              url: req.url,
              requestHost,
              requestProtocol,
              requestUrl,
              isNgrok,
            },
            timestamp: Date.now(),
            sessionId: 'debug-session',
            runId: 'run1',
            hypothesisId: 'D',
          }),
        }).catch(() => {});
        // #endregion
        try {
          logger.info('OAuth callback received', {
            query: Object.keys(req.query),
            hasCode: !!req.query.code,
            hasError: !!req.query.error,
            state: req.query.state,
          });

          // SECURITY: Validate and sanitize OAuth callback parameters to prevent type confusion
          // Note: redirect_uri can be URLs, URNs (urn:ietf:wg:oauth:2.0:oob), or custom schemes
          // Microsoft authorization codes can be 2500+ characters long
          const oauthCallbackSchema = z.object({
            code: z.string().max(4000).optional(), // Microsoft codes can be 2500+ chars
            error: z.string().max(200).optional(),
            error_description: z.string().max(2000).optional(), // Increased for verbose MS error descriptions
            state: z.string().max(4000).optional(), // Increased for base64-encoded state with client info
            format: z.enum(['html', 'json']).optional(),
            exchange_token: z.enum(['true', '1', 'false', '0']).optional(),
            code_verifier: z.string().max(200).optional(),
            redirect_uri: z.string().max(2000).optional(), // Removed .url() - can be URN or custom scheme
          });

          let validatedParams: z.infer<typeof oauthCallbackSchema>;
          try {
            validatedParams = oauthCallbackSchema.parse({
              code: req.query.code,
              error: req.query.error,
              error_description: req.query.error_description,
              state: req.query.state,
              format: req.query.format,
              exchange_token: req.query.exchange_token,
              code_verifier: req.query.code_verifier,
              redirect_uri: req.query.redirect_uri,
            });
          } catch (validationError) {
            logger.warn('Invalid OAuth callback parameters', {
              error: validationError,
              query: {
                hasCode: !!req.query.code,
                hasError: !!req.query.error,
                hasState: !!req.query.state,
                stateLength: typeof req.query.state === 'string' ? req.query.state.length : 0,
              },
            });
            return res.status(400).json({
              error: 'invalid_request',
              error_description: 'Invalid request parameters',
            });
          }

          const code = validatedParams.code;
          const error = validatedParams.error;
          const errorDescription = validatedParams.error_description;
          const state = validatedParams.state;

          // Handle OAuth errors from Microsoft
          if (error) {
            logger.error('OAuth callback error from Microsoft', {
              error,
              errorDescription,
              state,
            });

            // Return error page - don't redirect to avoid loops
            return res.status(400).send(`
              <!DOCTYPE html>
              <html>
                <head>
                  <title>OAuth Error</title>
                  <meta charset="utf-8">
                  <style>
                    body {
                      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
                      max-width: 600px;
                      margin: 50px auto;
                      padding: 20px;
                      background: #f5f5f5;
                    }
                    .container {
                      background: white;
                      padding: 30px;
                      border-radius: 8px;
                      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                    }
                    h1 {
                      color: #d13438;
                      margin-top: 0;
                    }
                    .error {
                      color: #d13438;
                      font-weight: bold;
                    }
                    code {
                      background: #f5f5f5;
                      padding: 2px 6px;
                      border-radius: 3px;
                      font-size: 0.9em;
                    }
                  </style>
                </head>
                <body>
                  <div class="container">
                    <h1>✗ OAuth Authorization Error</h1>
                    <p class="error"><strong>Error:</strong> ${error}</p>
                    ${errorDescription ? `<p><strong>Description:</strong> ${errorDescription}</p>` : ''}
                    <p>Please try again or contact support if the problem persists.</p>
                  </div>
                  <script>
                    // Send error to parent window if in iframe
                    if (window.parent !== window) {
                      window.parent.postMessage({
                        type: 'oauth-error',
                        error: '${error}',
                        error_description: '${errorDescription || ''}',
                        state: '${state || ''}'
                      }, '*');
                    }
                  </script>
                </body>
              </html>
            `);
          }

          // Handle successful authorization
          if (code) {
            logger.info('Authorization code received successfully', {
              codeLength: code.length,
              state,
            });

            // Check if client wants HTML response (via Accept header or format parameter)
            // Default to JSON for programmatic clients (MCP protocol expects structured responses)
            // Only return HTML if explicitly requested (browsers with text/html Accept header)
            const acceptHeader = req.get('Accept') || '';
            const formatParam = validatedParams.format;
            // Only return HTML if:
            // 1. format=html is explicitly set, OR
            // 2. Accept header includes text/html AND doesn't include application/json
            const wantsHtml =
              formatParam === 'html' ||
              (acceptHeader.includes('text/html') &&
                !acceptHeader.includes('application/json') &&
                formatParam !== 'json');
            // Default to JSON for all MCP clients - HTML only if explicitly requested
            const wantsJson = !wantsHtml;

            // #region agent log
            const userAgent = req.get('User-Agent') || '';
            fetch('http://127.0.0.1:7245/ingest/76c7865f-57f2-4bf0-8001-38b29d141bbc', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                location: 'server.ts:394',
                message: 'Checking response format',
                data: {
                  wantsJson,
                  wantsHtml,
                  acceptHeader: req.get('Accept'),
                  userAgent: userAgent.substring(0, 100),
                  formatParam: req.query.format,
                  headersSent: res.headersSent,
                },
                timestamp: Date.now(),
                sessionId: 'debug-session',
                runId: 'run1',
                hypothesisId: 'B',
              }),
            }).catch(() => {});
            // #endregion

            // Default to JSON for programmatic clients (MCP protocol expects structured responses)
            // Only return HTML if explicitly requested (browsers with text/html Accept header)

            // Check if client explicitly requests token exchange (via query parameter)
            const requestToken =
              validatedParams.exchange_token === 'true' || validatedParams.exchange_token === '1';

            // Try to automatically exchange code for token ONLY if:
            // 1. Client explicitly requests it (exchange_token=true), AND
            // 2. We have a code_verifier (for PKCE flows) OR client_secret (for confidential clients)
            // Otherwise, return the code and let the MCP client (like OpenWebUI) handle the exchange
            const codeVerifier = validatedParams.code_verifier;
            const hasCodeVerifier = !!codeVerifier;
            const hasClientSecret = !!this.secrets?.clientSecret;

            // Only attempt automatic token exchange if explicitly requested AND we have the necessary credentials
            const shouldAttemptAutoExchange = requestToken && (hasCodeVerifier || hasClientSecret);

            if (shouldAttemptAutoExchange) {
              const protocol = req.secure ? 'https' : 'http';
              const baseUrl = `${protocol}://${req.get('host')}`;
              // SECURITY: Use validated redirect_uri or fallback to our callback
              const redirectUri = validatedParams.redirect_uri || `${baseUrl}/callback`;

              // Attempt automatic token exchange for MCP clients
              try {
                const tenantId = this.secrets?.tenantId || 'common';
                const clientId = this.secrets!.clientId;
                const clientSecret = this.secrets?.clientSecret;

                logger.info('Attempting automatic token exchange for MCP client', {
                  redirectUri,
                  hasClientSecret: !!clientSecret,
                  hasCodeVerifier,
                });

                const { exchangeCodeForToken } = await import('./lib/microsoft-auth.js');
                const tokenResult = await exchangeCodeForToken(
                  code,
                  redirectUri,
                  clientId,
                  clientSecret,
                  tenantId,
                  codeVerifier, // Use code_verifier if provided (for PKCE)
                  this.secrets!.cloudType
                );

                // Return token directly to MCP client - always JSON for MCP clients
                logger.info('Token exchange successful, returning token to client');
                if (wantsJson) {
                  return res.json({
                    ...tokenResult,
                    message: 'Token exchange successful',
                  });
                }
                // Only return HTML if explicitly requested
                return res.send(`
                  <!DOCTYPE html>
                  <html>
                    <head>
                      <title>Authorization Successful</title>
                      <meta charset="utf-8">
                      <script>
                        // Return token via postMessage
                        const tokenData = ${JSON.stringify(tokenResult)};
                        
                        // Send to parent/opener
                        if (window.parent !== window) {
                          window.parent.postMessage({ type: 'oauth-token', ...tokenData }, '*');
                        }
                        if (window.opener && !window.opener.closed) {
                          window.opener.postMessage({ type: 'oauth-token', ...tokenData }, '*');
                        }
                        
                        // Store token
                        try {
                          localStorage.setItem('oauth_access_token', tokenData.access_token);
                          localStorage.setItem('oauth_refresh_token', tokenData.refresh_token);
                        } catch(e) {}
                        
                        // Close window
                        setTimeout(() => { if (window.opener) window.close(); }, 500);
                      </script>
                    </head>
                    <body>
                      <h1>Authorization Successful</h1>
                      <p>Token received. Closing window...</p>
                    </body>
                  </html>
                `);
              } catch (tokenError) {
                const errorMessage = (tokenError as Error).message;
                const isPKCEError =
                  errorMessage.includes('code_verifier') ||
                  errorMessage.includes('code_challenge') ||
                  errorMessage.includes('AADSTS50148');

                if (isPKCEError) {
                  logger.info(
                    'PKCE detected - skipping automatic token exchange, returning code to client',
                    {
                      error: errorMessage,
                    }
                  );
                } else {
                  logger.warn('Automatic token exchange failed, falling back to code return', {
                    error: errorMessage,
                  });
                }
                // Fall through to return code (JSON or HTML based on wantsJson)
              }
            }

            // Extract the client URL from the state parameter
            // The state parameter contains base64url-encoded JSON with client_url, redirect_uri and original_state
            let clientUrl: string | undefined;
            let originalRedirectUri: string | undefined;
            let originalState: string | null = null;

            if (state) {
              try {
                // Try to decode the state parameter as base64url-encoded JSON
                // base64url uses - and _ instead of + and /, and no padding
                const base64 = state.replace(/-/g, '+').replace(/_/g, '/');
                const padding = base64.length % 4;
                const paddedBase64 = base64 + (padding ? '='.repeat(4 - padding) : '');
                const stateData = JSON.parse(Buffer.from(paddedBase64, 'base64').toString('utf8'));

                // Extract the client URL and redirect_uri separately
                clientUrl = stateData.client_url;
                originalRedirectUri = stateData.redirect_uri;
                // IMPORTANT: Preserve original_state exactly as sent, including empty strings
                // Use explicit check for undefined to avoid losing empty string values
                // This is critical for CSRF protection - state must match exactly
                originalState =
                  stateData.original_state !== undefined ? stateData.original_state : null;

                logger.info('Decoded state parameter from Microsoft callback', {
                  hasClientUrl: !!clientUrl,
                  hasRedirectUri: !!originalRedirectUri,
                  originalStatePresent: stateData.original_state !== undefined,
                  originalStateValue:
                    originalState !== null ? `[${originalState.length} chars]` : 'null',
                });
              } catch {
                // State is not encoded, use it as-is (raw state from client)
                logger.info('State is not base64url-encoded, using raw value');
                originalState = state;
              }
            }

            // PRIORITY 1: Always prefer originalRedirectUri from the MCP client's OAuth request
            // This is the redirect_uri the client specified and expects to receive the callback
            if (originalRedirectUri) {
              const protocol = req.secure ? 'https' : 'http';
              const ourCallbackUrl = `${protocol}://${req.get('host')}/callback`;

              // Only redirect if it's not our own callback (to avoid infinite loops)
              if (originalRedirectUri !== ourCallbackUrl) {
                try {
                  const redirectUrl = new URL(originalRedirectUri);
                  redirectUrl.searchParams.set('code', code);
                  // CRITICAL: Always include state if it was provided, even if empty string
                  // MCP clients perform CSRF validation - state must match exactly
                  if (originalState !== null) {
                    redirectUrl.searchParams.set('state', originalState);
                  }

                  logger.info('Redirecting to MCP client redirect_uri with authorization code', {
                    redirectUri: originalRedirectUri.substring(0, 100), // SECURITY: Truncate in logs
                    codeLength: code.length,
                    stateIncluded: originalState !== null,
                    stateLength: originalState?.length ?? 0,
                  });

                  return res.redirect(redirectUrl.toString());
                } catch (urlError) {
                  logger.warn('Invalid originalRedirectUri, falling back', {
                    originalRedirectUri,
                    error: (urlError as Error).message,
                  });
                }
              }
            }

            // PRIORITY 2: Use clientUrl (from Referer header) if no valid redirect_uri
            if (clientUrl) {
              try {
                const redirectUrl = new URL(clientUrl);
                // If client_url is just an origin (no path), append /oauth/callback
                // per MCP OAuth specification
                if (redirectUrl.pathname === '/' || redirectUrl.pathname === '') {
                  redirectUrl.pathname = '/oauth/callback';
                }

                redirectUrl.searchParams.set('code', code);
                if (originalState !== null) {
                  redirectUrl.searchParams.set('state', originalState);
                }

                logger.info('Redirecting to MCP client (from Referer) with authorization code', {
                  clientUrl: redirectUrl.toString().substring(0, 100), // SECURITY: Truncate in logs
                  codeLength: code.length,
                  stateIncluded: originalState !== null,
                });

                return res.redirect(redirectUrl.toString());
              } catch {
                logger.warn('Invalid clientUrl from Referer', { clientUrl });
              }
            }

            // Fallback: If no redirect_uri is provided, return JSON for programmatic clients
            if (wantsJson) {
              logger.info('Returning authorization code as JSON response for MCP client');
              return res.json({
                code,
                state: state || null,
                message: 'Authorization code received successfully',
              });
            }

            // Last resort: Return HTML only if explicitly requested (browsers with text/html Accept header)
            const protocol = req.secure ? 'https' : 'http';
            const baseUrl = `${protocol}://${req.get('host')}`;

            logger.info(
              'No redirect_uri provided, returning authorization code via HTML page with postMessage',
              {
                codeLength: code.length,
                stateLength: state?.length ?? 0, // SECURITY: Log length, not content
              }
            );

            // Return success page - code is in URL, client can extract it
            // Also provide multiple ways for the client to receive the code
            return res.send(`
              <!DOCTYPE html>
              <html>
                <head>
                  <title>Authorization Successful</title>
                  <meta charset="utf-8">
                  <style>
                    body {
                      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
                      max-width: 600px;
                      margin: 50px auto;
                      padding: 20px;
                      background: #f5f5f5;
                    }
                    .container {
                      background: white;
                      padding: 30px;
                      border-radius: 8px;
                      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
                    }
                    h1 {
                      color: #107c10;
                      margin-top: 0;
                    }
                    .success {
                      color: #107c10;
                      font-weight: bold;
                    }
                    code {
                      background: #f5f5f5;
                      padding: 2px 6px;
                      border-radius: 3px;
                      font-size: 0.9em;
                    }
                  </style>
                </head>
                <body>
                  <div class="container">
                    <h1>✓ Authorization Successful</h1>
                    <p class="success">Authorization code received successfully.</p>
                    <p>You can close this window. The authorization code has been processed.</p>
                    <p><small>Code: <code>${code.substring(0, 20)}...</code></small></p>
                  </div>
                  <script>
                    (function() {
                      const code = '${code}';
                      const state = '${state || ''}';
                      const callbackData = {
                        type: 'oauth-callback',
                        code: code,
                        state: state,
                        url: window.location.href
                      };
                      
                      // Also send in MCP-compatible format
                      const mcpCallbackData = {
                        type: 'mcp-oauth-callback',
                        code: code,
                        state: state || null
                      };

                      // Method 1: Send to parent window if in iframe (for OpenWebUI and other MCP clients)
                      if (window.parent !== window) {
                        try {
                          // Send both formats for compatibility
                          window.parent.postMessage(callbackData, '*');
                          window.parent.postMessage(mcpCallbackData, '*');
                          console.log('Sent authorization code to parent window via postMessage');
                          // #region agent log
                          fetch('http://127.0.0.1:7245/ingest/76c7865f-57f2-4bf0-8001-38b29d141bbc', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({
                              location: 'server.ts:588',
                              message: 'postMessage sent to parent window',
                              data: { hasParent: window.parent !== window, codeLength: callbackData.code.length, sentBothFormats: true },
                              timestamp: Date.now(),
                              sessionId: 'debug-session',
                              runId: 'run1',
                              hypothesisId: 'C'
                            })
                          }).catch(() => {});
                          // #endregion
                        } catch (e) {
                          console.error('Failed to send to parent:', e);
                          // #region agent log
                          fetch('http://127.0.0.1:7245/ingest/76c7865f-57f2-4bf0-8001-38b29d141bbc', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({
                              location: 'server.ts:602',
                              message: 'postMessage to parent failed',
                              data: { error: String(e) },
                              timestamp: Date.now(),
                              sessionId: 'debug-session',
                              runId: 'run1',
                              hypothesisId: 'C'
                            })
                          }).catch(() => {});
                          // #endregion
                        }
                      }

                      // Method 2: Send to opener window if opened as popup (for OpenWebUI and other MCP clients)
                      if (window.opener && !window.opener.closed) {
                        try {
                          // Send both formats for compatibility
                          window.opener.postMessage(callbackData, '*');
                          window.opener.postMessage(mcpCallbackData, '*');
                          console.log('Sent authorization code to opener window via postMessage');
                          // #region agent log
                          fetch('http://127.0.0.1:7245/ingest/76c7865f-57f2-4bf0-8001-38b29d141bbc', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({
                              location: 'server.ts:620',
                              message: 'postMessage sent to opener window',
                              data: { hasOpener: !!window.opener, openerClosed: window.opener?.closed, codeLength: callbackData.code.length, sentBothFormats: true },
                              timestamp: Date.now(),
                              sessionId: 'debug-session',
                              runId: 'run1',
                              hypothesisId: 'C'
                            })
                          }).catch(() => {});
                          // #endregion
                        } catch (e) {
                          console.error('Failed to send to opener:', e);
                          // #region agent log
                          fetch('http://127.0.0.1:7245/ingest/76c7865f-57f2-4bf0-8001-38b29d141bbc', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({
                              location: 'server.ts:632',
                              message: 'postMessage to opener failed',
                              data: { error: String(e) },
                              timestamp: Date.now(),
                              sessionId: 'debug-session',
                              runId: 'run1',
                              hypothesisId: 'C'
                            })
                          }).catch(() => {});
                          // #endregion
                        }
                      }

                      // Method 3: Store in localStorage (for same-origin clients)
                      try {
                        localStorage.setItem('oauth_code', code);
                        localStorage.setItem('oauth_state', state);
                        localStorage.setItem('oauth_callback_time', Date.now().toString());
                        console.log('Stored code in localStorage');
                      } catch (e) {
                        console.warn('Could not store in localStorage:', e);
                      }

                      // Method 4: Store in sessionStorage
                      try {
                        sessionStorage.setItem('oauth_code', code);
                        sessionStorage.setItem('oauth_state', state);
                        console.log('Stored code in sessionStorage');
                      } catch (e) {
                        console.warn('Could not store in sessionStorage:', e);
                      }

                      // Method 5: Dispatch custom event
                      try {
                        const event = new CustomEvent('oauth-callback', {
                          detail: callbackData,
                          bubbles: true,
                          cancelable: true
                        });
                        window.dispatchEvent(event);
                        document.dispatchEvent(event);
                        console.log('Dispatched oauth-callback event');
                      } catch (e) {
                        console.warn('Could not dispatch event:', e);
                      }

                      // Method 6: Set window.name (for cross-origin popups)
                      try {
                        window.name = JSON.stringify(callbackData);
                        console.log('Set window.name with callback data');
                      } catch (e) {
                        console.warn('Could not set window.name:', e);
                      }

                      // Auto-close popup windows after a delay
                      if (window.opener && !window.opener.closed) {
                        setTimeout(function() {
                          try {
                            window.close();
                          } catch (e) {
                            console.warn('Could not close window:', e);
                          }
                        }, 1000);
                      }
                    })();
                  </script>
                </body>
              </html>
            `);
          }

          // No code and no error - invalid request
          logger.warn('OAuth callback received without code or error');
          res.status(400).send(`
            <html>
              <head><title>Invalid Request</title></head>
              <body>
                <h1>Invalid OAuth Callback</h1>
                <p>No authorization code or error received.</p>
              </body>
            </html>
          `);
        } catch (error) {
          logger.error('OAuth callback error:', error);
          if (!res.headersSent) {
            res.status(500).send(`
              <html>
                <head><title>Server Error</title></head>
                <body>
                  <h1>Server Error</h1>
                  <p>An error occurred processing the OAuth callback.</p>
                </body>
              </html>
            `);
          }
        }
      });

      // OAuth callback endpoint - also handle POST requests (if Microsoft uses POST)
      app.post('/callback', async (req, res) => {
        try {
          logger.info('OAuth callback POST received', {
            body: Object.keys(req.body || {}),
            hasCode: !!req.body?.code,
            hasError: !!req.body?.error,
          });

          // Handle POST callback similar to GET
          const code = req.body?.code as string | undefined;
          const error = req.body?.error as string | undefined;
          const errorDescription = req.body?.error_description as string | undefined;
          const state = req.body?.state as string | undefined;

          if (error) {
            logger.error('OAuth callback POST error', { error, errorDescription, state });
            return res.status(400).json({
              error,
              error_description: errorDescription,
              state,
            });
          }

          if (code) {
            logger.info('Authorization code received via POST', { codeLength: code.length, state });
            // Return JSON response for POST requests
            return res.json({
              code,
              state,
              message: 'Authorization code received successfully',
            });
          }

          res.status(400).json({
            error: 'invalid_request',
            error_description: 'No authorization code or error received',
          });
        } catch (error) {
          logger.error('OAuth callback POST error:', error);
          res.status(500).json({
            error: 'server_error',
            error_description: 'Internal server error processing callback',
          });
        }
      });

      // Token exchange endpoint
      app.post('/token', async (req, res) => {
        try {
          // Log token endpoint call (redact sensitive data)
          logger.info('Token endpoint called', {
            method: req.method,
            url: req.url,
            contentType: req.get('Content-Type'),
            grant_type: req.body?.grant_type,
          });

          const body = req.body;

          // Add debugging and validation
          if (!body) {
            logger.error('Token endpoint: Request body is undefined');
            res.status(400).json({
              error: 'invalid_request',
              error_description: 'Request body is required',
            });
            return;
          }

          if (!body.grant_type) {
            logger.error('Token endpoint: grant_type is missing', { body });
            res.status(400).json({
              error: 'invalid_request',
              error_description: 'grant_type parameter is required',
            });
            return;
          }

          if (body.grant_type === 'authorization_code') {
            const tenantId = this.secrets?.tenantId || 'common';
            const clientId = this.secrets!.clientId;
            const clientSecret = this.secrets?.clientSecret;

            // Log whether using public or confidential client
            if (clientSecret) {
              logger.info('Token endpoint: Using confidential client with client_secret');
            } else {
              logger.info('Token endpoint: Using public client without client_secret');
            }

            // CRITICAL: redirect_uri must match EXACTLY what was sent to Microsoft in /authorize
            // Microsoft redirects to OUR callback URL, so we must use OUR callback URL here
            const protocol = req.secure ? 'https' : 'http';
            const ourCallbackUrl = `${protocol}://${req.get('host')}/callback`;

            // Use our callback URL (what Microsoft redirected to) instead of client's redirect_uri
            // The client's redirect_uri is only used for the initial authorization redirect
            const redirectUri = ourCallbackUrl;

            logger.info('Token exchange request', {
              hasCode: !!body.code,
              hasCodeVerifier: !!body.code_verifier,
              hasClientSecret: !!clientSecret,
              redirectUri,
              clientRedirectUri: body.redirect_uri,
            });

            const result = await exchangeCodeForToken(
              body.code as string,
              redirectUri, // Use our callback URL, not the client's redirect_uri
              clientId,
              clientSecret,
              tenantId,
              body.code_verifier as string | undefined,
              this.secrets!.cloudType
            );
            res.json(result);
          } else if (body.grant_type === 'refresh_token') {
            const tenantId = this.secrets?.tenantId || 'common';
            const clientId = this.secrets!.clientId;
            const clientSecret = this.secrets?.clientSecret;

            // Log whether using public or confidential client
            if (clientSecret) {
              logger.info('Refresh endpoint: Using confidential client with client_secret');
            } else {
              logger.info('Refresh endpoint: Using public client without client_secret');
            }

            const result = await refreshAccessToken(
              body.refresh_token as string,
              clientId,
              clientSecret,
              tenantId,
              this.secrets!.cloudType
            );
            res.json(result);
          } else {
            res.status(400).json({
              error: 'unsupported_grant_type',
              error_description: `Grant type '${body.grant_type}' is not supported`,
            });
          }
        } catch (error) {
          logger.error('Token endpoint error:', error);

          // Extract more details from the error
          let errorDescription = 'Internal server error during token exchange';
          let statusCode = 500;

          if (error instanceof Error) {
            errorDescription = error.message;
            // Check if it's a Microsoft API error (400/401/403)
            if (
              error.message.includes('400') ||
              error.message.includes('401') ||
              error.message.includes('403')
            ) {
              statusCode = 400; // Return 400 for client errors from Microsoft
            }
            // Log the full error for debugging
            logger.error('Token exchange error details:', {
              message: error.message,
              stack: error.stack,
              body: req.body
                ? { ...req.body, code: req.body.code ? '[REDACTED]' : undefined }
                : undefined,
            });
          }

          res.status(statusCode).json({
            error: 'server_error',
            error_description: errorDescription,
          });
        }
      });

      // OAuth Dynamic Client Registration endpoint (RFC 7591)
      // This endpoint allows MCP clients to dynamically register for OAuth flows
      // @see https://tools.ietf.org/html/rfc7591
      // @see https://modelcontextprotocol.io
      app.post('/register', async (req, res) => {
        try {
          const protocol = req.secure ? 'https' : 'http';
          const baseUrl = `${protocol}://${req.get('host')}`;

          logger.info('OAuth Dynamic Client Registration request (RFC 7591)', {
            method: req.method,
            contentType: req.get('Content-Type'),
            baseUrl,
          });

          // Update provider base URL
          oauthProvider.setBaseUrl(baseUrl);

          const body = req.body || {};

          // RFC 7591 Section 2: Client Registration Request
          // redirect_uris is REQUIRED per RFC 7591
          // However, for MCP compatibility, we provide defaults if not specified
          const registrationRequest = {
            redirect_uris: body.redirect_uris || [
              `${baseUrl}/callback`,
              `http://localhost:${port}/callback`,
              `http://127.0.0.1:${port}/callback`,
            ],
            client_name: body.client_name,
            client_uri: body.client_uri,
            logo_uri: body.logo_uri,
            scope: body.scope,
            contacts: body.contacts,
            tos_uri: body.tos_uri,
            policy_uri: body.policy_uri,
            jwks_uri: body.jwks_uri,
            jwks: body.jwks,
            software_id: body.software_id,
            software_version: body.software_version,
            grant_types: body.grant_types || ['authorization_code', 'refresh_token'],
            response_types: body.response_types || ['code'],
            token_endpoint_auth_method: body.token_endpoint_auth_method,
          };

          // Register the client using the OAuth provider
          const result = oauthProvider.registerClient(registrationRequest);

          // Check for registration errors (RFC 7591 Section 3.2.2)
          if ('error' in result) {
            logger.warn('OAuth client registration failed', {
              error: result.error,
              error_description: result.error_description,
            });

            return res.status(400).json(result);
          }

          // RFC 7591 Section 3.2.1: Client Registration Response
          logger.info('OAuth client registered successfully (RFC 7591)', {
            client_id: result.client_id.substring(0, 16) + '...',
            client_name: result.client_name,
            redirect_uris_count: result.redirect_uris.length,
          });

          res.status(201).json(result);
        } catch (error) {
          logger.error('OAuth client registration error:', error);
          res.status(500).json({
            error: 'server_error',
            error_description: 'Failed to process client registration',
          });
        }
      });

      // Microsoft Graph MCP endpoints with bearer token auth
      // Handle both GET and POST methods as required by MCP Streamable HTTP specification
      // IMPORTANT: OAuth authentication is REQUIRED - no access without valid Bearer token
      app.get(
        '/mcp',
        requireOAuthMiddleware,
        async (
          req: Request & { microsoftAuth?: { accessToken: string; refreshToken: string } },
          res: Response
        ) => {
          // #region agent log
          fetch('http://127.0.0.1:7245/ingest/76c7865f-57f2-4bf0-8001-38b29d141bbc', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              location: 'server.ts:1037',
              message: 'MCP GET request received',
              data: {
                origin: req.headers.origin,
                host: req.get('host'),
                userAgent: req.get('user-agent')?.substring(0, 100),
                hasAuth: !!req.microsoftAuth,
                query: req.query,
                url: req.url,
              },
              timestamp: Date.now(),
              sessionId: 'debug-session',
              runId: 'run1',
              hypothesisId: 'E',
            }),
          }).catch(() => {});
          // #endregion

          // Handle simple verification requests (GET requests without query params are likely verification)
          const hasQueryParams = Object.keys(req.query).length > 0;
          const isVerificationRequest = !hasQueryParams || req.query.verify === 'true';

          if (isVerificationRequest) {
            logger.info(
              'Simple verification GET request detected, returning verification response'
            );
            return res.json({
              status: 'ok',
              service: 'Microsoft 365 MCP Server',
              version: this.version,
              mcp: {
                endpoint: '/mcp',
                protocol: 'streamable-http',
              },
            });
          }

          const handler = async () => {
            const transport = new StreamableHTTPServerTransport({
              sessionIdGenerator: undefined, // Stateless mode
            });

            res.on('close', () => {
              transport.close();
            });

            await this.server!.connect(transport);
            await transport.handleRequest(req as any, res as any, undefined);
          };

          try {
            // Extract chat ID and user ID for memory context
            const chatId = extractChatId(req);
            const userId = extractUserId(req);

            logger.debug('MCP GET request context', { chatId, userId: userId?.substring(0, 8) });

            if (req.microsoftAuth) {
              // SECURITY: Include token hash for secure logging/correlation
              const tokenHash = createTokenHash(req.microsoftAuth.accessToken);
              await requestContext.run(
                {
                  accessToken: req.microsoftAuth.accessToken,
                  refreshToken: req.microsoftAuth.refreshToken,
                  chatId,
                  userId,
                  tokenHash,
                },
                handler
              );
            } else {
              // Even without auth, provide chat context
              await requestContext.run(
                { accessToken: '', chatId, userId, tokenHash: 'no-auth' },
                handler
              );
            }
          } catch (error) {
            logger.error('Error handling MCP GET request:', error);
            // #region agent log
            fetch('http://127.0.0.1:7245/ingest/76c7865f-57f2-4bf0-8001-38b29d141bbc', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                location: 'server.ts:1085',
                message: 'MCP GET request error',
                data: {
                  error: String(error),
                  errorMessage: error instanceof Error ? error.message : 'Unknown error',
                  headersSent: res.headersSent,
                  origin: req.headers.origin,
                  host: req.get('host'),
                },
                timestamp: Date.now(),
                sessionId: 'debug-session',
                runId: 'run1',
                hypothesisId: 'G',
              }),
            }).catch(() => {});
            // #endregion
            if (!res.headersSent) {
              res.status(500).json({
                jsonrpc: '2.0',
                error: {
                  code: -32603,
                  message: 'Internal server error',
                },
                id: null,
              });
            }
          }
        }
      );

      app.post(
        '/mcp',
        requireOAuthMiddleware,
        async (
          req: Request & { microsoftAuth?: { accessToken: string; refreshToken: string } },
          res: Response
        ) => {
          // #region agent log
          fetch('http://127.0.0.1:7245/ingest/76c7865f-57f2-4bf0-8001-38b29d141bbc', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              location: 'server.ts:1101',
              message: 'MCP POST request received',
              data: {
                origin: req.headers.origin,
                host: req.get('host'),
                userAgent: req.get('user-agent')?.substring(0, 100),
                contentType: req.get('Content-Type'),
                hasBody: !!req.body,
                hasAuth: !!req.microsoftAuth,
                method: req.body?.method,
                jsonrpc: req.body?.jsonrpc,
                url: req.url,
              },
              timestamp: Date.now(),
              sessionId: 'debug-session',
              runId: 'run1',
              hypothesisId: 'E',
            }),
          }).catch(() => {});
          // #endregion

          // Log incoming request for debugging
          logger.info('MCP POST request received', {
            contentType: req.get('Content-Type'),
            hasBody: !!req.body,
            bodyType: typeof req.body,
            bodyKeys: req.body && typeof req.body === 'object' ? Object.keys(req.body) : undefined,
            method: req.body?.method,
            jsonrpc: req.body?.jsonrpc,
          });

          // Handle simple verification requests (not proper MCP JSON-RPC)
          const body = req.body;
          const contentType = req.get('Content-Type') || '';
          const userAgent = req.get('User-Agent') || '';

          // Check if this is a verification request from Open WebUI or similar tools
          const isVerificationRequest =
            !body ||
            (typeof body === 'object' && Object.keys(body).length === 0) ||
            (typeof body === 'string' && body.trim() === '') ||
            (contentType.includes('application/x-www-form-urlencoded') && !body?.jsonrpc) ||
            (body?.method === 'initialize' && body?.params === undefined) ||
            userAgent.includes('open-webui') ||
            req.query?.verify === 'true';

          if (
            isVerificationRequest &&
            (!body?.jsonrpc || !body?.method || (body?.method === 'initialize' && !body?.params))
          ) {
            logger.info('Verification request detected, returning verification response', {
              hasBody: !!body,
              method: body?.method,
              userAgent,
            });
            return res.json({
              status: 'ok',
              service: 'Microsoft 365 MCP Server',
              version: this.version,
              mcp: {
                endpoint: '/mcp',
                protocol: 'streamable-http',
              },
            });
          }

          // Log MCP tool calls to QueryStore for dashboard analytics
          const startTime = Date.now();
          const isToolCall = body?.method === 'tools/call';
          const toolName = isToolCall ? body?.params?.name || 'unknown' : null;
          const toolParams = isToolCall ? body?.params?.arguments || {} : null;

          const handler = async () => {
            const transport = new StreamableHTTPServerTransport({
              sessionIdGenerator: undefined, // Stateless mode
            });

            res.on('close', () => {
              transport.close();
            });

            await this.server!.connect(transport);

            // Ensure response headers are set before handling request
            if (!res.headersSent) {
              res.setHeader('Content-Type', 'application/json');
            }

            await transport.handleRequest(req as any, res as any, req.body);
          };

          try {
            // Extract chat ID and user ID for memory context
            const chatId = extractChatId(req);
            const userId = extractUserId(req);

            // Log tool call to QueryStore when dashboard is enabled
            if (isToolCall && isDashboardEnabled()) {
              res.on('finish', () => {
                try {
                  const queryStore = getQueryStore();
                  const durationMs = Date.now() - startTime;
                  const success = res.statusCode >= 200 && res.statusCode < 400;

                  queryStore.storeQuery({
                    userIdHash: queryStore.hashUserId(userId || 'anonymous'),
                    chatId: chatId || undefined,
                    toolName: toolName || 'unknown',
                    parameters: sanitizeToolParams(toolParams || {}),
                    success,
                    errorMessage: success ? undefined : `HTTP ${res.statusCode}`,
                    durationMs,
                    ipAnonymized: req.ip ? queryStore.anonymizeIp(req.ip) : undefined,
                    userAgent: req.get('User-Agent'),
                  });

                  logger.debug('MCP tool call logged to QueryStore', {
                    toolName,
                    success,
                    durationMs,
                  });
                } catch (logError) {
                  logger.error('Failed to log MCP tool call:', logError);
                }
              });
            }

            logger.debug('MCP POST request context', { chatId, userId: userId?.substring(0, 8) });

            if (req.microsoftAuth) {
              // SECURITY: Include token hash for secure logging/correlation
              const tokenHash = createTokenHash(req.microsoftAuth.accessToken);
              await requestContext.run(
                {
                  accessToken: req.microsoftAuth.accessToken,
                  refreshToken: req.microsoftAuth.refreshToken,
                  chatId,
                  userId,
                  tokenHash,
                },
                handler
              );
            } else {
              // Even without auth, provide chat context
              await requestContext.run(
                { accessToken: '', chatId, userId, tokenHash: 'no-auth' },
                handler
              );
            }
          } catch (error) {
            logger.error('Error handling MCP POST request:', error);
            // #region agent log
            fetch('http://127.0.0.1:7245/ingest/76c7865f-57f2-4bf0-8001-38b29d141bbc', {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({
                location: 'server.ts:1251',
                message: 'MCP POST request error',
                data: {
                  error: String(error),
                  errorMessage: error instanceof Error ? error.message : 'Unknown error',
                  headersSent: res.headersSent,
                  origin: req.headers.origin,
                  host: req.get('host'),
                  hasAuth: !!req.microsoftAuth,
                },
                timestamp: Date.now(),
                sessionId: 'debug-session',
                runId: 'run1',
                hypothesisId: 'G',
              }),
            }).catch(() => {});
            // #endregion
            if (!res.headersSent) {
              res.status(500).json({
                jsonrpc: '2.0',
                error: {
                  code: -32603,
                  message: 'Internal server error',
                },
                id: null,
              });
            }
          }
        }
      );

      // Query Dashboard (password protected)
      if (isDashboardEnabled()) {
        logger.info('Query Dashboard enabled - access at /dashboard');
        app.use('/dashboard', createDashboardRouter());
      } else {
        logger.info(
          'Query Dashboard disabled - set DASHBOARD_PASSWORD environment variable to enable'
        );
        // Explicitly handle /dashboard routes when disabled
        // Returns 503 Service Unavailable to indicate the feature is not configured
        app.use('/dashboard', (_req, res) => {
          res.status(503).json({
            error: 'Dashboard not available',
            message:
              'The Query Dashboard is disabled. Set DASHBOARD_PASSWORD environment variable to enable it.',
            hint: 'Contact your administrator to configure the dashboard.',
          });
        });
      }

      // Health check endpoint
      app.get('/', (req, res) => {
        res.json({
          status: 'ok',
          service: 'Microsoft 365 MCP Server',
          version: this.version,
          endpoints: {
            mcp: '/mcp',
            authorize: '/authorize',
            token: '/token',
            oauthDiscovery: '/.well-known/oauth-authorization-server',
            protectedResourceDiscovery: '/.well-known/oauth-protected-resource',
            ...(isDashboardEnabled() ? { dashboard: '/dashboard' } : {}),
          },
        });
      });

      // POST handler for root endpoint (for health checks or other purposes)
      app.post('/', (req, res) => {
        res.json({
          status: 'ok',
          service: 'Microsoft 365 MCP Server',
          message: 'POST requests to root endpoint are supported',
          version: this.version,
        });
      });

      // Favicon handler - return 204 No Content to prevent 404 errors
      app.get('/favicon.ico', (req, res) => {
        res.status(204).end();
      });

      // Verification endpoint for tool server verification (Open WebUI compatibility)
      app.get('/verify', async (req, res) => {
        try {
          res.json({
            status: 'ok',
            service: 'Microsoft 365 MCP Server',
            version: this.version,
            mcp: {
              endpoint: '/mcp',
              protocol: 'streamable-http',
            },
          });
        } catch (error) {
          logger.error('Error in verification endpoint:', error);
          res.status(500).json({
            status: 'error',
            message: 'Verification failed',
          });
        }
      });

      // POST verification endpoint
      app.post('/verify', async (req, res) => {
        try {
          res.json({
            status: 'ok',
            service: 'Microsoft 365 MCP Server',
            version: this.version,
            mcp: {
              endpoint: '/mcp',
              protocol: 'streamable-http',
            },
          });
        } catch (error) {
          logger.error('Error in verification endpoint:', error);
          res.status(500).json({
            status: 'error',
            message: 'Verification failed',
          });
        }
      });

      if (host) {
        const isAllInterfaces = host === '0.0.0.0' || host === '::';
        app.listen(port, host, () => {
          if (isAllInterfaces) {
            logger.info(`Server listening on all interfaces (${host}:${port})`);
            logger.info(
              `  - MCP endpoint: http://0.0.0.0:${port}/mcp (accessible from any interface)`
            );
            logger.info(
              `  - OAuth endpoints: http://0.0.0.0:${port}/auth/* (accessible from any interface)`
            );
            logger.info(
              `  - OAuth discovery: http://0.0.0.0:${port}/.well-known/oauth-authorization-server (accessible from any interface)`
            );
          } else {
            logger.info(`Server listening on ${host}:${port}`);
            logger.info(`  - MCP endpoint: http://${host}:${port}/mcp`);
            logger.info(`  - OAuth endpoints: http://${host}:${port}/auth/*`);
            logger.info(
              `  - OAuth discovery: http://${host}:${port}/.well-known/oauth-authorization-server`
            );
          }
        });
      } else {
        app.listen(port, () => {
          logger.info(`Server listening on all interfaces (0.0.0.0:${port})`);
          logger.info(
            `  - MCP endpoint: http://0.0.0.0:${port}/mcp (accessible from any interface)`
          );
          logger.info(
            `  - OAuth endpoints: http://0.0.0.0:${port}/auth/* (accessible from any interface)`
          );
          logger.info(
            `  - OAuth discovery: http://0.0.0.0:${port}/.well-known/oauth-authorization-server (accessible from any interface)`
          );
        });
      }
    } else {
      const transport = new StdioServerTransport();
      await this.server!.connect(transport);
      logger.info('Server connected to stdio transport');
    }
  }
}

export default MicrosoftGraphServer;
