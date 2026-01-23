/**
 * Microsoft OAuth Provider for MCP Server
 * Implements OAuth 2.1 and RFC 7591 Dynamic Client Registration
 *
 * MCP servers function as Resource Servers (not Authorization Servers),
 * validating tokens and providing protected resources.
 *
 * @see https://modelcontextprotocol.io
 * @see https://tools.ietf.org/html/rfc7591
 */

import { ProxyOAuthServerProvider } from '@modelcontextprotocol/sdk/server/auth/providers/proxyProvider.js';
import type { AuthInfo } from '@modelcontextprotocol/sdk/server/auth/types.js';
import logger from './logger.js';
import AuthManager from './auth.js';
import type { AppSecrets } from './secrets.js';
import { getCloudEndpoints } from './cloud-config.js';
import {
  OAuthClientStore,
  initializeClientStore,
  type ClientRegistrationRequest,
  type ClientRegistrationResponse,
  type ClientRegistrationError,
  type RegisteredClient,
} from './oauth-client-store.js';

/**
 * Configuration options for the Microsoft OAuth Provider
 */
export interface MicrosoftOAuthProviderOptions {
  /** Base URL of the MCP server (for constructing redirect URIs) */
  baseUrl?: string;
  /** Port the server is running on */
  port?: number;
  /** Available scopes for registered clients */
  scopes?: string[];
}

/**
 * Microsoft OAuth Provider extending ProxyOAuthServerProvider
 *
 * This provider:
 * - Validates access tokens against Microsoft Graph API
 * - Supports dynamic client registration (RFC 7591)
 * - Manages registered OAuth clients
 */
export class MicrosoftOAuthProvider extends ProxyOAuthServerProvider {
  private authManager: AuthManager;
  private clientStore: OAuthClientStore;
  private options: MicrosoftOAuthProviderOptions;
  private secrets: AppSecrets;

  constructor(
    authManager: AuthManager,
    secrets: AppSecrets,
    options: MicrosoftOAuthProviderOptions = {}
  ) {
    const tenantId = secrets.tenantId || 'common';
    const clientId = secrets.clientId;
    const cloudEndpoints = getCloudEndpoints(secrets.cloudType);

    // Initialize the client store
    const clientStore = initializeClientStore(clientId, secrets.clientSecret);

    super({
      endpoints: {
        authorizationUrl: `${cloudEndpoints.authority}/${tenantId}/oauth2/v2.0/authorize`,
        tokenUrl: `${cloudEndpoints.authority}/${tenantId}/oauth2/v2.0/token`,
        revocationUrl: `${cloudEndpoints.authority}/${tenantId}/oauth2/v2.0/logout`,
      },
      /**
       * Verify access token by calling Microsoft Graph API
       * MCP servers act as Resource Servers, validating tokens
       */
      verifyAccessToken: async (token: string): Promise<AuthInfo> => {
        try {
          const response = await fetch(`${cloudEndpoints.graphApi}/v1.0/me`, {
            headers: {
              Authorization: `Bearer ${token}`,
            },
          });

          if (response.ok) {
            const userData = await response.json();
            logger.info(`OAuth token verified for user: ${userData.userPrincipalName}`);

            // Store the token in the auth manager for subsequent API calls
            await authManager.setOAuthToken(token);

            return {
              token,
              clientId,
              scopes: [],
            };
          } else {
            const errorText = await response.text();
            logger.error(`Token verification failed: ${response.status}`, { error: errorText });
            throw new Error(`Token verification failed: ${response.status}`);
          }
        } catch (error) {
          logger.error(`OAuth token verification error: ${error}`);
          throw error;
        }
      },
      /**
       * Get client by ID from the client store
       * Supports both dynamically registered clients and the primary Azure AD client
       */
      getClient: async (requestedClientId: string) => {
        const client = clientStore.getClient(requestedClientId);
        if (client) {
          return {
            client_id: client.client_id,
            client_secret: client.client_secret,
            redirect_uris: client.redirect_uris,
          };
        }

        // Return null if client not found (MCP SDK will handle error)
        logger.warn(`Client not found: ${requestedClientId}`);
        return null;
      },
    });

    this.authManager = authManager;
    this.clientStore = clientStore;
    this.options = options;
    this.secrets = secrets;

    logger.info('MicrosoftOAuthProvider initialized', {
      tenantId,
      clientId: clientId.substring(0, 8) + '...',
      hasClientSecret: !!secrets.clientSecret,
    });
  }

  /**
   * Register a new OAuth client (RFC 7591)
   *
   * @param request - Client registration request
   * @returns Registration response or error
   */
  registerClient(
    request: ClientRegistrationRequest
  ): ClientRegistrationResponse | ClientRegistrationError {
    return this.clientStore.registerClient(request, {
      baseUrl: this.options.baseUrl,
      isConfidential: false, // MCP clients are typically public clients
      scopes: this.options.scopes,
    });
  }

  /**
   * Get a registered client by ID
   */
  getRegisteredClient(clientId: string): RegisteredClient | null {
    return this.clientStore.getClient(clientId);
  }

  /**
   * Check if a redirect URI is valid for a client
   */
  isValidRedirectUri(clientId: string, redirectUri: string): boolean {
    return this.clientStore.isValidRedirectUri(clientId, redirectUri);
  }

  /**
   * Get the primary Azure AD client ID
   */
  getPrimaryClientId(): string {
    return this.clientStore.getPrimaryClientId();
  }

  /**
   * Get the client store instance
   */
  getClientStore(): OAuthClientStore {
    return this.clientStore;
  }

  /**
   * Update the base URL (useful when server starts on dynamic port)
   */
  setBaseUrl(baseUrl: string): void {
    this.options.baseUrl = baseUrl;
  }

  /**
   * Update available scopes
   */
  setScopes(scopes: string[]): void {
    this.options.scopes = scopes;
  }

  /**
   * Get OAuth authorization server metadata (RFC 8414)
   */
  getAuthorizationServerMetadata(baseUrl: string): Record<string, unknown> {
    return {
      issuer: baseUrl,
      authorization_endpoint: `${baseUrl}/authorize`,
      token_endpoint: `${baseUrl}/token`,
      registration_endpoint: `${baseUrl}/register`,
      revocation_endpoint: `${baseUrl}/revoke`,
      response_types_supported: ['code'],
      response_modes_supported: ['query'],
      grant_types_supported: ['authorization_code', 'refresh_token'],
      token_endpoint_auth_methods_supported: ['none', 'client_secret_post'],
      code_challenge_methods_supported: ['S256'],
      scopes_supported: this.options.scopes || [],
    };
  }

  /**
   * Get protected resource metadata (RFC 9728)
   */
  getProtectedResourceMetadata(baseUrl: string): Record<string, unknown> {
    return {
      resource: `${baseUrl}/mcp`,
      authorization_servers: [baseUrl],
      scopes_supported: this.options.scopes || [],
      bearer_methods_supported: ['header'],
      resource_documentation: baseUrl,
    };
  }
}

export default MicrosoftOAuthProvider;
