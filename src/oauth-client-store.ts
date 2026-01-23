/**
 * OAuth 2.0 Dynamic Client Registration Store
 * Implements RFC 7591 - OAuth 2.0 Dynamic Client Registration Protocol
 *
 * @see https://tools.ietf.org/html/rfc7591
 * @see https://modelcontextprotocol.io
 */

import { randomUUID } from 'crypto';
import logger from './logger.js';

/**
 * Client registration request as per RFC 7591 Section 2
 */
export interface ClientRegistrationRequest {
  /** Array of redirect URIs for the client */
  redirect_uris: string[];
  /** Human-readable name of the client */
  client_name?: string;
  /** URL of the client's home page */
  client_uri?: string;
  /** URL of the client's logo */
  logo_uri?: string;
  /** Space-separated list of scopes the client can request */
  scope?: string;
  /** Array of contact email addresses */
  contacts?: string[];
  /** URL to the client's terms of service */
  tos_uri?: string;
  /** URL to the client's privacy policy */
  policy_uri?: string;
  /** URL to the client's JWK Set document */
  jwks_uri?: string;
  /** JWK Set document embedded in the registration */
  jwks?: object;
  /** Unique identifier for the client software */
  software_id?: string;
  /** Version of the client software */
  software_version?: string;
  /** Array of grant types the client will use */
  grant_types?: string[];
  /** Array of response types the client will use */
  response_types?: string[];
  /** Token endpoint authentication method */
  token_endpoint_auth_method?: 'none' | 'client_secret_post' | 'client_secret_basic';
}

/**
 * Registered client as stored in the client store
 */
export interface RegisteredClient {
  /** Unique client identifier */
  client_id: string;
  /** Client secret for confidential clients */
  client_secret?: string;
  /** Timestamp when the client_id was issued (Unix timestamp) */
  client_id_issued_at: number;
  /** Timestamp when the client_secret expires (0 = never, Unix timestamp) */
  client_secret_expires_at?: number;
  /** Array of redirect URIs */
  redirect_uris: string[];
  /** Human-readable name of the client */
  client_name?: string;
  /** URL of the client's home page */
  client_uri?: string;
  /** URL of the client's logo */
  logo_uri?: string;
  /** Space-separated list of scopes */
  scope?: string;
  /** Array of contact email addresses */
  contacts?: string[];
  /** URL to the client's terms of service */
  tos_uri?: string;
  /** URL to the client's privacy policy */
  policy_uri?: string;
  /** URL to the client's JWK Set document */
  jwks_uri?: string;
  /** JWK Set document */
  jwks?: object;
  /** Unique identifier for the client software */
  software_id?: string;
  /** Version of the client software */
  software_version?: string;
  /** Array of grant types */
  grant_types: string[];
  /** Array of response types */
  response_types: string[];
  /** Token endpoint authentication method */
  token_endpoint_auth_method: string;
  /** Registration access token for client management */
  registration_access_token?: string;
  /** Client management URI */
  registration_client_uri?: string;
}

/**
 * Client registration response as per RFC 7591 Section 3.2.1
 */
export interface ClientRegistrationResponse extends RegisteredClient {
  /** Extended metadata from the server */
  issuer?: string;
  /** Authorization endpoint URL */
  authorization_endpoint?: string;
  /** Token endpoint URL */
  token_endpoint?: string;
  /** Scopes supported by this registration */
  scopes?: string[];
}

/**
 * Client registration error as per RFC 7591 Section 3.2.2
 */
export interface ClientRegistrationError {
  error:
    | 'invalid_redirect_uri'
    | 'invalid_client_metadata'
    | 'invalid_software_statement'
    | 'unapproved_software_statement';
  error_description?: string;
}

/**
 * In-memory OAuth Client Store
 * Stores dynamically registered clients for RFC 7591 compliance
 */
export class OAuthClientStore {
  private clients: Map<string, RegisteredClient> = new Map();
  private azureClientId: string;
  private azureClientSecret?: string;

  constructor(azureClientId: string, azureClientSecret?: string) {
    this.azureClientId = azureClientId;
    this.azureClientSecret = azureClientSecret;
  }

  /**
   * Validate redirect URIs according to RFC 7591
   */
  private validateRedirectUris(redirectUris: string[]): { valid: boolean; error?: string } {
    if (!redirectUris || redirectUris.length === 0) {
      return { valid: false, error: 'redirect_uris is required and must not be empty' };
    }

    for (const uri of redirectUris) {
      try {
        const parsed = new URL(uri);

        // Allow localhost and 127.0.0.1 for development
        const isLocalhost = parsed.hostname === 'localhost' || parsed.hostname === '127.0.0.1';

        // In production, require HTTPS unless localhost
        if (!isLocalhost && parsed.protocol !== 'https:') {
          // For development, allow HTTP with any host
          if (process.env.NODE_ENV === 'production') {
            return { valid: false, error: `redirect_uri must use HTTPS: ${uri}` };
          }
        }

        // Disallow fragment identifiers
        if (parsed.hash) {
          return { valid: false, error: `redirect_uri must not contain a fragment: ${uri}` };
        }
      } catch {
        return { valid: false, error: `Invalid redirect_uri: ${uri}` };
      }
    }

    return { valid: true };
  }

  /**
   * Generate a unique client ID
   */
  private generateClientId(): string {
    return `mcp-client-${randomUUID()}`;
  }

  /**
   * Generate a client secret for confidential clients
   */
  private generateClientSecret(): string {
    const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_';
    let secret = '';
    for (let i = 0; i < 48; i++) {
      secret += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    return secret;
  }

  /**
   * Register a new OAuth client (RFC 7591)
   */
  registerClient(
    request: ClientRegistrationRequest,
    options?: {
      baseUrl?: string;
      isConfidential?: boolean;
      scopes?: string[];
    }
  ): ClientRegistrationResponse | ClientRegistrationError {
    // Validate redirect URIs
    const validation = this.validateRedirectUris(request.redirect_uris);
    if (!validation.valid) {
      logger.warn('Client registration failed: invalid redirect_uri', { error: validation.error });
      return {
        error: 'invalid_redirect_uri',
        error_description: validation.error,
      };
    }

    // Generate client credentials
    const clientId = this.generateClientId();
    const now = Math.floor(Date.now() / 1000);

    // Determine if client needs a secret
    const isConfidential = options?.isConfidential ?? false;
    const clientSecret = isConfidential ? this.generateClientSecret() : undefined;

    // Default grant types and response types
    const grantTypes = request.grant_types || ['authorization_code'];
    const responseTypes = request.response_types || ['code'];
    const tokenEndpointAuthMethod =
      request.token_endpoint_auth_method || (isConfidential ? 'client_secret_post' : 'none');

    // Create registered client
    const registeredClient: RegisteredClient = {
      client_id: clientId,
      client_secret: clientSecret,
      client_id_issued_at: now,
      client_secret_expires_at: clientSecret ? 0 : undefined, // 0 = never expires
      redirect_uris: request.redirect_uris,
      client_name: request.client_name,
      client_uri: request.client_uri,
      logo_uri: request.logo_uri,
      scope: request.scope,
      contacts: request.contacts,
      tos_uri: request.tos_uri,
      policy_uri: request.policy_uri,
      jwks_uri: request.jwks_uri,
      jwks: request.jwks,
      software_id: request.software_id,
      software_version: request.software_version,
      grant_types: grantTypes,
      response_types: responseTypes,
      token_endpoint_auth_method: tokenEndpointAuthMethod,
    };

    // Store the client
    this.clients.set(clientId, registeredClient);

    logger.info('OAuth client registered successfully', {
      client_id: clientId,
      client_name: request.client_name,
      redirect_uris: request.redirect_uris.length,
      grant_types: grantTypes,
    });

    // Build response with additional server metadata
    const response: ClientRegistrationResponse = {
      ...registeredClient,
    };

    if (options?.baseUrl) {
      response.issuer = options.baseUrl;
      response.authorization_endpoint = `${options.baseUrl}/authorize`;
      response.token_endpoint = `${options.baseUrl}/token`;
    }

    if (options?.scopes) {
      response.scopes = options.scopes;
    }

    return response;
  }

  /**
   * Get a registered client by ID
   */
  getClient(clientId: string): RegisteredClient | null {
    // First check dynamically registered clients
    const dynamicClient = this.clients.get(clientId);
    if (dynamicClient) {
      return dynamicClient;
    }

    // Fall back to the Azure AD client (for backward compatibility)
    if (clientId === this.azureClientId) {
      return {
        client_id: this.azureClientId,
        client_secret: this.azureClientSecret,
        client_id_issued_at: 0,
        redirect_uris: ['http://localhost:3000/callback'],
        grant_types: ['authorization_code', 'refresh_token'],
        response_types: ['code'],
        token_endpoint_auth_method: this.azureClientSecret ? 'client_secret_post' : 'none',
      };
    }

    return null;
  }

  /**
   * Check if a redirect URI is valid for a client
   */
  isValidRedirectUri(clientId: string, redirectUri: string): boolean {
    const client = this.getClient(clientId);
    if (!client) {
      return false;
    }

    return client.redirect_uris.some((uri) => {
      // Exact match
      if (uri === redirectUri) return true;

      // Allow matching with different ports for localhost development
      try {
        const registered = new URL(uri);
        const requested = new URL(redirectUri);

        if (
          registered.hostname === 'localhost' &&
          requested.hostname === 'localhost' &&
          registered.pathname === requested.pathname
        ) {
          return true;
        }

        if (
          registered.hostname === '127.0.0.1' &&
          requested.hostname === '127.0.0.1' &&
          registered.pathname === requested.pathname
        ) {
          return true;
        }
      } catch {
        return false;
      }

      return false;
    });
  }

  /**
   * Delete a registered client
   */
  deleteClient(clientId: string): boolean {
    if (clientId === this.azureClientId) {
      // Cannot delete the primary Azure AD client
      return false;
    }

    const deleted = this.clients.delete(clientId);
    if (deleted) {
      logger.info('OAuth client deleted', { client_id: clientId });
    }
    return deleted;
  }

  /**
   * Get all registered clients (for debugging)
   */
  getAllClients(): RegisteredClient[] {
    return Array.from(this.clients.values());
  }

  /**
   * Get the primary Azure AD client ID
   */
  getPrimaryClientId(): string {
    return this.azureClientId;
  }

  /**
   * Clear all dynamically registered clients
   */
  clearClients(): void {
    this.clients.clear();
    logger.info('All dynamically registered OAuth clients cleared');
  }
}

// Singleton instance
let clientStoreInstance: OAuthClientStore | null = null;

/**
 * Initialize the OAuth client store
 */
export function initializeClientStore(
  azureClientId: string,
  azureClientSecret?: string
): OAuthClientStore {
  clientStoreInstance = new OAuthClientStore(azureClientId, azureClientSecret);
  return clientStoreInstance;
}

/**
 * Get the OAuth client store instance
 */
export function getClientStore(): OAuthClientStore | null {
  return clientStoreInstance;
}

export default OAuthClientStore;
